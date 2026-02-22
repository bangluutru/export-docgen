/**
 * PDF Renderer — Breakthrough approach: Excel → HTML Table → html2canvas → PDF
 * 
 * Strategy: Build a complete HTML table from the template XML with ALL merge cells,
 * styles, column widths, and row heights. Render it VISIBLE in the viewport
 * (behind loading overlay) so html2canvas can capture it. Then place the
 * captured canvas image into jsPDF pages.
 *
 * This leverages the browser's native text rendering engine for perfect layout.
 */
const SVGPDFRenderer = (() => {
    'use strict';

    const PAGE_SIZES = {
        a4: { w: 210, h: 297 },
        a3: { w: 297, h: 420 },
        letter: { w: 215.9, h: 279.4 },
    };

    const THEME_COLORS = [
        'FFFFFF', '000000', 'E7E6E6', '44546A',
        '4472C4', 'ED7D31', 'A5A5A5', 'FFC000',
        '5B9BD5', '70AD47',
    ];

    // ===== Main entry =====

    async function renderToPDF(opts) {
        const {
            headers, rows,
            templateData = null,
            title = '',
            pageSize = 'a4',
            landscape = false,
            margins = { top: 8, right: 6, bottom: 8, left: 6 },
        } = opts;

        if (templateData && templateData.zip) {
            return await renderFromTemplate(templateData, rows, pageSize, landscape, margins);
        }
        return await renderSimpleTable(headers, rows, title, pageSize, landscape, margins);
    }

    // ===== Template-based rendering =====

    async function renderFromTemplate(templateData, dataRows, pageSize, landscape, margins) {
        const zip = templateData.zip;
        const analysis = templateData.analysis;

        // 1. Parse template XML
        const sheetXml = await zip.file(templateData.sheetPaths[0]).async('string');
        const sheetDoc = new DOMParser().parseFromString(sheetXml, 'application/xml');

        // 2. Shared strings
        const ssFile = zip.file('xl/sharedStrings.xml');
        const ssXml = ssFile ? await ssFile.async('string') : '<sst/>';
        const strings = parseStrings(ssXml);

        // 3. Styles
        const stylesXml = await zip.file('xl/styles.xml').async('string');
        const styles = parseStyles(stylesXml);

        // 4. Column widths
        const colWidths = parseColumnWidths(sheetDoc);

        // 5. Merge cells from template
        const merges = parseMergeCells(sheetDoc);

        // 6. All template rows
        const allRows = parseAllRows(sheetDoc, strings);

        // 7. Zone detection
        const dataStart = analysis.dataZone.startRowNum;
        const dataEnd = analysis.dataZone.endRowNum;
        const maxCol = analysis.maxCol || 8;

        // 8. Build final rows: header zone + new data + shifted footer
        const headerRows = allRows.filter(r => r.rowNum < dataStart);
        const tplDataRows = allRows.filter(r => r.rowNum >= dataStart && r.rowNum <= dataEnd);
        const footerRows = allRows.filter(r => r.rowNum > dataEnd);

        // FIX 1: Use ONLY the first data row style pattern for consistent font size.
        // Template may have 2+ data rows with DIFFERENT font sizes — cycling them causes
        // alternating large/small text. Using only pattern[0] ensures uniformity.
        const firstPattern = { cellStyles: {}, ht: 18 };
        if (tplDataRows.length > 0) {
            firstPattern.ht = tplDataRows[0].ht || 18;
            for (const cell of tplDataRows[0].cells) {
                firstPattern.cellStyles[cell.colNum] = cell.s;
            }
        }

        // New data rows — all use the SAME style pattern
        const newDataRows = [];
        for (let ri = 0; ri < dataRows.length; ri++) {
            const rowData = dataRows[ri];
            const rowNum = dataStart + ri;

            const cells = [];
            for (let ci = 0; ci < Math.min(rowData.length, maxCol); ci++) {
                cells.push({
                    colNum: ci + 1,
                    s: firstPattern.cellStyles[ci + 1] || 0,
                    display: String(rowData[ci] ?? ''),
                    t: '',
                });
            }
            newDataRows.push({ rowNum, ht: firstPattern.ht, cells });
        }

        // Shift footer
        const shift = newDataRows.length - (dataEnd - dataStart + 1);
        const shiftedFooter = footerRows.map(r => ({
            ...r,
            rowNum: r.rowNum + shift,
            cells: r.cells.map(c => {
                const copy = { ...c };
                // Clear cached formula values (numeric cells) in footer.
                // These are stale results from the template's original data.
                // Labels like 小計/税金 are shared strings (t='s') and are preserved.
                if (copy.t !== 's' && copy.display && !isNaN(parseFloat(copy.display))) {
                    copy.display = '';
                }
                return copy;
            }),
        }));

        // Adjust merges
        const adjustedMerges = [];
        for (const m of merges) {
            if (m.startRow >= dataStart && m.endRow <= dataEnd) continue;
            if (m.startRow > dataEnd) {
                adjustedMerges.push({
                    ...m,
                    startRow: m.startRow + shift,
                    endRow: m.endRow + shift,
                });
            } else {
                adjustedMerges.push(m);
            }
        }

        // Build merge map
        const mergeMap = buildMergeMap(adjustedMerges);

        // DEFINITIVE FIX: Collapse consecutive empty rows.
        // Previous approaches (merge-aware filter, trailing trim) failed because
        // empty rows are INSIDE merge cells whose origin has content (e.g. 備考 spans 10 rows).
        // New approach: simply walk through ALL rows and collapse any run of 2+ consecutive
        // empty rows into a single spacer. This is universal and handles all cases.
        const allRowsUnsorted = [...headerRows, ...newDataRows, ...shiftedFooter];
        allRowsUnsorted.sort((a, b) => a.rowNum - b.rowNum);

        const hasContent = (row) => row.cells.some(c => c.display && String(c.display).trim().length > 0);

        const allFinalRows = [];
        let consecutiveEmptyCount = 0;
        for (const row of allRowsUnsorted) {
            if (hasContent(row)) {
                consecutiveEmptyCount = 0;
                allFinalRows.push(row);
            } else {
                consecutiveEmptyCount++;
                if (consecutiveEmptyCount <= 1) {
                    allFinalRows.push(row); // Keep max 1 spacer row
                }
                // Skip all additional consecutive empty rows
            }
        }

        // 9. Calculate page dimensions (in pixels for HTML)
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const pw = landscape ? pgBase.h : pgBase.w; // mm
        const contentWidthMM = pw - margins.left - margins.right;
        const PX_PER_MM = 3.78; // 96 dpi
        const contentWidthPX = Math.round(contentWidthMM * PX_PER_MM);

        // Use 100% Excel template column widths — the template designer set these proportions.
        // Header text uses overflow:visible so it doesn't need wider columns.
        let totalExcelW = 0;
        for (let c = 1; c <= maxCol; c++) totalExcelW += (colWidths[c] || 8.43);
        const colPxWidths = {};
        for (let c = 1; c <= maxCol; c++) {
            colPxWidths[c] = Math.round(((colWidths[c] || 8.43) / totalExcelW) * contentWidthPX);
        }

        // FIX 5: Extract seal/stamp images from template with position info
        let sealImages = [];
        try {
            sealImages = await extractTemplateImages(zip, dataStart, dataEnd, shift);
        } catch (e) { /* ignore if no images */ }

        // 10. Build HTML table
        const html = buildFullHTML(allFinalRows, mergeMap, styles, maxCol, colPxWidths, contentWidthPX, dataStart);

        // 11. Render HTML → Canvas → PDF (with row-aware page breaks)
        return await htmlCanvasToPDF(html, pageSize, landscape, margins, contentWidthPX, sealImages);
    }

    // ===== Build HTML table =====

    function buildFullHTML(rows, mergeMap, styles, maxCol, colPxWidths, totalWidth, dataStartRowNum) {
        let html = `<table style="border-collapse:collapse; width:${totalWidth}px; table-layout:fixed; font-family:'MS PGothic','Yu Gothic','Meiryo',sans-serif; font-size:10px;">`;

        // Colgroup for fixed widths
        html += '<colgroup>';
        for (let c = 1; c <= maxCol; c++) {
            html += `<col style="width:${colPxWidths[c]}px">`;
        }
        html += '</colgroup>';

        for (const row of rows) {
            const cellMap = {};
            for (const cell of row.cells) cellMap[cell.colNum] = cell;

            const rowH = Math.max(Math.round((row.ht || 16) * 1.15), 14);
            html += `<tr style="height:${rowH}px">`;

            let skipUntilCol = 0;

            for (let c = 1; c <= maxCol; c++) {
                if (c <= skipUntilCol) continue;

                const key = `${row.rowNum},${c}`;
                const merge = mergeMap[key];

                if (merge && !merge.isOrigin) continue; // covered by explicit merge

                let colspanAttr = 1;
                let rowspanAttr = 1;
                let cellW = colPxWidths[c];

                if (merge && merge.isOrigin) {
                    colspanAttr = merge.colspan || 1;
                    rowspanAttr = merge.rowspan || 1;
                    cellW = 0;
                    for (let mc = merge.startCol; mc <= merge.endCol; mc++) {
                        cellW += (colPxWidths[mc] || 40);
                    }
                } else {
                    // Breakthrough Auto-Colspan: If this cell has text, and adjacent cells are empty,
                    // absorb them so the text has room to breathe, mirroring Excel's visual overflow.
                    const cell = cellMap[c];
                    if (cell && cell.display && String(cell.display).trim().length > 0) {
                        for (let nc = c + 1; nc <= maxCol; nc++) {
                            const nextKey = `${row.rowNum},${nc}`;
                            const nextMerge = mergeMap[nextKey];
                            const nextCell = cellMap[nc];

                            // Stop if hitting an explicit merge
                            if (nextMerge) break;

                            // Stop if next cell has its own text
                            if (nextCell && nextCell.display && String(nextCell.display).trim().length > 0) break;

                            // Stop if next cell has a distinct background color (don't erase visual blocks)
                            let hasDistinctFill = false;
                            if (nextCell && nextCell.s) {
                                const nxf = styles.xfs[nextCell.s] || {};
                                const nfill = styles.fills[nxf.fillId] || {};
                                if (nfill.pattern && nfill.pattern !== 'none' && nfill.pattern !== 'gray125') {
                                    hasDistinctFill = true;
                                }
                            }
                            if (hasDistinctFill) break;

                            // Safe to absorb this empty column!
                            colspanAttr++;
                        }
                    }
                }

                if (colspanAttr > 1 && !(merge && merge.isOrigin)) {
                    skipUntilCol = c + colspanAttr - 1;
                }

                const cell = cellMap[c];
                // Header zone rows: use overflow:visible so long text spills into adjacent empty cells,
                // exactly like Excel's visual overflow behavior. Data zone keeps overflow:hidden.
                const isHeaderZone = dataStartRowNum && row.rowNum < dataStartRowNum;
                const css = buildCellCSS(cell, styles, isHeaderZone);
                let content = '';

                if (cell && cell.display) {
                    // Format numbers
                    const xf = styles.xfs[cell.s] || {};
                    const numFmtId = xf.numFmtId || 0;
                    if (numFmtId > 0 && cell.t !== 's' && !isNaN(parseFloat(cell.display))) {
                        content = formatNumber(cell.display, numFmtId, styles.numFmts);
                    } else {
                        content = cell.display;
                    }
                }

                const colspanStr = colspanAttr > 1 ? ` colspan="${colspanAttr}"` : '';
                const rowspanStr = rowspanAttr > 1 ? ` rowspan="${rowspanAttr}"` : '';
                html += `<td${colspanStr}${rowspanStr} style="${css}">${escapeHtml(content)}</td>`;
            }
            html += '</tr>';
        }

        html += '</table>';
        return html;
    }

    function buildCellCSS(cell, styles, isHeaderZone) {
        // Header zone: overflow:visible so text flows into adjacent empty cells (like Excel)
        // Data zone: overflow:hidden to keep rows clean
        const overflow = isHeaderZone ? 'overflow:visible; white-space:nowrap;' : 'overflow:hidden; white-space:nowrap;';
        let css = `padding:1px 3px; ${overflow} vertical-align:middle;`;
        if (!cell) return css;

        const xf = styles.xfs[cell.s] || {};
        const font = styles.fonts[xf.fontId] || {};
        const fill = styles.fills[xf.fillId] || {};
        const border = isHeaderZone ? {} : (styles.borders[xf.borderId] || {}); // FIX 2: No borders in header zone
        const align = xf.alignment || {};

        // Font
        const fontSize = Math.min(font.size || 10, 14);
        css += `font-size:${fontSize}px;`;
        if (font.bold) css += 'font-weight:bold;';

        // Font color: default to BLACK. Only override with explicit RGB (not theme+tint).
        // Theme colors with tint produce gray text — we want solid black for readability.
        css += 'color:#000000;';
        if (font.colorRgb) {
            const rgb = font.colorRgb.length === 8 ? font.colorRgb.substring(2) : font.colorRgb;
            // Only apply non-white, non-near-black colors (actual colored text)
            if (rgb !== '000000' && rgb !== 'FFFFFF' && rgb !== 'FF000000') {
                css += `color:#${rgb};`;
            }
        }

        // Fill
        const bg = resolveFillColor(fill);
        if (bg) css += `background-color:#${bg};`;

        // Alignment
        if (align.horizontal === 'center') css += 'text-align:center;';
        else if (align.horizontal === 'right') css += 'text-align:right;';
        else if (cell.display && !isNaN(parseFloat(cell.display)) && cell.t !== 's') css += 'text-align:right;';

        if (align.vertical === 'center') css += 'vertical-align:middle;';
        else if (align.vertical === 'top') css += 'vertical-align:top;';

        // Borders — only draw if explicitly defined in template
        css += borderSideCSS('top', border.top);
        css += borderSideCSS('bottom', border.bottom);
        css += borderSideCSS('left', border.left);
        css += borderSideCSS('right', border.right);

        return css;
    }

    function borderSideCSS(side, b) {
        if (!b || !b.style) return '';
        const widths = { thin: '0.5px', medium: '1.5px', thick: '2px', hair: '0.5px' };
        const w = widths[b.style] || '0.5px';
        // Use subtle gray borders for aesthetics
        let color = '#aaa';
        if (b.style === 'medium' || b.style === 'thick') color = '#666';
        if (b.color) {
            const resolved = b.color.length === 8 ? b.color.substring(2) : b.color;
            // Only use explicit color if it's not auto-black
            if (resolved !== '000000' && resolved !== 'FF000000') {
                color = '#' + resolved;
            }
        }
        return `border-${side}:${w} solid ${color};`;
    }

    // ===== HTML → Canvas → PDF =====

    async function htmlCanvasToPDF(html, pageSize, landscape, margins, contentWidthPX, sealImages) {
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const pw = landscape ? pgBase.h : pgBase.w;
        const ph = landscape ? pgBase.w : pgBase.h;
        const contentW = pw - margins.left - margins.right;
        const contentH = ph - margins.top - margins.bottom;
        const PX_PER_MM = 3.78;
        const SCALE = 2;
        const contentWpx = Math.round(contentW * PX_PER_MM);

        // Create visible container (behind loading overlay)
        const container = document.createElement('div');
        container.style.cssText = `
            position: fixed; left: 0; top: 0; z-index: 9998;
            width: ${contentWpx}px; background: white; padding: 0;
            font-family: 'MS PGothic', 'Yu Gothic', 'Meiryo', sans-serif;
        `;
        container.innerHTML = html;
        document.body.appendChild(container);

        // Wait for layout
        await new Promise(r => setTimeout(r, 100));

        try {
            // FIX 3: Get row Y positions for row-boundary-aware page breaks
            const table = container.querySelector('table');
            const trElements = table ? table.querySelectorAll('tr') : [];
            const rowBottoms = []; // Y bottom of each row in CSS pixels
            for (const tr of trElements) {
                const rect = tr.getBoundingClientRect();
                const tableRect = table.getBoundingClientRect();
                rowBottoms.push(Math.round(rect.bottom - tableRect.top));
            }

            // Capture with html2canvas
            const canvas = await html2canvas(container, {
                scale: SCALE,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff',
                width: contentWpx,
                windowWidth: contentWpx,
            });

            const imgW = canvas.width;
            const imgH = canvas.height;
            const pageContentHpx = Math.round(contentH * PX_PER_MM * SCALE);

            // Calculate page break points at row boundaries
            const breakPoints = [0]; // start of each page in scaled pixels
            let currentLimit = pageContentHpx;

            for (let i = 0; i < rowBottoms.length; i++) {
                const rowBottomScaled = rowBottoms[i] * SCALE;
                if (rowBottomScaled > currentLimit) {
                    // This row exceeds the page. Break BEFORE this row.
                    // Find the previous row's bottom as the break point.
                    const prevBottom = (i > 0) ? rowBottoms[i - 1] * SCALE : 0;
                    if (prevBottom > breakPoints[breakPoints.length - 1]) {
                        breakPoints.push(prevBottom);
                        currentLimit = prevBottom + pageContentHpx;
                    } else {
                        // Single row taller than page — force break at page limit
                        breakPoints.push(currentLimit);
                        currentLimit += pageContentHpx;
                    }
                }
            }

            const totalPages = breakPoints.length;

            const jsPDFLib = window.jspdf || window.jsPDF;
            const doc = new jsPDFLib.jsPDF({
                orientation: landscape ? 'landscape' : 'portrait',
                unit: 'mm', format: pageSize, compress: true,
            });

            for (let page = 0; page < totalPages; page++) {
                if (page > 0) doc.addPage();

                const srcY = breakPoints[page];
                const srcEnd = (page + 1 < breakPoints.length) ? breakPoints[page + 1] : imgH;
                const srcH = srcEnd - srcY;
                if (srcH <= 0) break;

                const pageCanvas = document.createElement('canvas');
                pageCanvas.width = imgW;
                pageCanvas.height = srcH;
                const ctx = pageCanvas.getContext('2d');
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, imgW, srcH);
                ctx.drawImage(canvas, 0, srcY, imgW, srcH, 0, 0, imgW, srcH);

                const imgData = pageCanvas.toDataURL('image/jpeg', 0.92);
                const sliceHmm = (srcH / (PX_PER_MM * SCALE));

                doc.addImage(imgData, 'JPEG', margins.left, margins.top, contentW, sliceHmm);

                // FIX 5: Add seal image on first page using anchor position from template
                if (page === 0 && sealImages.length > 0) {
                    for (const seal of sealImages) {
                        try {
                            const sealW = 18; // mm
                            const sealH = 18;
                            // Convert anchor col/row to mm position
                            // Approximate: each col ≈ contentW/maxCol, each row ≈ 4.5mm
                            const approxColW = contentW / 8; // assume ~8 cols
                            const approxRowH = 4.5; // mm per row
                            const sealX = margins.left + (seal.anchorCol * approxColW);
                            const sealY = margins.top + (seal.anchorRow * approxRowH);
                            doc.addImage(seal.dataUrl, 'PNG', sealX, sealY, sealW, sealH);
                        } catch (e) { /* skip bad images */ }
                    }
                }

                // Page number
                doc.setFontSize(8);
                doc.setTextColor(150, 150, 150);
                doc.text(`${page + 1} / ${totalPages}`, pw / 2, ph - 3, { align: 'center' });
                doc.setTextColor(0, 0, 0);
            }

            return doc.output('blob');
        } finally {
            document.body.removeChild(container);
        }
    }

    // ===== Simple table (no template) =====

    async function renderSimpleTable(headers, rows, title, pageSize, landscape, margins) {
        const jsPDFLib = window.jspdf || window.jsPDF;
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const pw = landscape ? pgBase.h : pgBase.w;
        const ph = landscape ? pgBase.w : pgBase.h;

        const doc = new jsPDFLib.jsPDF({
            orientation: landscape ? 'landscape' : 'portrait',
            unit: 'mm', format: pageSize, compress: true,
        });

        let fontName = 'helvetica';
        if (typeof FontLoader !== 'undefined') {
            await FontLoader.registerFont(doc);
            if (FontLoader.isLoaded()) fontName = 'NotoSans';
        }

        if (title) {
            doc.setFont(fontName, 'bold');
            doc.setFontSize(16);
            doc.text(title, pw / 2, 15, { align: 'center' });
        }

        doc.autoTable({
            head: [headers],
            body: rows,
            startY: title ? 25 : 10,
            styles: { font: fontName, fontSize: 9, cellPadding: 3 },
            headStyles: { fillColor: [99, 102, 241], textColor: 255, fontStyle: 'bold', halign: 'center' },
            alternateRowStyles: { fillColor: [245, 245, 255] },
            margin: { top: 10, right: 10, bottom: 10, left: 10 },
        });

        return doc.output('blob');
    }

    // ===== Parsers =====

    function parseStrings(ssXml) {
        const doc = new DOMParser().parseFromString(ssXml, 'application/xml');
        const siNodes = doc.getElementsByTagName('si');
        const arr = [];
        for (let i = 0; i < siNodes.length; i++) {
            const tNodes = siNodes[i].getElementsByTagName('t');
            let text = '';
            for (let j = 0; j < tNodes.length; j++) text += tNodes[j].textContent || '';
            arr.push(text);
        }
        return arr;
    }

    function parseColumnWidths(sheetDoc) {
        const colNodes = sheetDoc.getElementsByTagName('col');
        const widths = {};
        for (let i = 0; i < colNodes.length; i++) {
            const min = parseInt(colNodes[i].getAttribute('min'));
            const max = parseInt(colNodes[i].getAttribute('max'));
            const w = parseFloat(colNodes[i].getAttribute('width'));
            for (let c = min; c <= max; c++) widths[c] = w;
        }
        return widths;
    }

    function parseMergeCells(sheetDoc) {
        const nodes = sheetDoc.getElementsByTagName('mergeCell');
        const merges = [];
        for (let i = 0; i < nodes.length; i++) {
            const ref = nodes[i].getAttribute('ref');
            const [start, end] = ref.split(':');
            const s = parseCellRef(start);
            const e = parseCellRef(end);
            merges.push({
                startCol: s.colNum, endCol: e.colNum,
                startRow: s.rowNum, endRow: e.rowNum,
                colspan: e.colNum - s.colNum + 1,
                rowspan: e.rowNum - s.rowNum + 1,
            });
        }
        return merges;
    }

    function parseAllRows(sheetDoc, strings) {
        const rowNodes = sheetDoc.getElementsByTagName('row');
        const rows = [];
        for (let i = 0; i < rowNodes.length; i++) {
            const row = rowNodes[i];
            const rn = parseInt(row.getAttribute('r'));
            const ht = parseFloat(row.getAttribute('ht')) || 16;
            const cells = [];
            const cellNodes = row.getElementsByTagName('c');
            for (let j = 0; j < cellNodes.length; j++) {
                const c = cellNodes[j];
                const ref = c.getAttribute('r');
                const { colNum } = parseCellRef(ref);
                const s = parseInt(c.getAttribute('s') || '0');
                const t = c.getAttribute('t') || '';
                const vEl = c.getElementsByTagName('v')[0];
                let val = vEl ? vEl.textContent : '';
                let display = val;
                if (t === 's') display = strings[parseInt(val)] || '';
                cells.push({ colNum, s, t, display });
            }
            rows.push({ rowNum: rn, ht, cells });
        }
        return rows;
    }

    function parseStyles(stylesXml) {
        const doc = new DOMParser().parseFromString(stylesXml, 'application/xml');

        // Fonts
        const fontParent = doc.getElementsByTagName('fonts')[0];
        const fontNodes = fontParent ? fontParent.querySelectorAll(':scope > font') : [];
        const fonts = [];
        for (const f of fontNodes) {
            const nm = f.getElementsByTagName('name')[0];
            const sz = f.getElementsByTagName('sz')[0];
            const bNodes = f.getElementsByTagName('b');
            const colorEl = f.getElementsByTagName('color')[0];
            fonts.push({
                name: nm?.getAttribute('val') || 'sans-serif',
                size: parseFloat(sz?.getAttribute('val') || '11'),
                bold: bNodes.length > 0,
                colorRgb: colorEl?.getAttribute('rgb') || null,
                colorTheme: colorEl?.getAttribute('theme') || null,
            });
        }

        // Fills
        const fillParent = doc.getElementsByTagName('fills')[0];
        const fillNodes = fillParent ? fillParent.querySelectorAll(':scope > fill') : [];
        const fills = [];
        for (const f of fillNodes) {
            const pf = f.getElementsByTagName('patternFill')[0];
            if (!pf) { fills.push({}); continue; }
            const fg = pf.getElementsByTagName('fgColor')[0];
            fills.push({
                pattern: pf.getAttribute('patternType') || 'none',
                fgRgb: fg?.getAttribute('rgb') || null,
                fgTheme: fg?.getAttribute('theme') || null,
            });
        }

        // Borders
        const borderParent = doc.getElementsByTagName('borders')[0];
        const borderNodes = borderParent ? borderParent.querySelectorAll(':scope > border') : [];
        const borders = [];
        for (const b of borderNodes) {
            const sides = {};
            for (const side of ['left', 'right', 'top', 'bottom']) {
                const el = b.getElementsByTagName(side)[0];
                if (el && el.getAttribute('style')) {
                    const cEl = el.getElementsByTagName('color')[0];
                    sides[side] = {
                        style: el.getAttribute('style'),
                        color: cEl?.getAttribute('rgb') || (cEl?.getAttribute('auto') === 'true' ? '000000' : '000000'),
                    };
                }
            }
            borders.push(sides);
        }

        // CellXfs
        const xfParent = doc.getElementsByTagName('cellXfs')[0];
        const xfNodes = xfParent ? xfParent.querySelectorAll(':scope > xf') : [];
        const xfs = [];
        for (const xf of xfNodes) {
            const al = xf.getElementsByTagName('alignment')[0];
            xfs.push({
                fontId: parseInt(xf.getAttribute('fontId') || '0'),
                fillId: parseInt(xf.getAttribute('fillId') || '0'),
                borderId: parseInt(xf.getAttribute('borderId') || '0'),
                numFmtId: parseInt(xf.getAttribute('numFmtId') || '0'),
                alignment: al ? {
                    horizontal: al.getAttribute('horizontal') || null,
                    vertical: al.getAttribute('vertical') || null,
                    wrapText: al.getAttribute('wrapText') === '1',
                } : null,
            });
        }

        // NumFmts
        const nfNodes = doc.getElementsByTagName('numFmt');
        const numFmts = {};
        for (let i = 0; i < nfNodes.length; i++) {
            numFmts[parseInt(nfNodes[i].getAttribute('numFmtId'))] = nfNodes[i].getAttribute('formatCode');
        }

        return { fonts, fills, borders, xfs, numFmts };
    }

    function buildMergeMap(merges) {
        const map = {};
        for (const m of merges) {
            for (let r = m.startRow; r <= m.endRow; r++) {
                for (let c = m.startCol; c <= m.endCol; c++) {
                    if (r === m.startRow && c === m.startCol) {
                        map[`${r},${c}`] = { ...m, isOrigin: true };
                    } else {
                        map[`${r},${c}`] = { isOrigin: false };
                    }
                }
            }
        }
        return map;
    }

    // ===== Utilities =====

    function parseCellRef(ref) {
        const match = ref.match(/^([A-Z]+)(\d+)$/);
        if (!match) return { colNum: 1, rowNum: 1 };
        const col = match[1];
        let colNum = 0;
        for (let i = 0; i < col.length; i++) colNum = colNum * 26 + (col.charCodeAt(i) - 64);
        return { colNum, rowNum: parseInt(match[2]) };
    }

    function resolveColor(rgb, theme) {
        if (rgb && rgb.length >= 6) return rgb.length === 8 ? rgb.substring(2) : rgb;
        if (theme !== null && theme !== undefined) {
            const t = parseInt(theme);
            if (t >= 0 && t < THEME_COLORS.length) return THEME_COLORS[t];
        }
        return null;
    }

    function resolveFillColor(fill) {
        if (!fill || fill.pattern === 'none' || fill.pattern === 'gray125') return null;
        return resolveColor(fill.fgRgb, fill.fgTheme);
    }

    function formatNumber(val, numFmtId, numFmts) {
        const num = parseFloat(val);
        if (isNaN(num)) return val;
        if ([3, 4, 38, 39, 40].includes(numFmtId)) return num.toLocaleString('ja-JP');
        const fmt = numFmts[numFmtId];
        if (fmt && fmt.includes('#,##0')) return num.toLocaleString('ja-JP');
        if ((numFmtId >= 14 && numFmtId <= 22) || (fmt && (fmt.includes('yy') || fmt.includes('dd')))) {
            const date = new Date(Date.UTC(1899, 11, 30) + num * 86400000);
            return date.toLocaleDateString('ja-JP');
        }
        if (numFmtId === 49) return val;
        return val;
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // ===== FIX 5: Extract template images (seal/stamp) =====

    async function extractTemplateImages(zip, dataStart, dataEnd, rowShift) {
        const images = [];

        // Find drawing relationships from sheet rels
        const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
        const relsFile = zip.file(relsPath);
        if (!relsFile) return images;

        const relsXml = await relsFile.async('string');
        const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
        const rels = relsDoc.getElementsByTagName('Relationship');

        // Find drawing file path
        let drawingPath = null;
        for (let i = 0; i < rels.length; i++) {
            const type = rels[i].getAttribute('Type') || '';
            if (type.includes('drawing')) {
                const target = rels[i].getAttribute('Target') || '';
                drawingPath = target.startsWith('/') ? target.substring(1) : 'xl/' + target.replace('../', '');
                break;
            }
        }
        if (!drawingPath) return images;

        // Parse drawing XML for anchor positions
        const drawingFile = zip.file(drawingPath);
        if (!drawingFile) return images;
        const drawingXml = await drawingFile.async('string');
        const drawingDoc = new DOMParser().parseFromString(drawingXml, 'application/xml');

        // Build a map of rId → anchor position from drawing XML
        const anchorMap = {};
        // Try twoCellAnchor elements
        const anchors = drawingDoc.getElementsByTagName('xdr:twoCellAnchor');
        for (let i = 0; i < anchors.length; i++) {
            const from = anchors[i].getElementsByTagName('xdr:from')[0];
            if (!from) continue;
            const fromRow = parseInt(from.getElementsByTagName('xdr:row')[0]?.textContent || '0');
            const fromCol = parseInt(from.getElementsByTagName('xdr:col')[0]?.textContent || '0');

            // Get the relationship ID from blipFill
            const blips = anchors[i].getElementsByTagName('a:blip');
            for (let j = 0; j < blips.length; j++) {
                const rEmbed = blips[j].getAttribute('r:embed');
                if (rEmbed) {
                    anchorMap[rEmbed] = { row: fromRow, col: fromCol };
                }
            }
        }
        // Also try oneCellAnchor
        const oneAnchors = drawingDoc.getElementsByTagName('xdr:oneCellAnchor');
        for (let i = 0; i < oneAnchors.length; i++) {
            const from = oneAnchors[i].getElementsByTagName('xdr:from')[0];
            if (!from) continue;
            const fromRow = parseInt(from.getElementsByTagName('xdr:row')[0]?.textContent || '0');
            const fromCol = parseInt(from.getElementsByTagName('xdr:col')[0]?.textContent || '0');

            const blips = oneAnchors[i].getElementsByTagName('a:blip');
            for (let j = 0; j < blips.length; j++) {
                const rEmbed = blips[j].getAttribute('r:embed');
                if (rEmbed) {
                    anchorMap[rEmbed] = { row: fromRow, col: fromCol };
                }
            }
        }

        // Find drawing rels to map rId → image file
        const drawingRelsPath = drawingPath.replace('drawings/', 'drawings/_rels/') + '.rels';
        const drawingRelsFile = zip.file(drawingRelsPath);
        if (!drawingRelsFile) return images;

        const drawingRelsXml = await drawingRelsFile.async('string');
        const drawingRelsDoc = new DOMParser().parseFromString(drawingRelsXml, 'application/xml');
        const imgRels = drawingRelsDoc.getElementsByTagName('Relationship');

        for (let i = 0; i < imgRels.length; i++) {
            const type = imgRels[i].getAttribute('Type') || '';
            if (!type.includes('image')) continue;

            const rId = imgRels[i].getAttribute('Id') || '';
            const target = imgRels[i].getAttribute('Target') || '';
            const imgPath = target.startsWith('/') ? target.substring(1) : 'xl/' + target.replace('../', '');

            const imgFile = zip.file(imgPath);
            if (!imgFile) continue;

            try {
                const imgData = await imgFile.async('base64');
                const ext = imgPath.split('.').pop().toLowerCase();
                const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : 'image/png';
                const anchor = anchorMap[rId] || { row: 3, col: 6 }; // default to header area if no anchor

                images.push({
                    dataUrl: `data:${mime};base64,${imgData}`,
                    path: imgPath,
                    anchorRow: anchor.row,
                    anchorCol: anchor.col,
                });
            } catch (e) { /* skip unreadable images */ }
        }

        return images;
    }

    return { renderToPDF };
})();
