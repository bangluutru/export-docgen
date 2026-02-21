/**
 * HTML-PDF Renderer — Cell-perfect PDF from Excel template.
 * 
 * Approach: Parse template XML cell-by-cell, map styles to rendering params, 
 * then draw directly with jsPDF (vector, no html2canvas raster).
 * 
 * This avoids the html2canvas blank-page issue and produces small vector PDFs.
 */
const SVGPDFRenderer = (() => {
    'use strict';

    const PT_TO_MM = 0.3528; // 1pt = 0.3528mm
    const EXCEL_COL_TO_MM = 2.3; // 1 Excel col width unit ≈ 2.3mm

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

    /**
     * Main entry point
     */
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

    /**
     * Cell-perfect rendering from template
     */
    async function renderFromTemplate(templateData, dataRows, pageSize, landscape, margins) {
        const zip = templateData.zip;
        const analysis = templateData.analysis;

        // 1. Parse template sheet XML
        const sheetXml = await zip.file(templateData.sheetPaths[0]).async('string');
        const sheetDoc = new DOMParser().parseFromString(sheetXml, 'application/xml');

        // 2. Parse shared strings
        const ssFile = zip.file('xl/sharedStrings.xml');
        const ssXml = ssFile ? await ssFile.async('string') : '<sst/>';
        const strings = parseStrings(ssXml);

        // 3. Parse styles
        const stylesXml = await zip.file('xl/styles.xml').async('string');
        const styles = parseStyles(stylesXml);

        // 4. Parse columns
        const colWidths = parseColumnWidths(sheetDoc);

        // 5. Parse merge cells
        const merges = parseMergeCells(sheetDoc);

        // 6. Parse all rows from template
        const allTemplateRows = parseAllRows(sheetDoc, strings);

        // 7. Determine zones
        const dataStart = analysis.dataZone.startRowNum;
        const dataEnd = analysis.dataZone.endRowNum;
        const headerRows = allTemplateRows.filter(r => r.rowNum < dataStart);
        const tplDataRows = allTemplateRows.filter(r => r.rowNum >= dataStart && r.rowNum <= dataEnd);
        const footerRows = allTemplateRows.filter(r => r.rowNum > dataEnd);

        // 8. Style patterns from template data rows
        const stylePatterns = tplDataRows.map(tr => ({
            styles: tr.cells.map(c => c.s),
            ht: tr.ht,
        }));

        // 9. Max columns
        const maxCol = Math.max(8, ...allTemplateRows.flatMap(r => r.cells.map(c => c.colNum)));

        // 10. Build new data rows
        const newDataRows = [];
        for (let ri = 0; ri < dataRows.length; ri++) {
            const rowData = dataRows[ri];
            const rowNum = dataStart + ri;
            const patIdx = ri % Math.max(1, stylePatterns.length);
            const pattern = stylePatterns[patIdx] || stylePatterns[0];

            const cells = [];
            for (let ci = 0; ci < Math.min(rowData.length, maxCol); ci++) {
                cells.push({
                    colNum: ci + 1,
                    s: pattern?.styles[ci] || 0,
                    display: String(rowData[ci] ?? ''),
                    t: '',
                });
            }
            newDataRows.push({ rowNum, ht: pattern?.ht || 18, cells });
        }

        // 11. Shift footer rows
        const shift = newDataRows.length - (dataEnd - dataStart + 1);
        const shiftedFooter = footerRows.map(r => ({ ...r, rowNum: r.rowNum + shift }));

        // 12. Adjust merges (remove data zone merges, shift footer merges)
        const adjustedMerges = [];
        for (const m of merges) {
            if (m.startRow >= dataStart && m.endRow <= dataEnd) continue; // data zone
            if (m.startRow > dataEnd) {
                adjustedMerges.push({ ...m, startRow: m.startRow + shift, endRow: m.endRow + shift });
            } else {
                adjustedMerges.push(m);
            }
        }

        // 13. Combine all rows
        const allFinalRows = [...headerRows, ...newDataRows, ...shiftedFooter];

        // 14. Calculate page layout
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const pw = landscape ? pgBase.h : pgBase.w;
        const ph = landscape ? pgBase.w : pgBase.h;
        const contentW = pw - margins.left - margins.right;

        // Scale column widths to fit page
        let totalExcelW = 0;
        for (let c = 1; c <= maxCol; c++) totalExcelW += (colWidths[c] || 8.43);
        const scaledWidths = {};
        for (let c = 1; c <= maxCol; c++) {
            scaledWidths[c] = ((colWidths[c] || 8.43) / totalExcelW) * contentW;
        }

        // 15. Build merge map
        const mergeMap = buildMergeMap(adjustedMerges);

        // 16. Create jsPDF and draw
        const jsPDFLib = window.jspdf || window.jsPDF;
        const doc = new jsPDFLib.jsPDF({
            orientation: landscape ? 'landscape' : 'portrait',
            unit: 'mm', format: pageSize, compress: true,
        });

        // Load Unicode font
        let fontName = 'helvetica';
        if (typeof FontLoader !== 'undefined') {
            await FontLoader.registerFont(doc);
            if (FontLoader.isLoaded()) fontName = 'NotoSans';
        }

        // Draw rows page by page
        let y = margins.top;
        let pageNum = 1;

        for (const row of allFinalRows) {
            const rowH = Math.max(row.ht * PT_TO_MM, 4);

            // Check page break
            if (y + rowH > ph - margins.bottom - 5) {
                // Page number
                drawPageNumber(doc, pageNum, pw, ph, fontName);
                doc.addPage();
                pageNum++;
                y = margins.top;
            }

            // Draw cells for this row
            drawRow(doc, row, scaledWidths, mergeMap, styles, maxCol, margins.left, y, rowH, fontName);
            y += rowH;
        }

        // Last page number
        drawPageNumber(doc, pageNum, pw, ph, fontName);

        return doc.output('blob');
    }

    /**
     * Draw a single row
     */
    function drawRow(doc, row, colWidths, mergeMap, styles, maxCol, x, y, h, fontName) {
        const cellMap = {};
        for (const cell of row.cells) cellMap[cell.colNum] = cell;

        let cellX = x;
        for (let c = 1; c <= maxCol; c++) {
            const w = colWidths[c] || 10;
            const mergeKey = `${row.rowNum},${c}`;
            const merge = mergeMap[mergeKey];

            if (merge && !merge.isOrigin) {
                cellX += w;
                continue; // Covered by merge
            }

            // Calculate actual width/height considering merge
            let actualW = w;
            let actualH = h;
            if (merge && merge.isOrigin) {
                actualW = 0;
                for (let mc = merge.startCol; mc <= merge.endCol; mc++) {
                    actualW += (colWidths[mc] || 10);
                }
                // For rowspan, we'd need to know subsequent row heights — use approximate
                if (merge.rowspan > 1) {
                    actualH = h * merge.rowspan;
                }
            }

            const cell = cellMap[c];
            const xf = cell ? (styles.xfs[cell.s] || {}) : {};
            const font = styles.fonts[xf.fontId] || {};
            const fill = styles.fills[xf.fillId] || {};
            const border = styles.borders[xf.borderId] || {};
            const align = xf.alignment || {};

            // Fill color
            const bgColor = resolveFillColor(fill);
            if (bgColor) {
                const rgb = hexToRgb(bgColor);
                doc.setFillColor(rgb[0], rgb[1], rgb[2]);
                doc.rect(cellX, y, actualW, actualH, 'F');
            }

            // Borders
            drawBorders(doc, border, cellX, y, actualW, actualH);

            // Text
            if (cell && cell.display && cell.display.trim() !== '') {
                const fontSize = Math.min(font.size || 10, 13);
                const isBold = font.bold;
                doc.setFont(fontName, isBold ? 'bold' : 'normal');
                doc.setFontSize(fontSize);

                const fontColor = resolveFontColor(font);
                if (fontColor) {
                    const rgb = hexToRgb(fontColor);
                    doc.setTextColor(rgb[0], rgb[1], rgb[2]);
                } else {
                    doc.setTextColor(0, 0, 0);
                }

                // Format number
                let text = cell.display;
                const numFmtId = xf.numFmtId || 0;
                if (numFmtId > 0 && cell.t !== 's') {
                    text = formatNumber(text, numFmtId, styles.numFmts);
                }

                // Truncate to fit
                text = truncateText(doc, text, actualW - 2);

                // Alignment
                const hAlign = align.horizontal || (isNumericStr(text) ? 'right' : 'left');
                let tx = cellX + 1;
                if (hAlign === 'center') tx = cellX + actualW / 2;
                else if (hAlign === 'right') tx = cellX + actualW - 1;

                const ty = y + actualH / 2 + fontSize * 0.12;
                doc.text(text, tx, ty, { align: hAlign, baseline: 'middle' });

                doc.setTextColor(0, 0, 0); // Reset
            }

            cellX += w;
        }
    }

    function drawBorders(doc, border, x, y, w, h) {
        if (!border) return;
        const sides = [
            { name: 'top', x1: x, y1: y, x2: x + w, y2: y },
            { name: 'bottom', x1: x, y1: y + h, x2: x + w, y2: y + h },
            { name: 'left', x1: x, y1: y, x2: x, y2: y + h },
            { name: 'right', x1: x + w, y1: y, x2: x + w, y2: y + h },
        ];
        for (const side of sides) {
            const b = border[side.name];
            if (b && b.style) {
                const lw = b.style === 'thin' ? 0.2 : b.style === 'medium' ? 0.5 : 0.1;
                doc.setLineWidth(lw);
                const color = b.color || '000000';
                const rgb = hexToRgb(color.replace(/^FF/, ''));
                doc.setDrawColor(rgb[0], rgb[1], rgb[2]);
                doc.line(side.x1, side.y1, side.x2, side.y2);
            }
        }
    }

    function drawPageNumber(doc, num, pw, ph, fontName) {
        doc.setFont(fontName, 'normal');
        doc.setFontSize(8);
        doc.setTextColor(150, 150, 150);
        doc.text(`${num}`, pw / 2, ph - 5, { align: 'center' });
        doc.setTextColor(0, 0, 0);
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
        const fontNodes = doc.getElementsByTagName('fonts')[0]?.querySelectorAll(':scope > font') || [];
        const fonts = [];
        for (const f of fontNodes) {
            const nameEl = f.getElementsByTagName('name')[0];
            const szEl = f.getElementsByTagName('sz')[0];
            const bNodes = f.getElementsByTagName('b');
            const colorEl = f.getElementsByTagName('color')[0];
            fonts.push({
                name: nameEl?.getAttribute('val') || 'serif',
                size: parseFloat(szEl?.getAttribute('val') || '11'),
                bold: bNodes.length > 0,
                colorRgb: colorEl?.getAttribute('rgb') || null,
                colorTheme: colorEl?.getAttribute('theme') || null,
            });
        }

        // Fills
        const fillNodes = doc.getElementsByTagName('fills')[0]?.querySelectorAll(':scope > fill') || [];
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
        const borderNodes = doc.getElementsByTagName('borders')[0]?.querySelectorAll(':scope > border') || [];
        const borders = [];
        for (const b of borderNodes) {
            const sides = {};
            for (const side of ['left', 'right', 'top', 'bottom']) {
                const el = b.getElementsByTagName(side)[0];
                if (el && el.getAttribute('style')) {
                    const colorEl = el.getElementsByTagName('color')[0];
                    sides[side] = {
                        style: el.getAttribute('style'),
                        color: colorEl?.getAttribute('rgb') || (colorEl?.getAttribute('auto') === 'true' ? '000000' : null),
                    };
                }
            }
            borders.push(sides);
        }

        // CellXfs
        const xfNodes = doc.getElementsByTagName('cellXfs')[0]?.querySelectorAll(':scope > xf') || [];
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

    // ===== Simple table (no template) =====

    async function renderSimpleTable(headers, rows, title, pageSize, landscape, margins) {
        const jsPDFLib = window.jspdf || window.jsPDF;
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const pw = landscape ? pgBase.h : pgBase.w;
        const ph = landscape ? pgBase.w : pgBase.h;
        const contentW = pw - margins.left - margins.right;

        const doc = new jsPDFLib.jsPDF({
            orientation: landscape ? 'landscape' : 'portrait',
            unit: 'mm', format: pageSize, compress: true,
        });

        let fontName = 'helvetica';
        if (typeof FontLoader !== 'undefined') {
            await FontLoader.registerFont(doc);
            if (FontLoader.isLoaded()) fontName = 'NotoSans';
        }

        let y = margins.top;

        // Title
        if (title) {
            doc.setFont(fontName, 'bold');
            doc.setFontSize(16);
            doc.text(title, pw / 2, y + 6, { align: 'center' });
            y += 12;
        }

        // Column widths proportional
        const colW = contentW / headers.length;

        // Headers
        doc.setFillColor(68, 114, 196);
        doc.rect(margins.left, y, contentW, 7, 'F');
        doc.setFont(fontName, 'bold');
        doc.setFontSize(9);
        doc.setTextColor(255, 255, 255);
        for (let i = 0; i < headers.length; i++) {
            const x = margins.left + i * colW;
            doc.text(truncateText(doc, String(headers[i]), colW - 2), x + colW / 2, y + 4.5, { align: 'center' });
        }
        y += 7;
        doc.setTextColor(0, 0, 0);

        // Rows
        for (let ri = 0; ri < rows.length; ri++) {
            if (y + 6 > ph - margins.bottom - 5) {
                doc.addPage();
                y = margins.top;
            }

            if (ri % 2 === 1) {
                doc.setFillColor(242, 242, 242);
                doc.rect(margins.left, y, contentW, 6, 'F');
            }

            doc.setFont(fontName, 'normal');
            doc.setFontSize(8);
            doc.setDrawColor(200, 200, 200);
            doc.setLineWidth(0.1);
            doc.rect(margins.left, y, contentW, 6, 'S');

            for (let ci = 0; ci < headers.length; ci++) {
                const x = margins.left + ci * colW;
                const val = String(rows[ri]?.[ci] ?? '');
                const align = isNumericStr(val) ? 'right' : 'left';
                const tx = align === 'right' ? x + colW - 1 : x + 1;
                doc.text(truncateText(doc, val, colW - 2), tx, y + 3.8, { align });
            }
            y += 6;
        }

        return doc.output('blob');
    }

    // ===== Utility =====

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

    function resolveFontColor(font) {
        return resolveColor(font.colorRgb, font.colorTheme);
    }

    function resolveFillColor(fill) {
        if (!fill || fill.pattern === 'none' || fill.pattern === 'gray125') return null;
        return resolveColor(fill.fgRgb, fill.fgTheme);
    }

    function hexToRgb(hex) {
        if (!hex) return [0, 0, 0];
        hex = hex.replace('#', '');
        return [
            parseInt(hex.substring(0, 2), 16) || 0,
            parseInt(hex.substring(2, 4), 16) || 0,
            parseInt(hex.substring(4, 6), 16) || 0,
        ];
    }

    function formatNumber(val, numFmtId, numFmts) {
        const num = parseFloat(val);
        if (isNaN(num)) return val;
        if (numFmtId === 38 || numFmtId === 39 || numFmtId === 40 || numFmtId === 3 || numFmtId === 4) {
            return num.toLocaleString('ja-JP');
        }
        const fmt = numFmts[numFmtId];
        if (fmt && fmt.includes('#,##0')) {
            return num.toLocaleString('ja-JP');
        }
        if ((numFmtId >= 14 && numFmtId <= 22) || (fmt && (fmt.includes('yy') || fmt.includes('dd')))) {
            const epoch = new Date(1899, 11, 30);
            const date = new Date(epoch.getTime() + num * 86400000);
            return date.toLocaleDateString('ja-JP');
        }
        if (numFmtId === 49) return val; // text
        return val;
    }

    function truncateText(doc, text, maxW) {
        if (!text || maxW <= 0) return '';
        if (doc.getTextWidth(text) <= maxW) return text;
        let lo = 0, hi = text.length;
        while (lo < hi - 1) {
            const mid = (lo + hi) >> 1;
            if (doc.getTextWidth(text.substring(0, mid) + '…') <= maxW) lo = mid;
            else hi = mid;
        }
        return text.substring(0, lo) + '…';
    }

    function isNumericStr(s) {
        if (!s) return false;
        return !isNaN(String(s).replace(/[,\s¥$€]/g, ''));
    }

    return { renderToPDF };
})();
