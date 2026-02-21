/**
 * HTML-PDF Renderer — Cell-perfect HTML table rendering from Excel template.
 * 
 * BREAKTHROUGH APPROACH:
 * Instead of trying to manually position elements with jsPDF drawing commands,
 * we build a pixel-perfect HTML table from the Excel template's XML structure,
 * then convert it to PDF using jsPDF.html() — leveraging the browser's own
 * rendering engine for perfect layout.
 * 
 * Flow:
 * 1. Parse template ZIP → extract sheet XML, styles XML, shared strings
 * 2. Map every cell style to CSS (font, size, color, fill, border, alignment)
 * 3. Build HTML table with exact merge cells (colspan/rowspan)
 * 4. Replace data zone rows with new data (preserving style patterns)
 * 5. Convert HTML → PDF via jsPDF.html() or window.print()
 */
const SVGPDFRenderer = (() => {
    'use strict';

    const PT_TO_PX = 1.33; // 1pt = 1.33px
    const EXCEL_COL_TO_PX = 7.5; // 1 Excel column width unit ≈ 7.5px

    // Page sizes in mm
    const PAGE_SIZES = {
        a4: { w: 210, h: 297 },
        a3: { w: 297, h: 420 },
        letter: { w: 215.9, h: 279.4 },
    };

    // OpenXML theme colors (Excel default theme)
    const THEME_COLORS = [
        'FFFFFF', '000000', 'E7E6E6', '44546A',
        '4472C4', 'ED7D31', 'A5A5A5', 'FFC000',
        '5B9BD5', '70AD47',
    ];

    /**
     * Main entry — render data to PDF using HTML intermediate
     */
    async function renderToPDF(opts) {
        const {
            headers,
            rows,
            templateData = null,
            title = '',
            pageSize = 'a4',
            landscape = false,
            margins = { top: 8, right: 6, bottom: 8, left: 6 },
        } = opts;

        // If we have template data, use the cell-perfect approach
        if (templateData && templateData.zip) {
            return await renderFromTemplate(templateData, rows, pageSize, landscape, margins);
        }

        // Fallback: generate a clean styled HTML table for non-template data
        return await renderSimpleTable(headers, rows, title, pageSize, landscape, margins);
    }

    /**
     * Cell-perfect rendering from template
     */
    async function renderFromTemplate(templateData, dataRows, pageSize, landscape, margins) {
        const zip = templateData.zip;
        const analysis = templateData.analysis;

        // 1. Parse raw XML from the template
        const sheetXml = await zip.file(templateData.sheetPaths[0]).async('string');
        const sheetDoc = new DOMParser().parseFromString(sheetXml, 'application/xml');

        const ssFile = zip.file('xl/sharedStrings.xml');
        const ssXml = ssFile ? await ssFile.async('string') : '<sst/>';
        const ssDoc = new DOMParser().parseFromString(ssXml, 'application/xml');
        const strings = parseSharedStringsArray(ssDoc);

        const stylesXml = await zip.file('xl/styles.xml').async('string');
        const stylesDoc = new DOMParser().parseFromString(stylesXml, 'application/xml');
        const styleMap = buildStyleMap(stylesDoc);

        // 2. Parse column widths
        const colNodes = sheetDoc.getElementsByTagName('col');
        const colWidths = {};
        for (let i = 0; i < colNodes.length; i++) {
            const min = parseInt(colNodes[i].getAttribute('min'));
            const max = parseInt(colNodes[i].getAttribute('max'));
            const w = parseFloat(colNodes[i].getAttribute('width'));
            for (let c = min; c <= max; c++) {
                colWidths[c] = w;
            }
        }

        // 3. Parse merge cells
        const mergeNodes = sheetDoc.getElementsByTagName('mergeCell');
        const merges = [];
        for (let i = 0; i < mergeNodes.length; i++) {
            const ref = mergeNodes[i].getAttribute('ref');
            merges.push(parseMergeRef(ref));
        }

        // 4. Parse all rows
        const rowNodes = sheetDoc.getElementsByTagName('row');
        const allRows = [];
        for (let i = 0; i < rowNodes.length; i++) {
            const row = rowNodes[i];
            const rn = parseInt(row.getAttribute('r'));
            const ht = parseFloat(row.getAttribute('ht')) || 16;
            const cells = [];
            const cellNodes = row.getElementsByTagName('c');
            for (let j = 0; j < cellNodes.length; j++) {
                const c = cellNodes[j];
                const ref = c.getAttribute('r');
                const { col, colNum } = parseCellRef(ref);
                const s = parseInt(c.getAttribute('s') || '0');
                const t = c.getAttribute('t') || '';
                const vEl = c.getElementsByTagName('v')[0];
                const fEl = c.getElementsByTagName('f')[0];
                let val = vEl ? vEl.textContent : '';
                let display = val;
                if (t === 's') display = strings[parseInt(val)] || '';
                cells.push({ ref, col, colNum, s, t, val, display, hasFormula: !!fEl });
            }
            allRows.push({ rowNum: rn, ht, cells });
        }

        // 5. Determine zones
        const dataStart = analysis.dataZone.startRowNum;
        const dataEnd = analysis.dataZone.endRowNum;
        const headerRows = allRows.filter(r => r.rowNum < dataStart);
        const templateDataRows = allRows.filter(r => r.rowNum >= dataStart && r.rowNum <= dataEnd);
        const footerRows = allRows.filter(r => r.rowNum > dataEnd);

        // 6. Get style patterns from template data rows
        const stylePatterns = [];
        for (const tr of templateDataRows) {
            stylePatterns.push(tr.cells.map(c => c.s));
        }

        // 7. Get max columns
        const maxCol = Math.max(8, ...allRows.flatMap(r => r.cells.map(c => c.colNum)));

        // 8. Build new data rows with template styles
        const newDataRows = buildNewDataRows(dataRows, dataStart, stylePatterns, templateDataRows, maxCol, analysis);

        // 9. Compute row shift for footer
        const originalDataCount = dataEnd - dataStart + 1;
        const newDataCount = newDataRows.length;
        const shift = newDataCount - originalDataCount;

        // 10. Adjust footer row numbers
        const adjustedFooterRows = footerRows.map(r => ({
            ...r,
            rowNum: r.rowNum + shift,
        }));

        // 11. Adjust merges
        const adjustedMerges = adjustMerges(merges, dataStart, dataEnd, shift, newDataCount);

        // 12. Build HTML table
        const allFinalRows = [...headerRows, ...newDataRows, ...adjustedFooterRows];
        const html = buildHTMLTable(allFinalRows, colWidths, adjustedMerges, styleMap, maxCol, pageSize, landscape, margins);

        // 13. Convert to PDF
        return await htmlToPDF(html, pageSize, landscape, margins);
    }

    /**
     * Simple table rendering (no template)
     */
    async function renderSimpleTable(headers, rows, title, pageSize, landscape, margins) {
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const pw = landscape ? pgBase.h : pgBase.w;

        let html = `<div style="font-family: 'Times New Roman', serif; width: ${pw - margins.left - margins.right}mm; margin: 0 auto;">`;
        if (title) {
            html += `<h1 style="text-align: center; font-size: 18pt; margin: 10px 0 15px;">${escapeHtml(title)}</h1>`;
        }
        html += `<table style="width: 100%; border-collapse: collapse; font-size: 10pt;">`;
        html += `<thead><tr>`;
        for (const h of headers) {
            html += `<th style="border: 1px solid #333; padding: 4px 6px; background: #4472C4; color: white; font-weight: bold; text-align: center;">${escapeHtml(String(h))}</th>`;
        }
        html += `</tr></thead><tbody>`;
        for (let ri = 0; ri < rows.length; ri++) {
            const bg = ri % 2 === 1 ? ' background: #f2f2f2;' : '';
            html += `<tr>`;
            for (let ci = 0; ci < headers.length; ci++) {
                const val = String(rows[ri]?.[ci] || '');
                const align = isNumericStr(val) ? 'right' : (ci === 0 ? 'center' : 'left');
                html += `<td style="border: 1px solid #999; padding: 3px 5px; text-align: ${align};${bg}">${escapeHtml(val)}</td>`;
            }
            html += `</tr>`;
        }
        html += `</tbody></table></div>`;

        return await htmlToPDF(html, pageSize, landscape, margins);
    }

    // ===== HELPER FUNCTIONS =====

    function parseSharedStringsArray(ssDoc) {
        const siNodes = ssDoc.getElementsByTagName('si');
        const arr = [];
        for (let i = 0; i < siNodes.length; i++) {
            const tNodes = siNodes[i].getElementsByTagName('t');
            let text = '';
            for (let j = 0; j < tNodes.length; j++) text += tNodes[j].textContent || '';
            arr.push(text);
        }
        return arr;
    }

    function parseCellRef(ref) {
        const match = ref.match(/^([A-Z]+)(\d+)$/);
        if (!match) return { col: 'A', colNum: 1, rowNum: 1 };
        const col = match[1];
        const rowNum = parseInt(match[2]);
        let colNum = 0;
        for (let i = 0; i < col.length; i++) {
            colNum = colNum * 26 + (col.charCodeAt(i) - 64);
        }
        return { col, colNum, rowNum };
    }

    function colNumToLetter(num) {
        let s = '';
        while (num > 0) {
            num--;
            s = String.fromCharCode(65 + (num % 26)) + s;
            num = Math.floor(num / 26);
        }
        return s;
    }

    function parseMergeRef(ref) {
        const [start, end] = ref.split(':');
        const s = parseCellRef(start);
        const e = parseCellRef(end);
        return {
            ref,
            startCol: s.colNum,
            endCol: e.colNum,
            startRow: s.rowNum,
            endRow: e.rowNum,
            colspan: e.colNum - s.colNum + 1,
            rowspan: e.rowNum - s.rowNum + 1,
        };
    }

    function buildStyleMap(stylesDoc) {
        // Parse fonts
        const fontNodes = stylesDoc.getElementsByTagName('fonts')[0]?.querySelectorAll(':scope > font') || [];
        const fonts = [];
        for (const f of fontNodes) {
            const nameEl = f.getElementsByTagName('name')[0];
            const szEl = f.getElementsByTagName('sz')[0];
            const bNodes = f.getElementsByTagName('b');
            const iNodes = f.getElementsByTagName('i');
            const colorEl = f.getElementsByTagName('color')[0];
            fonts.push({
                name: nameEl?.getAttribute('val') || 'serif',
                size: parseFloat(szEl?.getAttribute('val') || '11'),
                bold: bNodes.length > 0,
                italic: iNodes.length > 0,
                colorRgb: colorEl?.getAttribute('rgb') || null,
                colorTheme: colorEl?.getAttribute('theme') || null,
            });
        }

        // Parse fills
        const fillNodes = stylesDoc.getElementsByTagName('fills')[0]?.querySelectorAll(':scope > fill') || [];
        const fills = [];
        for (const f of fillNodes) {
            const pf = f.getElementsByTagName('patternFill')[0];
            if (!pf) { fills.push(null); continue; }
            const fg = pf.getElementsByTagName('fgColor')[0];
            const bg = pf.getElementsByTagName('bgColor')[0];
            fills.push({
                pattern: pf.getAttribute('patternType') || 'none',
                fgRgb: fg?.getAttribute('rgb') || null,
                fgTheme: fg?.getAttribute('theme') || null,
                bgRgb: bg?.getAttribute('rgb') || null,
            });
        }

        // Parse borders
        const borderNodes = stylesDoc.getElementsByTagName('borders')[0]?.querySelectorAll(':scope > border') || [];
        const borders = [];
        for (const b of borderNodes) {
            const sides = {};
            for (const side of ['left', 'right', 'top', 'bottom']) {
                const el = b.getElementsByTagName(side)[0];
                if (el && el.getAttribute('style')) {
                    const colorEl = el.getElementsByTagName('color')[0];
                    sides[side] = {
                        style: el.getAttribute('style'),
                        color: colorEl?.getAttribute('rgb') || colorEl?.getAttribute('auto') === 'true' ? '000000' : null,
                    };
                }
            }
            borders.push(sides);
        }

        // Parse cellXfs
        const xfNodes = stylesDoc.getElementsByTagName('cellXfs')[0]?.querySelectorAll(':scope > xf') || [];
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

        // Parse number formats
        const nfNodes = stylesDoc.getElementsByTagName('numFmt');
        const numFmts = {};
        for (let i = 0; i < nfNodes.length; i++) {
            numFmts[parseInt(nfNodes[i].getAttribute('numFmtId'))] = nfNodes[i].getAttribute('formatCode');
        }

        return { fonts, fills, borders, xfs, numFmts };
    }

    function styleToCss(styleIdx, styleMap) {
        const xf = styleMap.xfs[styleIdx];
        if (!xf) return '';

        const parts = [];

        // Font
        const font = styleMap.fonts[xf.fontId];
        if (font) {
            const fontFamily = font.name.includes('Gothic') || font.name.includes('Sans')
                ? `'${font.name}', 'Noto Sans', sans-serif`
                : `'${font.name}', 'Times New Roman', serif`;
            parts.push(`font-family: ${fontFamily}`);
            parts.push(`font-size: ${Math.min(font.size, 14)}pt`);
            if (font.bold) parts.push('font-weight: bold');
            if (font.italic) parts.push('font-style: italic');
            const color = resolveColor(font.colorRgb, font.colorTheme);
            if (color && color !== '000000') parts.push(`color: #${color}`);
        }

        // Fill
        const fill = styleMap.fills[xf.fillId];
        if (fill && fill.pattern !== 'none' && fill.pattern !== 'gray125') {
            const bgColor = resolveColor(fill.fgRgb, fill.fgTheme);
            if (bgColor) parts.push(`background-color: #${bgColor}`);
        }

        // Border
        const border = styleMap.borders[xf.borderId];
        if (border) {
            for (const side of ['left', 'right', 'top', 'bottom']) {
                if (border[side]) {
                    const bStyle = mapBorderStyle(border[side].style);
                    const bColor = border[side].color || '000000';
                    parts.push(`border-${side}: ${bStyle} #${bColor.replace(/^FF/, '')}`);
                }
            }
        }

        // Alignment
        if (xf.alignment) {
            if (xf.alignment.horizontal) parts.push(`text-align: ${xf.alignment.horizontal}`);
            if (xf.alignment.vertical) {
                const vMap = { top: 'top', center: 'middle', bottom: 'bottom' };
                parts.push(`vertical-align: ${vMap[xf.alignment.vertical] || 'middle'}`);
            }
            if (xf.alignment.wrapText) parts.push('word-wrap: break-word');
        }

        return parts.join('; ');
    }

    function resolveColor(rgb, theme) {
        if (rgb && rgb.length >= 6) {
            return rgb.length === 8 ? rgb.substring(2) : rgb;
        }
        if (theme !== null && theme !== undefined) {
            const t = parseInt(theme);
            if (t >= 0 && t < THEME_COLORS.length) return THEME_COLORS[t];
        }
        return null;
    }

    function mapBorderStyle(xlStyle) {
        const map = {
            thin: '1px solid', medium: '2px solid', thick: '3px solid',
            dashed: '1px dashed', dotted: '1px dotted',
            hair: '0.5px solid', double: '3px double',
        };
        return map[xlStyle] || '1px solid';
    }

    function formatNumber(val, numFmtId, numFmts) {
        if (!val || val === '') return '';
        const num = parseFloat(val);
        if (isNaN(num)) return val;

        // Built-in format IDs
        if (numFmtId === 38 || numFmtId === 39 || numFmtId === 40) {
            // Japanese number format: #,##0 / #,##0.00
            return num.toLocaleString('ja-JP');
        }
        if (numFmtId === 3 || numFmtId === 4) {
            return num.toLocaleString();
        }

        // Custom format
        const fmt = numFmts[numFmtId];
        if (fmt && fmt.includes('#,##0')) {
            const decimals = (fmt.match(/\.0+/) || [''])[0].length - 1;
            return num.toLocaleString('ja-JP', {
                minimumFractionDigits: Math.max(0, decimals),
                maximumFractionDigits: Math.max(0, decimals),
            });
        }

        // Date format (for numFmtId 14-22 or custom date formats)
        if ((numFmtId >= 14 && numFmtId <= 22) || (fmt && (fmt.includes('yy') || fmt.includes('mm') || fmt.includes('dd')))) {
            const date = excelDateToJS(num);
            if (date) {
                return date.toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' });
            }
        }

        // numFmtId 49 = text format
        if (numFmtId === 49) return val;

        return val;
    }

    function excelDateToJS(serial) {
        if (serial < 1) return null;
        const epoch = new Date(1899, 11, 30);
        return new Date(epoch.getTime() + serial * 86400000);
    }

    function buildNewDataRows(dataRows, dataStart, stylePatterns, templateDataRows, maxCol, analysis) {
        const newRows = [];

        for (let ri = 0; ri < dataRows.length; ri++) {
            const rowData = dataRows[ri];
            const rowNum = dataStart + ri;

            // Cycle through style patterns
            const patternIdx = ri % Math.max(1, stylePatterns.length);
            const patternStyles = stylePatterns[patternIdx] || [];

            // Use template data row structure as a guide
            const templateRow = templateDataRows[patternIdx] || templateDataRows[0];
            const ht = templateRow?.ht || 18;

            const cells = [];
            for (let ci = 0; ci < Math.min(rowData.length, maxCol); ci++) {
                const colNum = ci + 1;
                const ref = colNumToLetter(colNum) + rowNum;
                const style = patternStyles[ci] !== undefined ? patternStyles[ci] : 0;
                cells.push({
                    ref,
                    col: colNumToLetter(colNum),
                    colNum,
                    s: style,
                    display: String(rowData[ci] || ''),
                    val: String(rowData[ci] || ''),
                });
            }

            newRows.push({ rowNum, ht, cells });
        }

        return newRows;
    }

    function adjustMerges(merges, dataStart, dataEnd, shift, newDataCount) {
        const result = [];
        const newDataEnd = dataStart + newDataCount - 1;

        for (const m of merges) {
            if (m.startRow >= dataStart && m.endRow <= dataEnd) {
                // Data zone merge — skip (data rows don't have merges by default)
                continue;
            } else if (m.startRow > dataEnd) {
                // Footer merge — shift
                result.push({
                    ...m,
                    startRow: m.startRow + shift,
                    endRow: m.endRow + shift,
                });
            } else {
                // Header merge — keep as-is
                result.push(m);
            }
        }

        return result;
    }

    function buildHTMLTable(allRows, colWidths, merges, styleMap, maxCol, pageSize, landscape, margins) {
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const totalWidthMm = (landscape ? pgBase.h : pgBase.w) - margins.left - margins.right;

        // Calculate total Excel width to scale columns proportionally
        let totalExcelWidth = 0;
        for (let c = 1; c <= maxCol; c++) {
            totalExcelWidth += (colWidths[c] || 8.43);
        }

        // Build merged cell lookup: key = "row,col" → { colspan, rowspan, isOrigin }
        const mergeMap = {};
        for (const m of merges) {
            for (let r = m.startRow; r <= m.endRow; r++) {
                for (let c = m.startCol; c <= m.endCol; c++) {
                    if (r === m.startRow && c === m.startCol) {
                        mergeMap[`${r},${c}`] = { colspan: m.colspan, rowspan: m.rowspan, isOrigin: true };
                    } else {
                        mergeMap[`${r},${c}`] = { isOrigin: false };
                    }
                }
            }
        }

        let html = `<table style="border-collapse: collapse; width: ${totalWidthMm}mm; table-layout: fixed;">`;

        // Column groups for widths
        html += '<colgroup>';
        for (let c = 1; c <= maxCol; c++) {
            const w = colWidths[c] || 8.43;
            const pct = (w / totalExcelWidth * 100).toFixed(2);
            html += `<col style="width: ${pct}%;">`;
        }
        html += '</colgroup>';

        // Render rows
        for (const row of allRows) {
            const htPx = Math.round(row.ht * PT_TO_PX);
            html += `<tr style="height: ${htPx}px;">`;

            // Build cell map for this row
            const cellMap = {};
            for (const cell of row.cells) {
                cellMap[cell.colNum] = cell;
            }

            for (let c = 1; c <= maxCol; c++) {
                const mergeKey = `${row.rowNum},${c}`;
                const merge = mergeMap[mergeKey];

                if (merge && !merge.isOrigin) {
                    // Skip — covered by a merge
                    continue;
                }

                const cell = cellMap[c];
                const colspanAttr = merge && merge.colspan > 1 ? ` colspan="${merge.colspan}"` : '';
                const rowspanAttr = merge && merge.rowspan > 1 ? ` rowspan="${merge.rowspan}"` : '';

                let css = 'padding: 2px 4px; overflow: hidden; white-space: nowrap;';
                let content = '';

                if (cell) {
                    css += styleToCss(cell.s, styleMap);

                    // Format numbers
                    const xf = styleMap.xfs[cell.s];
                    const numFmtId = xf?.numFmtId || 0;
                    if (cell.display && numFmtId > 0 && cell.t !== 's') {
                        content = formatNumber(cell.display, numFmtId, styleMap.numFmts);
                    } else {
                        content = cell.display || '';
                    }
                }

                html += `<td${colspanAttr}${rowspanAttr} style="${css}">${escapeHtml(content)}</td>`;
            }

            html += '</tr>';
        }

        html += '</table>';
        return html;
    }

    async function htmlToPDF(html, pageSize, landscape, margins) {
        const jsPDFLib = window.jspdf || window.jsPDF;

        const doc = new jsPDFLib.jsPDF({
            orientation: landscape ? 'landscape' : 'portrait',
            unit: 'mm',
            format: pageSize,
            compress: true,
        });

        // Create a hidden container
        const container = document.createElement('div');
        container.style.cssText = `
            position: absolute; left: -9999px; top: 0;
            font-family: 'Times New Roman', serif;
        `;
        container.innerHTML = html;
        document.body.appendChild(container);

        try {
            const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
            const pw = landscape ? pgBase.h : pgBase.w;
            const ph = landscape ? pgBase.w : pgBase.h;

            await doc.html(container, {
                x: margins.left,
                y: margins.top,
                width: pw - margins.left - margins.right,
                windowWidth: (pw - margins.left - margins.right) * 3.78, // mm to px ratio
                autoPaging: 'text',
                margin: [margins.top, margins.right, margins.bottom, margins.left],
                html2canvas: {
                    scale: 2,
                    useCORS: true,
                    logging: false,
                },
            });

            return doc.output('blob');
        } finally {
            document.body.removeChild(container);
        }
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function isNumericStr(s) {
        if (!s || s === '') return false;
        const cleaned = String(s).replace(/[,\s¥$€]/g, '');
        return !isNaN(cleaned) && !isNaN(parseFloat(cleaned));
    }

    return { renderToPDF };
})();
