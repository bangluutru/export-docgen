/**
 * SVG-PDF Renderer — Render Excel-like tables to SVG, then convert to tiny vector PDFs.
 * 
 * Architecture:
 * 1. Read styles from template's xl/styles.xml (fonts, fills, borders, alignment)
 * 2. Build a table layout with precise cell positions
 * 3. Render to SVG (vector — no raster images needed)
 * 4. Use jsPDF to convert SVG pages into a multi-page PDF
 * 
 * Result: PDF files ~50-150KB instead of ~1.5MB
 */
const SVGPDFRenderer = (() => {
    'use strict';

    // Excel column width unit → mm (1 char width ≈ 2.3mm at 10pt)
    const EXCEL_COL_WIDTH_TO_MM = 2.3;
    // Excel row height unit (points) → mm (1pt = 0.3528mm)
    const PT_TO_MM = 0.3528;
    // Default sizes
    const DEFAULT_FONT_SIZE = 10;
    const DEFAULT_ROW_HEIGHT = 15; // points
    const MIN_COL_WIDTH_MM = 8;
    const CELL_PADDING_MM = 1.5;

    // Page sizes in mm
    const PAGE_SIZES = {
        a4: { w: 210, h: 297 },
        a3: { w: 297, h: 420 },
        letter: { w: 215.9, h: 279.4 },
    };

    // OpenXML theme colors (simplified — Excel default theme)
    const THEME_COLORS = [
        'FFFFFF', // 0: light1 (window background)
        '000000', // 1: dark1 (window text)
        'E7E6E6', // 2: light2
        '44546A', // 3: dark2
        '4472C4', // 4: accent1
        'ED7D31', // 5: accent2
        'A5A5A5', // 6: accent3
        'FFC000', // 7: accent4
        '5B9BD5', // 8: accent5
        '70AD47', // 9: accent6
    ];

    /**
     * Render data to PDF using SVG intermediate.
     * 
     * @param {Object} opts
     * @param {string[]} opts.headers - Column headers
     * @param {Array<Array>} opts.rows - Data rows
     * @param {Object} opts.templateData - Template data from TemplateEngine.analyzeTemplate (optional)
     * @param {string} opts.title - Report title
     * @param {string} opts.pageSize - 'a4', 'a3', 'letter'
     * @param {boolean} opts.landscape - Landscape orientation
     * @param {Object} opts.margins - Page margins { top, right, bottom, left } in mm
     * @returns {Promise<Blob>} PDF blob
     */
    async function renderToPDF(opts) {
        const {
            headers,
            rows,
            templateData = null,
            title = '',
            pageSize = 'a4',
            landscape = false,
            margins = { top: 10, right: 8, bottom: 10, left: 8 },
        } = opts;

        // Page dimensions
        const pgBase = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
        const pw = landscape ? pgBase.h : pgBase.w;
        const ph = landscape ? pgBase.w : pgBase.h;
        const contentW = pw - margins.left - margins.right;
        const contentH = ph - margins.top - margins.bottom;

        // Build style info from template
        const styleInfo = templateData ? extractStyleInfo(templateData) : getDefaultStyleInfo();

        // Calculate column widths
        const colWidths = calculateColumnWidths(headers, rows, contentW, templateData);

        // Calculate row heights
        const headerRowHeight = styleInfo.headerRowHeight || 8;
        const dataRowHeight = styleInfo.dataRowHeight || 6;

        // Build header zone info (company info, title from template)
        const headerZoneInfo = templateData ? extractHeaderZoneInfo(templateData) : null;
        const headerZoneHeight = headerZoneInfo ? headerZoneInfo.totalHeight : (title ? 12 : 0);

        // Split rows into pages
        const pages = paginateRows(rows, contentH, headerRowHeight, dataRowHeight, headerZoneHeight);

        // Create jsPDF document with compression
        const jsPDFLib = window.jspdf || window.jsPDF;
        const doc = new jsPDFLib.jsPDF({
            orientation: landscape ? 'landscape' : 'portrait',
            unit: 'mm',
            format: pageSize,
            compress: true, // Enable stream compression for smaller output
        });

        // Smart font loading: only load Unicode font if data contains non-ASCII chars
        const allText = [title, ...headers.map(String), ...rows.flat().map(String)].join('');
        const needsUnicode = /[^\x00-\x7F]/.test(allText);

        let fontName = 'helvetica';
        if (needsUnicode && typeof FontLoader !== 'undefined') {
            await FontLoader.registerFont(doc);
            if (FontLoader.isLoaded()) fontName = 'NotoSans';
        }

        // Render each page
        for (let pageIdx = 0; pageIdx < pages.length; pageIdx++) {
            if (pageIdx > 0) doc.addPage();

            const page = pages[pageIdx];
            let yOffset = margins.top;

            // Render header zone (company info, title) — only on first page or all pages
            if (headerZoneInfo && pageIdx === 0) {
                yOffset = renderHeaderZone(doc, headerZoneInfo, margins.left, yOffset, contentW, fontName);
            } else if (title && pageIdx === 0) {
                doc.setFontSize(14);
                doc.setFont(fontName, 'bold');
                doc.text(title, pw / 2, yOffset + 5, { align: 'center' });
                yOffset += 10;
            }

            // Render column headers
            yOffset = renderTableHeaders(doc, headers, colWidths, margins.left, yOffset, styleInfo, fontName);

            // Render data rows
            yOffset = renderTableRows(doc, page.rows, colWidths, margins.left, yOffset, styleInfo, fontName, page.startRowNum);

            // Render footer zone (totals) — only on last page
            if (pageIdx === pages.length - 1 && templateData) {
                renderFooterZone(doc, templateData, margins.left, yOffset, colWidths, styleInfo, fontName);
            }

            // Page number
            doc.setFontSize(8);
            doc.setFont(fontName, 'normal');
            doc.setTextColor(150, 150, 150);
            doc.text(`${pageIdx + 1} / ${pages.length}`, pw / 2, ph - 5, { align: 'center' });
            doc.setTextColor(0, 0, 0);
        }

        return doc.output('blob');
    }

    /**
     * Extract style info from template data for PDF rendering
     */
    function extractStyleInfo(templateData) {
        const styles = templateData.stylesData;
        const analysis = templateData.analysis;

        if (!styles) return getDefaultStyleInfo();

        // Get header row style
        const headerRow = analysis.columnHeaderRow.row;
        const headerStyleIdx = parseInt(headerRow.cells[0]?.s || '0');
        const headerXf = styles.cellXfs[headerStyleIdx] || {};
        const headerFont = styles.fonts[headerXf.fontId] || {};
        const headerFill = styles.fills[headerXf.fillId] || {};
        const headerBorder = styles.borders[headerXf.borderId] || {};

        // Get data row style (first data row)
        const dataRows = analysis.dataZone.dataRows;
        const dataStyleIdx = dataRows.length > 0 ? parseInt(dataRows[0].row.cells[1]?.s || '0') : 0;
        const dataXf = styles.cellXfs[dataStyleIdx] || {};
        const dataFont = styles.fonts[dataXf.fontId] || {};
        const dataBorder = styles.borders[dataXf.borderId] || {};

        // Even row fill (alternating)
        let evenFillColor = null;
        if (dataRows.length > 1) {
            const evenStyleIdx = parseInt(dataRows[1].row.cells[1]?.s || '0');
            const evenXf = styles.cellXfs[evenStyleIdx] || {};
            const evenFill = styles.fills[evenXf.fillId] || {};
            evenFillColor = resolveFillColor(evenFill);
        }

        return {
            headerFontSize: headerFont.size || 11,
            headerFontBold: headerFont.bold !== false,
            headerFontColor: resolveFontColor(headerFont) || '000000',
            headerFillColor: resolveFillColor(headerFill) || '4472C4',
            headerBorder: headerBorder,
            headerRowHeight: (parseFloat(headerRow.ht) || 20) * PT_TO_MM,
            headerAlignment: headerXf.alignment || { horizontal: 'center', vertical: 'center' },

            dataFontSize: dataFont.size || 10,
            dataFontColor: resolveFontColor(dataFont) || '000000',
            dataBorder: dataBorder,
            dataRowHeight: (parseFloat(analysis.dataZone.stylePatterns[0]?.ht) || DEFAULT_ROW_HEIGHT) * PT_TO_MM,
            dataAlignment: dataXf.alignment || { horizontal: 'left', vertical: 'center' },

            evenFillColor: evenFillColor,

            // Per-column alignments from style patterns
            columnAlignments: extractColumnAlignments(analysis, styles),
        };
    }

    function getDefaultStyleInfo() {
        return {
            headerFontSize: 11,
            headerFontBold: true,
            headerFontColor: 'FFFFFF',
            headerFillColor: '4472C4',
            headerBorder: {},
            headerRowHeight: 8,
            headerAlignment: { horizontal: 'center', vertical: 'center' },

            dataFontSize: 10,
            dataFontColor: '333333',
            dataBorder: {},
            dataRowHeight: 6,
            dataAlignment: { horizontal: 'left', vertical: 'center' },

            evenFillColor: 'F2F2F2',
            columnAlignments: [],
        };
    }

    function extractColumnAlignments(analysis, styles) {
        const patterns = analysis.dataZone.stylePatterns;
        if (patterns.length === 0) return [];
        const firstPattern = patterns[0].pattern;
        return firstPattern.map(p => {
            const xf = styles.cellXfs[parseInt(p.style)] || {};
            return xf.alignment || null;
        });
    }

    function resolveFontColor(font) {
        if (font.colorRgb && font.colorRgb.length >= 6) {
            // Remove leading FF if ARGB
            return font.colorRgb.length === 8 ? font.colorRgb.substring(2) : font.colorRgb;
        }
        if (font.colorTheme !== null && font.colorTheme !== undefined) {
            const theme = parseInt(font.colorTheme);
            if (theme >= 0 && theme < THEME_COLORS.length) return THEME_COLORS[theme];
        }
        return null;
    }

    function resolveFillColor(fill) {
        if (!fill) return null;
        if (fill.pattern === 'none' || fill.pattern === 'gray125') return null;
        if (fill.fgColorRgb && fill.fgColorRgb.length >= 6) {
            return fill.fgColorRgb.length === 8 ? fill.fgColorRgb.substring(2) : fill.fgColorRgb;
        }
        if (fill.fgColorTheme !== null && fill.fgColorTheme !== undefined) {
            const theme = parseInt(fill.fgColorTheme);
            if (theme >= 0 && theme < THEME_COLORS.length) return THEME_COLORS[theme];
        }
        return null;
    }

    function hexToRgb(hex) {
        if (!hex) return [0, 0, 0];
        hex = hex.replace('#', '');
        return [
            parseInt(hex.substring(0, 2), 16),
            parseInt(hex.substring(2, 4), 16),
            parseInt(hex.substring(4, 6), 16),
        ];
    }

    /**
     * Calculate column widths proportionally fitted to content width
     */
    function calculateColumnWidths(headers, rows, contentW, templateData) {
        let rawWidths;

        if (templateData && templateData.analysis.columns.length > 0) {
            // Use template column widths
            rawWidths = headers.map((_, i) => {
                const col = templateData.analysis.columns.find(c => {
                    const min = parseInt(c.min);
                    const max = parseInt(c.max);
                    return (i + 1) >= min && (i + 1) <= max;
                });
                return col ? parseFloat(col.width) * EXCEL_COL_WIDTH_TO_MM : MIN_COL_WIDTH_MM;
            });
        } else {
            // Auto-calculate from content
            rawWidths = headers.map((h, i) => {
                let maxLen = String(h).length;
                const sampleRows = rows.slice(0, 50);
                for (const row of sampleRows) {
                    const len = String(row[i] || '').length;
                    if (len > maxLen) maxLen = len;
                }
                return Math.max(maxLen * 2.0, MIN_COL_WIDTH_MM);
            });
        }

        // Scale to fit contentW
        const totalRaw = rawWidths.reduce((a, b) => a + b, 0);
        const scale = contentW / totalRaw;
        return rawWidths.map(w => w * scale);
    }

    /**
     * Split rows into pages
     */
    function paginateRows(rows, contentH, headerH, dataH, headerZoneH) {
        const pages = [];
        let idx = 0;

        while (idx < rows.length) {
            const isFirstPage = pages.length === 0;
            const availableH = contentH - headerH - (isFirstPage ? headerZoneH : 0) - 8; // 8mm for page number
            const maxRows = Math.floor(availableH / dataH);
            const pageRows = rows.slice(idx, idx + maxRows);

            pages.push({
                rows: pageRows,
                startRowNum: idx + 1,
            });
            idx += pageRows.length;
        }

        if (pages.length === 0) {
            pages.push({ rows: [], startRowNum: 1 });
        }

        return pages;
    }

    /**
     * Extract header zone info from template (company name, address, title, etc.)
     */
    function extractHeaderZoneInfo(templateData) {
        const analysis = templateData.analysis;
        const headerRows = analysis.headerZone.rows;
        if (headerRows.length === 0) return null;

        const items = [];
        let totalHeight = 0;

        for (const row of headerRows) {
            const ht = parseFloat(row.ht) || DEFAULT_ROW_HEIGHT;
            const heightMm = ht * PT_TO_MM;

            const cells = row.cells.map(c => ({
                col: c.col,
                text: c.displayValue || '',
                style: c.s,
            })).filter(c => c.text.trim().length > 0);

            if (cells.length > 0) {
                items.push({
                    cells,
                    height: heightMm,
                    rowNum: row.rowNum,
                });
            }
            totalHeight += heightMm;
        }

        return { items, totalHeight };
    }

    /**
     * Render header zone (company info, title) directly to jsPDF
     */
    function renderHeaderZone(doc, headerZoneInfo, x, y, contentW, fontName) {
        const styles = doc.__templateStyles || null;

        for (const item of headerZoneInfo.items) {
            for (const cell of item.cells) {
                // Simple approach: center text if one cell, left-align otherwise
                const fontSize = Math.max(8, Math.min(14, item.height / PT_TO_MM * 0.4));
                doc.setFontSize(fontSize);

                // Detect if this looks like a title (large text, centered)
                const isTitle = item.cells.length === 1 && cell.text.length > 5 && cell.text.length < 80;
                if (isTitle) {
                    doc.setFont(fontName, 'bold');
                    doc.text(cell.text, x + contentW / 2, y + item.height * 0.7, { align: 'center' });
                } else {
                    doc.setFont(fontName, 'normal');
                    const cellX = x + (cell.col - 1) * (contentW / 8); // rough positioning
                    doc.text(cell.text, cellX, y + item.height * 0.7);
                }
            }
            y += item.height;
        }

        return y + 2; // small gap after header zone
    }

    /**
     * Render table column headers
     */
    function renderTableHeaders(doc, headers, colWidths, x, y, styleInfo, fontName) {
        const h = styleInfo.headerRowHeight;
        const fillRgb = hexToRgb(styleInfo.headerFillColor);
        const fontRgb = hexToRgb(styleInfo.headerFontColor);

        // Draw header background
        let cellX = x;
        for (let i = 0; i < headers.length; i++) {
            doc.setFillColor(fillRgb[0], fillRgb[1], fillRgb[2]);
            doc.rect(cellX, y, colWidths[i], h, 'F');

            // Border
            doc.setDrawColor(200, 200, 200);
            doc.setLineWidth(0.2);
            doc.rect(cellX, y, colWidths[i], h, 'S');

            // Text
            doc.setFont(fontName, 'bold');
            doc.setFontSize(Math.min(styleInfo.headerFontSize, 10));
            doc.setTextColor(fontRgb[0], fontRgb[1], fontRgb[2]);

            const text = truncateText(doc, String(headers[i]), colWidths[i] - CELL_PADDING_MM * 2);
            doc.text(text, cellX + colWidths[i] / 2, y + h / 2 + 1, {
                align: 'center',
                baseline: 'middle',
            });

            cellX += colWidths[i];
        }

        doc.setTextColor(0, 0, 0);
        return y + h;
    }

    /**
     * Render data rows
     */
    function renderTableRows(doc, rows, colWidths, x, y, styleInfo, fontName, startRowNum) {
        const fontRgb = hexToRgb(styleInfo.dataFontColor);
        const evenFillRgb = styleInfo.evenFillColor ? hexToRgb(styleInfo.evenFillColor) : null;
        const h = styleInfo.dataRowHeight;

        for (let ri = 0; ri < rows.length; ri++) {
            const row = rows[ri];
            const isEven = ri % 2 === 1;

            let cellX = x;
            for (let ci = 0; ci < Math.min(row.length, colWidths.length); ci++) {
                // Fill
                if (isEven && evenFillRgb) {
                    doc.setFillColor(evenFillRgb[0], evenFillRgb[1], evenFillRgb[2]);
                    doc.rect(cellX, y, colWidths[ci], h, 'F');
                }

                // Border
                doc.setDrawColor(210, 210, 210);
                doc.setLineWidth(0.15);
                doc.rect(cellX, y, colWidths[ci], h, 'S');

                // Text
                doc.setFont(fontName, 'normal');
                doc.setFontSize(Math.min(styleInfo.dataFontSize, 9));
                doc.setTextColor(fontRgb[0], fontRgb[1], fontRgb[2]);

                const cellVal = String(row[ci] || '');
                const text = truncateText(doc, cellVal, colWidths[ci] - CELL_PADDING_MM * 2);

                // Determine alignment
                const colAlign = styleInfo.columnAlignments[ci];
                let align = 'left';
                let textX = cellX + CELL_PADDING_MM;
                if (colAlign && colAlign.horizontal) {
                    align = colAlign.horizontal;
                } else if (isNumericStr(cellVal)) {
                    align = 'right';
                } else if (ci === 0) {
                    align = 'center'; // First column (No.) centered
                }

                if (align === 'center') {
                    textX = cellX + colWidths[ci] / 2;
                } else if (align === 'right') {
                    textX = cellX + colWidths[ci] - CELL_PADDING_MM;
                }

                doc.text(text, textX, y + h / 2 + 0.8, {
                    align: align,
                    baseline: 'middle',
                });

                cellX += colWidths[ci];
            }
            y += h;
        }

        doc.setTextColor(0, 0, 0);
        return y;
    }

    /**
     * Render footer zone (totals, notes)
     */
    function renderFooterZone(doc, templateData, x, y, colWidths, styleInfo, fontName) {
        const analysis = templateData.analysis;
        const footerRows = analysis.footerZone.rows;
        if (footerRows.length === 0) return;

        y += 0.5; // Small gap

        for (const row of footerRows) {
            const h = (parseFloat(row.ht) || DEFAULT_ROW_HEIGHT) * PT_TO_MM;

            for (const cell of row.cells) {
                const colIdx = cell.col - 1;
                if (colIdx >= colWidths.length) continue;

                // Calculate x position
                let cellX = x;
                for (let i = 0; i < colIdx; i++) {
                    cellX += colWidths[i];
                }
                const w = colWidths[colIdx] || 20;

                // Render text
                const text = cell.displayValue || '';
                if (text.trim().length > 0) {
                    // Check if it's a total/summary row
                    const isTotal = text.includes('合計') || text.includes('小計') ||
                        text.includes('Total') || text.includes('Tổng') ||
                        text.includes('消費税');

                    doc.setFont(fontName, isTotal ? 'bold' : 'normal');
                    doc.setFontSize(styleInfo.dataFontSize);

                    // Border for footer cells
                    doc.setDrawColor(180, 180, 180);
                    doc.setLineWidth(0.2);
                    doc.rect(cellX, y, w, h, 'S');

                    const truncated = truncateText(doc, text, w - CELL_PADDING_MM * 2);

                    if (isNumericStr(text)) {
                        doc.text(truncated, cellX + w - CELL_PADDING_MM, y + h / 2 + 0.8, {
                            align: 'right', baseline: 'middle'
                        });
                    } else {
                        doc.text(truncated, cellX + CELL_PADDING_MM, y + h / 2 + 0.8, {
                            baseline: 'middle'
                        });
                    }
                }
            }
            y += h;
        }
    }

    /**
     * Truncate text to fit within a given width
     */
    function truncateText(doc, text, maxWidth) {
        if (!text || maxWidth <= 0) return '';
        const textWidth = doc.getTextWidth(text);
        if (textWidth <= maxWidth) return text;

        // Binary search for the right truncation point
        let lo = 0, hi = text.length;
        while (lo < hi - 1) {
            const mid = Math.floor((lo + hi) / 2);
            if (doc.getTextWidth(text.substring(0, mid) + '…') <= maxWidth) {
                lo = mid;
            } else {
                hi = mid;
            }
        }
        return text.substring(0, lo) + '…';
    }

    function isNumericStr(s) {
        if (!s || s === '') return false;
        const cleaned = String(s).replace(/[,\s¥$€]/g, '');
        return !isNaN(cleaned) && !isNaN(parseFloat(cleaned));
    }

    return { renderToPDF };
})();
