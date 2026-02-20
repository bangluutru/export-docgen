/**
 * Template Engine — Analyze template XLSX and generate new files using template formatting
 * Approach: Clone the template ZIP entirely, only replacing data area rows
 */
const TemplateEngine = (() => {
    const XLSX_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    // --- Helper functions ---
    function colToRef(col) {
        let ref = '';
        while (col > 0) {
            col--;
            ref = String.fromCharCode(65 + (col % 26)) + ref;
            col = Math.floor(col / 26);
        }
        return ref;
    }

    function refToCol(ref) {
        const match = ref.match(/^([A-Z]+)/);
        if (!match) return 1;
        let col = 0;
        for (const ch of match[1]) {
            col = col * 26 + (ch.charCodeAt(0) - 64);
        }
        return col;
    }

    function escapeXml(str) {
        if (str == null) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    function isNumeric(v) {
        if (v === '' || v == null) return false;
        return !isNaN(v) && !isNaN(parseFloat(v));
    }

    // --- Template Analysis ---
    /**
     * Analyze a template XLSX file to extract its structure
     * @param {ArrayBuffer} buffer - The template file as ArrayBuffer
     * @returns {Object} Template data including zones, styles, and original ZIP
     */
    async function analyzeTemplate(buffer) {
        const zip = await JSZip.loadAsync(buffer);

        // Parse shared strings
        const sharedStrings = await parseSharedStrings(zip);

        // Parse workbook to get sheet names
        const wbXml = await zip.file('xl/workbook.xml').async('string');
        const wbDoc = new DOMParser().parseFromString(wbXml, 'application/xml');
        const sheetNodes = wbDoc.getElementsByTagName('sheet');
        const sheets = [];
        for (let i = 0; i < sheetNodes.length; i++) {
            const node = sheetNodes[i];
            sheets.push({
                name: node.getAttribute('name'),
                rId: node.getAttribute('r:id') ||
                    node.getAttributeNS(REL_NS, 'id') ||
                    `rId${i + 1}`
            });
        }

        // Parse workbook rels to map rId to file paths
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
        const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
        const relNodes = relsDoc.getElementsByTagName('Relationship');
        const relMap = {};
        for (let i = 0; i < relNodes.length; i++) {
            const rel = relNodes[i];
            let target = rel.getAttribute('Target');
            if (target.startsWith('/')) target = target.substring(1);
            if (!target.startsWith('xl/')) target = 'xl/' + target;
            relMap[rel.getAttribute('Id')] = target;
        }

        // Analyze first sheet as the reference template
        const firstSheetPath = relMap[sheets[0].rId];
        const analysis = await analyzeSheet(zip, firstSheetPath, sharedStrings);

        return {
            zip: zip,
            sheets: sheets,
            relMap: relMap,
            sharedStrings: sharedStrings,
            firstSheetPath: firstSheetPath,
            analysis: analysis,
            sheetNames: sheets.map(s => s.name),
        };
    }

    async function parseSharedStrings(zip) {
        const file = zip.file('xl/sharedStrings.xml');
        if (!file) return [];
        const xml = await file.async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const items = doc.getElementsByTagName('si');
        const strings = [];
        for (let i = 0; i < items.length; i++) {
            const tNodes = items[i].getElementsByTagName('t');
            let text = '';
            for (let j = 0; j < tNodes.length; j++) {
                text += tNodes[j].textContent || '';
            }
            strings.push(text);
        }
        return strings;
    }

    /**
     * Analyze a single sheet to detect header/data/footer zones
     */
    async function analyzeSheet(zip, sheetPath, sharedStrings) {
        const xml = await zip.file(sheetPath).async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');

        // Parse all rows
        const rowNodes = doc.getElementsByTagName('row');
        const rows = [];
        for (let i = 0; i < rowNodes.length; i++) {
            const rowNode = rowNodes[i];
            const rowNum = parseInt(rowNode.getAttribute('r'), 10);
            const cells = [];
            const cellNodes = rowNode.getElementsByTagName('c');
            for (let j = 0; j < cellNodes.length; j++) {
                const cell = cellNodes[j];
                const ref = cell.getAttribute('r');
                const s = cell.getAttribute('s') || '0';
                const t = cell.getAttribute('t') || '';
                const vEl = cell.getElementsByTagName('v')[0];
                const fEl = cell.getElementsByTagName('f')[0];
                let value = vEl ? vEl.textContent : '';
                let displayValue = value;

                // Resolve shared string
                if (t === 's' && sharedStrings[parseInt(value)]) {
                    displayValue = sharedStrings[parseInt(value)];
                }

                cells.push({
                    ref, s, t, value, displayValue,
                    col: refToCol(ref),
                    formula: fEl ? fEl.textContent : null,
                });
            }
            rows.push({
                rowNum,
                ht: rowNode.getAttribute('ht'),
                hidden: rowNode.getAttribute('hidden') === '1',
                customHeight: rowNode.getAttribute('customHeight'),
                cells
            });
        }

        // Detect column headers row
        // Strategy: find a row where multiple cells look like headers (short text, often bold font)
        // and there's a "No." or sequential numbering pattern starting after it
        let headerRowIdx = -1;
        let headerRow = null;

        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            if (r.cells.length >= 3) {
                // Check if this looks like a header row
                const texts = r.cells.map(c => c.displayValue);
                const looksLikeHeaders = texts.every(t => t.length < 30) && texts.filter(t => t.length > 0).length >= 3;
                // Check if next row has numbered first column
                if (looksLikeHeaders && i + 1 < rows.length) {
                    const nextRow = rows[i + 1];
                    const hasNextData = nextRow && nextRow.cells.length > 0;
                    // Check if row after this looks like data (has sequential numbers or values)  
                    if (hasNextData) {
                        // Look for typical header indicators
                        const hasNoColumn = texts.some(t =>
                            t === 'No.' || t === 'No' || t === 'STT' || t === '#' || t === '番号'
                        );
                        if (hasNoColumn) {
                            headerRowIdx = i;
                            headerRow = r;
                            break;
                        }
                    }
                }
            }
        }

        // Fallback: if no header found, look for the first row with many cells
        if (headerRowIdx === -1) {
            let maxCells = 0;
            for (let i = 0; i < Math.min(rows.length, 20); i++) {
                if (rows[i].cells.length > maxCells) {
                    maxCells = rows[i].cells.length;
                    headerRowIdx = i;
                    headerRow = rows[i];
                }
            }
        }

        // Everything before header = header zone (company info, title, etc.)
        const headerZoneRows = rows.slice(0, headerRowIdx);

        // Find data zone: rows after header row where first column has sequential numbers
        let dataStartIdx = headerRowIdx + 1;
        // Skip any sub-category rows (merged rows between header and first data)
        while (dataStartIdx < rows.length) {
            const r = rows[dataStartIdx];
            const firstCellVal = r.cells[0]?.value;
            // Check if it looks like a data row (first cell is a number)
            if (firstCellVal && !isNaN(parseInt(firstCellVal))) break;
            dataStartIdx++;
        }

        // Find where data ends (look for footer with formulas like SUM)
        let dataEndIdx = rows.length - 1;
        for (let i = dataStartIdx; i < rows.length; i++) {
            const r = rows[i];
            const hasFormula = r.cells.some(c => c.formula && c.formula.includes('SUM'));
            const hasTotalText = r.cells.some(c =>
                c.displayValue.includes('小計') || c.displayValue.includes('合計') ||
                c.displayValue.includes('Total') || c.displayValue.includes('Subtotal') ||
                c.displayValue.includes('Tổng') || c.displayValue.includes('消費税')
            );
            if (hasFormula || hasTotalText) {
                dataEndIdx = i - 1;
                break;
            }
        }

        // Handle the category rows within data (like "Reagent", "Calibrator & Control")
        // These are merged rows that separate data groups
        const categoryRows = [];
        const actualDataRows = [];
        for (let i = dataStartIdx; i <= dataEndIdx; i++) {
            const r = rows[i];
            const firstVal = r.cells[0]?.value;
            if (!firstVal || isNaN(parseInt(firstVal))) {
                // This might be a category/separator row
                if (r.cells.length >= 1 && r.cells[0].displayValue.length > 0) {
                    categoryRows.push({ index: i, row: r });
                }
            } else {
                actualDataRows.push({ index: i, row: r });
            }
        }

        // Footer zone
        const footerZoneRows = rows.slice(dataEndIdx + 1);

        // Extract column headers
        const columnHeaders = headerRow.cells.map(c => ({
            col: c.col,
            ref: c.ref,
            label: c.displayValue,
            style: c.s,
        }));

        // Extract style patterns for data rows
        const dataStylePatterns = [];
        for (const { row } of actualDataRows.slice(0, 4)) {
            const pattern = row.cells.map(c => ({
                col: c.col,
                style: c.s,
                type: c.t,
                hasFormula: !!c.formula,
                formulaPattern: c.formula,
            }));
            dataStylePatterns.push({ rowNum: row.rowNum, ht: row.ht, pattern });
        }

        // Parse merge cells
        const mergeNodes = doc.getElementsByTagName('mergeCell');
        const mergeCells = [];
        for (let i = 0; i < mergeNodes.length; i++) {
            mergeCells.push(mergeNodes[i].getAttribute('ref'));
        }

        // Parse columns
        const colNodes = doc.getElementsByTagName('col');
        const columns = [];
        for (let i = 0; i < colNodes.length; i++) {
            const c = colNodes[i];
            columns.push({
                min: c.getAttribute('min'),
                max: c.getAttribute('max'),
                width: c.getAttribute('width'),
                customWidth: c.getAttribute('customWidth'),
                style: c.getAttribute('style'),
            });
        }

        // Count max column from header
        const maxCol = Math.max(...headerRow.cells.map(c => c.col));

        return {
            headerZone: {
                rows: headerZoneRows,
                endRow: headerRow.rowNum - 1,
            },
            columnHeaderRow: {
                rowNum: headerRow.rowNum,
                row: headerRow,
                headers: columnHeaders,
            },
            dataZone: {
                startRowNum: rows[dataStartIdx]?.rowNum || headerRow.rowNum + 2,
                endRowNum: rows[dataEndIdx]?.rowNum || headerRow.rowNum + 10,
                startIdx: dataStartIdx,
                endIdx: dataEndIdx,
                dataRows: actualDataRows,
                categoryRows: categoryRows,
                stylePatterns: dataStylePatterns,
            },
            footerZone: {
                startIdx: dataEndIdx + 1,
                rows: footerZoneRows,
            },
            mergeCells,
            columns,
            maxCol,
            totalRows: rows.length,
            rawRows: rows,
        };
    }

    // --- Template-Based Generation ---
    /**
     * Generate a new XLSX file using template formatting and new data
     * @param {Object} templateData - Result from analyzeTemplate()
     * @param {Object} options - { rows, sheetName, headerOverrides, updateHeaderFields }
     * @returns {Blob} The generated XLSX file
     */
    async function generateFromTemplate(templateData, options) {
        const {
            rows: newDataRows,
            sheetName,
            headerFieldUpdates = {},
        } = options;

        const { zip, analysis, sharedStrings: origSharedStrings } = templateData;

        // Clone the template ZIP
        const newZipData = await zip.generateAsync({ type: 'arraybuffer' });
        const newZip = await JSZip.loadAsync(newZipData);

        // Build new shared strings set
        const newSharedStrings = [...origSharedStrings];
        const ssIndexMap = {};
        function getOrAddSharedString(text) {
            const key = String(text);
            if (ssIndexMap[key] !== undefined) return ssIndexMap[key];
            // Check if already exists
            const existingIdx = newSharedStrings.indexOf(key);
            if (existingIdx !== -1) {
                ssIndexMap[key] = existingIdx;
                return existingIdx;
            }
            const idx = newSharedStrings.length;
            newSharedStrings.push(key);
            ssIndexMap[key] = idx;
            return idx;
        }

        // Pre-index existing strings
        origSharedStrings.forEach((s, i) => { ssIndexMap[s] = i; });

        // --- Rebuild the worksheet XML ---
        const sheetPath = templateData.firstSheetPath;
        const origXml = await zip.file(sheetPath).async('string');
        const newSheetXml = rebuildSheet(origXml, analysis, newDataRows, getOrAddSharedString, headerFieldUpdates);

        newZip.file(sheetPath, newSheetXml);

        // --- Rebuild shared strings XML ---
        const newSSXml = buildSharedStringsXml(newSharedStrings);
        newZip.file('xl/sharedStrings.xml', newSSXml);

        // Update sheet name if provided
        if (sheetName) {
            const wbXml = await newZip.file('xl/workbook.xml').async('string');
            // Update first sheet name
            const updatedWb = wbXml.replace(
                /(<sheet[^>]*name=")([^"]*)(")/,
                `$1${escapeXml(sheetName)}$3`
            );
            newZip.file('xl/workbook.xml', updatedWb);
        }

        // Generate blob
        const blob = await newZip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });

        return blob;
    }

    /**
     * Rebuild sheet XML with new data rows
     */
    function rebuildSheet(origXml, analysis, newDataRows, getOrAddSS, headerFieldUpdates) {
        const doc = new DOMParser().parseFromString(origXml, 'application/xml');
        const { headerZone, columnHeaderRow, dataZone, footerZone, rawRows, maxCol } = analysis;

        // Determine style patterns to use for data rows
        const stylePatterns = dataZone.stylePatterns;
        if (stylePatterns.length === 0) return origXml; // Shouldn't happen

        // Get column count from header
        const colCount = columnHeaderRow.headers.length;

        // Build all row XMLs
        const allRows = [];

        // === 1. Header Zone (keep as-is, but apply field updates) ===
        for (const row of headerZone.rows) {
            let rowXml = buildRowXml(row, headerFieldUpdates, getOrAddSS);
            allRows.push(rowXml);
        }

        // === 2. Column Header Row (keep as-is) ===
        allRows.push(buildRowXml(columnHeaderRow.row, {}, getOrAddSS));

        // === 3. Data rows with category support ===
        const dataStartRowNum = dataZone.startRowNum;
        let currentRowNum = dataStartRowNum;

        // Check if there are category rows in the original — preserve them
        // But for simplicity, we'll just handle the data rows
        // If original had category rows, check if we need them
        const origCategoryRows = dataZone.categoryRows;

        // Determine if original uses R1/R2 pattern (2 rows per item)
        const hasR1R2 = stylePatterns.length >= 2;
        const patternCycleLen = hasR1R2 ? 2 : 1;

        let itemNum = 1;
        for (let i = 0; i < newDataRows.length; i++) {
            const dataRow = newDataRows[i];
            const patternIdx = i % patternCycleLen;
            const pattern = stylePatterns[Math.min(patternIdx, stylePatterns.length - 1)];

            const cells = [];
            for (let c = 0; c < colCount; c++) {
                const colNum = c + 1;
                const cellRef = colToRef(colNum) + currentRowNum;
                const patternCell = pattern.pattern.find(p => p.col === colNum);
                const styleIdx = patternCell ? patternCell.style : '0';

                const cellValue = (dataRow[c] !== undefined && dataRow[c] !== null) ? String(dataRow[c]) : '';

                if (colNum === 1 && isNumeric(cellValue)) {
                    // First column (No.) — always numeric centered
                    cells.push(`<c r="${cellRef}" s="${styleIdx}"><v>${cellValue}</v></c>`);
                } else if (isNumeric(cellValue)) {
                    // Numeric value
                    cells.push(`<c r="${cellRef}" s="${styleIdx}"><v>${cellValue}</v></c>`);
                } else if (cellValue === '') {
                    cells.push(`<c r="${cellRef}" s="${styleIdx}"/>`);
                } else {
                    // String value — use shared strings
                    const ssIdx = getOrAddSS(cellValue);
                    cells.push(`<c r="${cellRef}" s="${styleIdx}" t="s"><v>${ssIdx}</v></c>`);
                }
            }

            const ht = pattern.ht || '18';
            allRows.push(`<row r="${currentRowNum}" ht="${ht}" customHeight="1">${cells.join('')}</row>`);
            currentRowNum++;
        }

        // === 4. Footer Zone ===
        // Adjust footer row numbers and formula ranges
        const footerStartRowNum = currentRowNum;
        const origDataStart = dataZone.startRowNum;
        const origDataEnd = dataZone.endRowNum;
        const newDataEnd = currentRowNum - 1;

        for (const row of footerZone.rows) {
            const newRowNum = footerStartRowNum + (footerZone.rows.indexOf(row));
            const cells = [];
            for (const cell of row.cells) {
                const colNum = cell.col;
                const cellRef = colToRef(colNum) + newRowNum;
                let cellXml;

                if (cell.formula) {
                    // Update formula ranges
                    let updatedFormula = cell.formula;
                    // Replace row references in formulas like SUM(H17:H45,H47:H53)
                    updatedFormula = updateFormulaRanges(updatedFormula, origDataStart, origDataEnd, dataStartRowNum, newDataEnd);
                    // Also update self-referencing formulas
                    updatedFormula = updatedFormula.replace(
                        /H(\d+)/g,
                        (match, rn) => {
                            const origRn = parseInt(rn);
                            if (origRn >= footerZone.rows[0]?.rowNum) {
                                // Footer row ref — adjust
                                const offset = origRn - footerZone.rows[0].rowNum;
                                return `H${footerStartRowNum + offset}`;
                            }
                            return match;
                        }
                    );

                    if (cell.t === 's') {
                        const ssIdx = getOrAddSS(cell.displayValue);
                        cellXml = `<c r="${cellRef}" s="${cell.s}" t="s"><f>${escapeXml(updatedFormula)}</f><v>${ssIdx}</v></c>`;
                    } else {
                        cellXml = `<c r="${cellRef}" s="${cell.s}"><f>${escapeXml(updatedFormula)}</f><v>${cell.value}</v></c>`;
                    }
                } else if (cell.t === 's') {
                    const ssIdx = getOrAddSS(cell.displayValue);
                    cellXml = `<c r="${cellRef}" s="${cell.s}" t="s"><v>${ssIdx}</v></c>`;
                } else if (cell.value) {
                    cellXml = `<c r="${cellRef}" s="${cell.s}"><v>${cell.value}</v></c>`;
                } else {
                    cellXml = `<c r="${cellRef}" s="${cell.s}"/>`;
                }
                cells.push(cellXml);
            }

            const ht = row.ht || '18';
            allRows.push(`<row r="${newRowNum}" ht="${ht}" customHeight="1">${cells.join('')}</row>`);
        }

        // === Rebuild merge cells ===
        const newMerges = rebuildMergeCells(analysis, newDataRows.length, dataStartRowNum, newDataEnd, footerStartRowNum);

        // === Assemble the full worksheet XML ===
        // Extract everything we need from original
        const sheetDataContent = allRows.join('\n');

        // Rebuild the complete XML
        // We need to preserve: sheetViews, cols, pageMargins, pageSetup, drawing, etc.
        return buildFullSheetXml(doc, sheetDataContent, newMerges, currentRowNum + footerZone.rows.length - 1, maxCol);
    }

    function buildRowXml(row, fieldUpdates, getOrAddSS) {
        const cells = [];
        for (const cell of row.cells) {
            const cellRef = cell.ref;
            let cellXml;

            // Check if this cell should be updated
            const updateValue = fieldUpdates[cellRef];

            if (updateValue !== undefined) {
                if (isNumeric(updateValue)) {
                    cellXml = `<c r="${cellRef}" s="${cell.s}"><v>${updateValue}</v></c>`;
                } else {
                    const ssIdx = getOrAddSS(String(updateValue));
                    cellXml = `<c r="${cellRef}" s="${cell.s}" t="s"><v>${ssIdx}</v></c>`;
                }
            } else if (cell.formula) {
                if (cell.t === 's') {
                    cellXml = `<c r="${cellRef}" s="${cell.s}" t="s"><f>${escapeXml(cell.formula)}</f><v>${cell.value}</v></c>`;
                } else {
                    cellXml = `<c r="${cellRef}" s="${cell.s}"><f>${escapeXml(cell.formula)}</f><v>${cell.value}</v></c>`;
                }
            } else if (cell.t === 's') {
                cellXml = `<c r="${cellRef}" s="${cell.s}" t="s"><v>${cell.value}</v></c>`;
            } else if (cell.value) {
                cellXml = `<c r="${cellRef}" s="${cell.s}"><v>${cell.value}</v></c>`;
            } else {
                cellXml = `<c r="${cellRef}" s="${cell.s}"/>`;
            }
            cells.push(cellXml);
        }

        const attrs = [`r="${row.rowNum}"`];
        if (row.ht) attrs.push(`ht="${row.ht}"`);
        if (row.customHeight) attrs.push(`customHeight="1"`);
        if (row.hidden) attrs.push(`hidden="1"`);

        return `<row ${attrs.join(' ')}>${cells.join('')}</row>`;
    }

    function updateFormulaRanges(formula, origStart, origEnd, newStart, newEnd) {
        // Replace cell range references like H17:H45 with adjusted ranges
        return formula.replace(
            /([A-Z]+)(\d+):([A-Z]+)(\d+)/g,
            (match, col1, row1, col2, row2) => {
                const r1 = parseInt(row1);
                const r2 = parseInt(row2);
                // If the range falls within the original data area, adjust
                if (r1 >= origStart && r2 <= origEnd + 20) {
                    return `${col1}${newStart}:${col2}${newEnd}`;
                }
                return match;
            }
        );
    }

    function rebuildMergeCells(analysis, newDataCount, dataStart, dataEnd, footerStart) {
        const merges = [];
        const origMerges = analysis.mergeCells;
        const origDataStart = analysis.dataZone.startRowNum;
        const origDataEnd = analysis.dataZone.endRowNum;
        const origFooterStart = analysis.footerZone.rows[0]?.rowNum || origDataEnd + 1;

        for (const merge of origMerges) {
            const match = merge.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
            if (!match) continue;
            const [, col1, row1str, col2, row2str] = match;
            const r1 = parseInt(row1str);
            const r2 = parseInt(row2str);

            if (r1 < origDataStart) {
                // Header zone merge — keep as-is
                merges.push(merge);
            } else if (r1 >= origFooterStart) {
                // Footer zone merge — adjust row numbers
                const offset = r1 - origFooterStart;
                const offset2 = r2 - origFooterStart;
                merges.push(`${col1}${footerStart + offset}:${col2}${footerStart + offset2}`);
            }
            // Data zone merges are not copied — new data has its own structure
        }

        return merges;
    }

    function buildFullSheetXml(origDoc, sheetDataContent, merges, lastRow, maxCol) {
        const ns = XLSX_NS;

        // Extract key elements from original
        let sheetViewsXml = '';
        let colsXml = '';
        let pageMarginsXml = '';
        let pageSetupXml = '';
        let drawingXml = '';
        let headerFooterXml = '';

        // SheetViews
        const svNodes = origDoc.getElementsByTagName('sheetViews');
        if (svNodes.length > 0) {
            sheetViewsXml = new XMLSerializer().serializeToString(svNodes[0]);
            // Clean up namespace duplication
            sheetViewsXml = cleanNamespaces(sheetViewsXml);
        }

        // SheetFormatPr
        let sheetFormatPrXml = '';
        const sfpNodes = origDoc.getElementsByTagName('sheetFormatPr');
        if (sfpNodes.length > 0) {
            sheetFormatPrXml = new XMLSerializer().serializeToString(sfpNodes[0]);
            sheetFormatPrXml = cleanNamespaces(sheetFormatPrXml);
        }

        // Cols
        const colNodes = origDoc.getElementsByTagName('cols');
        if (colNodes.length > 0) {
            colsXml = new XMLSerializer().serializeToString(colNodes[0]);
            colsXml = cleanNamespaces(colsXml);
        }

        // Page margins
        const pmNodes = origDoc.getElementsByTagName('pageMargins');
        if (pmNodes.length > 0) {
            pageMarginsXml = new XMLSerializer().serializeToString(pmNodes[0]);
            pageMarginsXml = cleanNamespaces(pageMarginsXml);
        }

        // Page setup
        const psNodes = origDoc.getElementsByTagName('pageSetup');
        if (psNodes.length > 0) {
            pageSetupXml = new XMLSerializer().serializeToString(psNodes[0]);
            pageSetupXml = cleanNamespaces(pageSetupXml);
        }

        // Drawing reference
        const drawNodes = origDoc.getElementsByTagName('drawing');
        if (drawNodes.length > 0) {
            drawingXml = new XMLSerializer().serializeToString(drawNodes[0]);
            drawingXml = cleanNamespaces(drawingXml);
        }

        // Header/Footer
        const hfNodes = origDoc.getElementsByTagName('headerFooter');
        if (hfNodes.length > 0) {
            headerFooterXml = new XMLSerializer().serializeToString(hfNodes[0]);
            headerFooterXml = cleanNamespaces(headerFooterXml);
        }

        // LegacyDrawing (for comments)
        let legacyDrawingXml = '';
        const ldNodes = origDoc.getElementsByTagName('legacyDrawing');
        if (ldNodes.length > 0) {
            legacyDrawingXml = new XMLSerializer().serializeToString(ldNodes[0]);
            legacyDrawingXml = cleanNamespaces(legacyDrawingXml);
        }

        // Merge cells XML
        let mergeCellsXml = '';
        if (merges.length > 0) {
            mergeCellsXml = `<mergeCells count="${merges.length}">` +
                merges.map(m => `<mergeCell ref="${m}"/>`).join('') +
                `</mergeCells>`;
        }

        // Build dimension
        const lastCol = colToRef(maxCol || 8);
        const dimRef = `A1:${lastCol}${lastRow}`;

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${ns}" xmlns:r="${REL_NS}">
<dimension ref="${dimRef}"/>
${sheetViewsXml}
${sheetFormatPrXml}
${colsXml}
<sheetData>
${sheetDataContent}
</sheetData>
${mergeCellsXml}
${pageMarginsXml}
${pageSetupXml}
${headerFooterXml}
${drawingXml}
${legacyDrawingXml}
</worksheet>`;
    }

    function cleanNamespaces(xml) {
        // Remove duplicate namespace declarations added by XMLSerializer
        return xml
            .replace(/\s+xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/g, '')
            .replace(/\s+xmlns:r="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/g, '');
    }

    function buildSharedStringsXml(strings) {
        const count = strings.length;
        const items = strings.map(s => `<si><t>${escapeXml(s)}</t></si>`).join('\n');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${count}">
${items}
</sst>`;
    }

    /**
     * Get a summary of the template for UI display
     */
    function getTemplateSummary(templateData) {
        const a = templateData.analysis;
        return {
            sheetCount: templateData.sheets.length,
            sheetNames: templateData.sheetNames,
            columnHeaders: a.columnHeaderRow.headers.map(h => h.label),
            headerRowCount: a.headerZone.rows.length + 1, // +1 for column header row
            dataRowCount: a.dataZone.dataRows.length,
            footerRowCount: a.footerZone.rows.length,
            hasCategories: a.dataZone.categoryRows.length > 0,
            categoryCount: a.dataZone.categoryRows.length,
            maxColumns: a.maxCol,
            mergeCount: a.mergeCells.length,
        };
    }

    return {
        analyzeTemplate,
        generateFromTemplate,
        getTemplateSummary,
    };
})();
