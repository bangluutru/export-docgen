/**
 * Template Engine — Analyze template XLSX and generate new files using template formatting
 * Approach: Surgical DOM manipulation — only replace data zone rows, preserve everything else
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

    function refToRow(ref) {
        const match = ref.match(/(\d+)$/);
        return match ? parseInt(match[1], 10) : 1;
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

        // Parse shared strings — keep both text and raw XML
        const { strings: sharedStrings, rawSiElements } = await parseSharedStrings(zip);

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

        // Parse styles.xml for later PDF rendering use
        const stylesData = await parseStyles(zip);

        return {
            zip: zip,
            sheets: sheets,
            relMap: relMap,
            sharedStrings: sharedStrings,
            rawSiElements: rawSiElements,
            firstSheetPath: firstSheetPath,
            analysis: analysis,
            sheetNames: sheets.map(s => s.name),
            stylesData: stylesData,
        };
    }

    /**
     * Parse shared strings, preserving raw XML for rich text
     */
    async function parseSharedStrings(zip) {
        const file = zip.file('xl/sharedStrings.xml');
        if (!file) return { strings: [], rawSiElements: [] };

        const xml = await file.async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const items = doc.getElementsByTagName('si');
        const strings = [];
        const rawSiElements = [];
        const serializer = new XMLSerializer();

        for (let i = 0; i < items.length; i++) {
            // Extract display text
            const tNodes = items[i].getElementsByTagName('t');
            let text = '';
            for (let j = 0; j < tNodes.length; j++) {
                text += tNodes[j].textContent || '';
            }
            strings.push(text);
            // Keep raw XML for preserving rich text formatting
            rawSiElements.push(serializer.serializeToString(items[i]));
        }
        return { strings, rawSiElements };
    }

    /**
     * Parse xl/styles.xml for colors, fonts, fills, borders, number formats
     */
    async function parseStyles(zip) {
        const file = zip.file('xl/styles.xml');
        if (!file) return null;
        const xml = await file.async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');

        // Parse number formats
        const numFmts = {};
        const nfNodes = doc.getElementsByTagName('numFmt');
        for (let i = 0; i < nfNodes.length; i++) {
            const id = nfNodes[i].getAttribute('numFmtId');
            numFmts[id] = nfNodes[i].getAttribute('formatCode');
        }

        // Parse fonts
        const fonts = [];
        const fontsNode = doc.getElementsByTagName('fonts')[0];
        if (fontsNode) {
            const fontNodes = fontsNode.querySelectorAll(':scope > font');
            for (const fn of fontNodes) {
                const font = {};
                const sz = fn.getElementsByTagName('sz')[0];
                if (sz) font.size = parseFloat(sz.getAttribute('val'));
                const name = fn.getElementsByTagName('name')[0];
                if (name) font.name = name.getAttribute('val');
                const color = fn.getElementsByTagName('color')[0];
                if (color) {
                    font.colorRgb = color.getAttribute('rgb');
                    font.colorTheme = color.getAttribute('theme');
                }
                font.bold = fn.getElementsByTagName('b').length > 0;
                font.italic = fn.getElementsByTagName('i').length > 0;
                font.underline = fn.getElementsByTagName('u').length > 0;
                fonts.push(font);
            }
        }

        // Parse fills
        const fills = [];
        const fillsNode = doc.getElementsByTagName('fills')[0];
        if (fillsNode) {
            const fillNodes = fillsNode.querySelectorAll(':scope > fill');
            for (const fl of fillNodes) {
                const fill = {};
                const pf = fl.getElementsByTagName('patternFill')[0];
                if (pf) {
                    fill.pattern = pf.getAttribute('patternType');
                    const fgColor = pf.getElementsByTagName('fgColor')[0];
                    if (fgColor) {
                        fill.fgColorRgb = fgColor.getAttribute('rgb');
                        fill.fgColorTheme = fgColor.getAttribute('theme');
                    }
                }
                fills.push(fill);
            }
        }

        // Parse borders
        const borders = [];
        const bordersNode = doc.getElementsByTagName('borders')[0];
        if (bordersNode) {
            const borderNodes = bordersNode.querySelectorAll(':scope > border');
            for (const bn of borderNodes) {
                const border = {};
                for (const side of ['left', 'right', 'top', 'bottom']) {
                    const sideNode = bn.getElementsByTagName(side)[0];
                    if (sideNode && sideNode.getAttribute('style')) {
                        const colorNode = sideNode.getElementsByTagName('color')[0];
                        border[side] = {
                            style: sideNode.getAttribute('style'),
                            colorRgb: colorNode ? colorNode.getAttribute('rgb') : null,
                        };
                    }
                }
                borders.push(border);
            }
        }

        // Parse cell style xfs (cellXfs)
        const cellXfs = [];
        const xfsNode = doc.getElementsByTagName('cellXfs')[0];
        if (xfsNode) {
            const xfNodes = xfsNode.querySelectorAll(':scope > xf');
            for (const xf of xfNodes) {
                cellXfs.push({
                    numFmtId: xf.getAttribute('numFmtId') || '0',
                    fontId: parseInt(xf.getAttribute('fontId') || '0'),
                    fillId: parseInt(xf.getAttribute('fillId') || '0'),
                    borderId: parseInt(xf.getAttribute('borderId') || '0'),
                    alignment: (() => {
                        const al = xf.getElementsByTagName('alignment')[0];
                        if (!al) return null;
                        return {
                            horizontal: al.getAttribute('horizontal'),
                            vertical: al.getAttribute('vertical'),
                            wrapText: al.getAttribute('wrapText') === '1',
                        };
                    })(),
                });
            }
        }

        return { numFmts, fonts, fills, borders, cellXfs };
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
        let headerRowIdx = -1;
        let headerRow = null;

        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            if (r.cells.length >= 3) {
                const texts = r.cells.map(c => c.displayValue);
                const looksLikeHeaders = texts.every(t => t.length < 30) && texts.filter(t => t.length > 0).length >= 3;
                if (looksLikeHeaders && i + 1 < rows.length) {
                    const nextRow = rows[i + 1];
                    const hasNextData = nextRow && nextRow.cells.length > 0;
                    if (hasNextData) {
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

        // Fallback: first row with many cells
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

        // Header zone
        const headerZoneRows = rows.slice(0, headerRowIdx);

        // Find data zone
        let dataStartIdx = headerRowIdx + 1;
        while (dataStartIdx < rows.length) {
            const r = rows[dataStartIdx];
            const firstCellVal = r.cells[0]?.value;
            if (firstCellVal && !isNaN(parseInt(firstCellVal))) break;
            dataStartIdx++;
        }

        // Find where data ends
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

        // Categorize data rows
        const categoryRows = [];
        const actualDataRows = [];
        for (let i = dataStartIdx; i <= dataEndIdx; i++) {
            const r = rows[i];
            const firstVal = r.cells[0]?.value;
            if (!firstVal || isNaN(parseInt(firstVal))) {
                if (r.cells.length >= 1 && r.cells[0].displayValue.length > 0) {
                    categoryRows.push({ index: i, row: r });
                }
            } else {
                actualDataRows.push({ index: i, row: r });
            }
        }

        // Footer zone
        const footerZoneRows = rows.slice(dataEndIdx + 1);

        // Column headers
        const columnHeaders = headerRow.cells.map(c => ({
            col: c.col,
            ref: c.ref,
            label: c.displayValue,
            style: c.s,
        }));

        // Style patterns for data rows
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
     * Generate a new XLSX file using template formatting and new data.
     * Uses surgical DOM manipulation — only replaces data rows, preserves everything else.
     */
    async function generateFromTemplate(templateData, options) {
        const {
            rows: newDataRows,
            sheetName,
            headerFieldUpdates = {},
        } = options;

        const { zip, analysis, sharedStrings: origSharedStrings, rawSiElements } = templateData;

        // Clone the template ZIP
        const newZipData = await zip.generateAsync({ type: 'arraybuffer' });
        const newZip = await JSZip.loadAsync(newZipData);

        // Build new shared strings — preserve originals, add new ones
        const newSharedStrings = [...origSharedStrings];
        const newRawSiElements = [...rawSiElements];
        const ssIndexMap = {};

        function getOrAddSharedString(text) {
            const key = String(text);
            if (ssIndexMap[key] !== undefined) return ssIndexMap[key];
            const existingIdx = newSharedStrings.indexOf(key);
            if (existingIdx !== -1) {
                ssIndexMap[key] = existingIdx;
                return existingIdx;
            }
            const idx = newSharedStrings.length;
            newSharedStrings.push(key);
            // New strings are plain text (no rich text)
            newRawSiElements.push(`<si xmlns="${XLSX_NS}"><t>${escapeXml(key)}</t></si>`);
            ssIndexMap[key] = idx;
            return idx;
        }

        // Pre-index existing strings
        origSharedStrings.forEach((s, i) => { ssIndexMap[s] = i; });

        // --- Surgical rebuild of the worksheet XML ---
        const sheetPath = templateData.firstSheetPath;
        const origXml = await zip.file(sheetPath).async('string');
        const newSheetXml = rebuildSheetSurgical(origXml, analysis, newDataRows, getOrAddSharedString, headerFieldUpdates);

        newZip.file(sheetPath, newSheetXml);

        // --- Rebuild shared strings XML preserving rich text ---
        const newSSXml = buildSharedStringsXmlPreserved(newRawSiElements);
        newZip.file('xl/sharedStrings.xml', newSSXml);

        // Update sheet name if provided
        if (sheetName) {
            const wbXml = await newZip.file('xl/workbook.xml').async('string');
            const updatedWb = wbXml.replace(
                /(<sheet[^>]*name=")([^"]*)(")/,
                `$1${escapeXml(sheetName)}$3`
            );
            newZip.file('xl/workbook.xml', updatedWb);
        }

        // Generate blob WITH compression
        const blob = await newZip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
        });

        return blob;
    }

    /**
     * Surgical rebuild — Modify DOM directly, only replace data zone rows.
     * Preserves ALL XML elements: conditionalFormatting, dataValidations,
     * printOptions, sheetPr, autoFilter, hyperlinks, drawings, comments, etc.
     */
    function rebuildSheetSurgical(origXml, analysis, newDataRows, getOrAddSS, headerFieldUpdates) {
        const doc = new DOMParser().parseFromString(origXml, 'application/xml');
        const sheetData = doc.getElementsByTagName('sheetData')[0];
        if (!sheetData) return origXml; // Safety fallback

        const { columnHeaderRow, dataZone, footerZone, maxCol } = analysis;
        const stylePatterns = dataZone.stylePatterns;
        if (stylePatterns.length === 0) return origXml;

        const colCount = columnHeaderRow.headers.length;
        const dataStartRowNum = dataZone.startRowNum;
        const origDataEndRowNum = dataZone.endRowNum;
        const origFooterStartRowNum = footerZone.rows[0]?.rowNum || origDataEndRowNum + 1;

        // === Step 1: Apply header field updates (if any) ===
        if (Object.keys(headerFieldUpdates).length > 0) {
            const allRows = sheetData.getElementsByTagName('row');
            for (let i = 0; i < allRows.length; i++) {
                const rowNum = parseInt(allRows[i].getAttribute('r'), 10);
                if (rowNum >= dataStartRowNum) break; // Only update header zone
                const cells = allRows[i].getElementsByTagName('c');
                for (let j = 0; j < cells.length; j++) {
                    const cellRef = cells[j].getAttribute('r');
                    if (headerFieldUpdates[cellRef] !== undefined) {
                        updateCellValue(doc, cells[j], headerFieldUpdates[cellRef], getOrAddSS);
                    }
                }
            }
        }

        // === Step 2: Collect and remove data zone rows ===
        const rowsToRemove = [];
        const allRows = sheetData.getElementsByTagName('row');
        for (let i = allRows.length - 1; i >= 0; i--) {
            const rowNum = parseInt(allRows[i].getAttribute('r'), 10);
            if (rowNum >= dataStartRowNum && rowNum <= origDataEndRowNum) {
                rowsToRemove.push(allRows[i]);
            }
        }
        rowsToRemove.forEach(row => sheetData.removeChild(row));

        // === Step 3: Build new data rows ===
        const hasR1R2 = stylePatterns.length >= 2;
        const patternCycleLen = hasR1R2 ? 2 : 1;
        let currentRowNum = dataStartRowNum;
        const newDataRowNodes = [];

        for (let i = 0; i < newDataRows.length; i++) {
            const dataRow = newDataRows[i];
            const patternIdx = i % patternCycleLen;
            const pattern = stylePatterns[Math.min(patternIdx, stylePatterns.length - 1)];

            const rowEl = doc.createElementNS(XLSX_NS, 'row');
            rowEl.setAttribute('r', String(currentRowNum));
            if (pattern.ht) {
                rowEl.setAttribute('ht', pattern.ht);
                rowEl.setAttribute('customHeight', '1');
            }

            for (let c = 0; c < colCount; c++) {
                const colNum = c + 1;
                const cellRef = colToRef(colNum) + currentRowNum;
                const patternCell = pattern.pattern.find(p => p.col === colNum);
                const styleIdx = patternCell ? patternCell.style : '0';
                const cellValue = (dataRow[c] !== undefined && dataRow[c] !== null) ? String(dataRow[c]) : '';

                const cellEl = doc.createElementNS(XLSX_NS, 'c');
                cellEl.setAttribute('r', cellRef);
                cellEl.setAttribute('s', styleIdx);

                if (cellValue === '') {
                    // Empty cell — just has style
                } else if (isNumeric(cellValue)) {
                    const vEl = doc.createElementNS(XLSX_NS, 'v');
                    vEl.textContent = cellValue;
                    cellEl.appendChild(vEl);
                } else {
                    const ssIdx = getOrAddSS(cellValue);
                    cellEl.setAttribute('t', 's');
                    const vEl = doc.createElementNS(XLSX_NS, 'v');
                    vEl.textContent = String(ssIdx);
                    cellEl.appendChild(vEl);
                }

                rowEl.appendChild(cellEl);
            }

            newDataRowNodes.push(rowEl);
            currentRowNum++;
        }

        // === Step 4: Find the insertion point (first row after data zone = footer) ===
        const newDataEnd = currentRowNum - 1;
        const rowShift = newDataEnd - origDataEndRowNum; // How many rows shifted

        // Find the first footer row in DOM to insert data before it
        let insertBeforeNode = null;
        const existingRows = sheetData.getElementsByTagName('row');
        for (let i = 0; i < existingRows.length; i++) {
            const rn = parseInt(existingRows[i].getAttribute('r'), 10);
            if (rn >= origFooterStartRowNum) {
                insertBeforeNode = existingRows[i];
                break;
            }
        }

        // Insert new data rows
        for (const rowNode of newDataRowNodes) {
            if (insertBeforeNode) {
                sheetData.insertBefore(rowNode, insertBeforeNode);
            } else {
                sheetData.appendChild(rowNode);
            }
        }

        // === Step 5: Adjust footer row numbers ===
        if (rowShift !== 0) {
            // Re-fetch rows after insertion
            const updatedRows = sheetData.getElementsByTagName('row');
            for (let i = 0; i < updatedRows.length; i++) {
                const rn = parseInt(updatedRows[i].getAttribute('r'), 10);
                if (rn >= origFooterStartRowNum) {
                    const newRn = rn + rowShift;
                    updatedRows[i].setAttribute('r', String(newRn));
                    // Update all cell refs in this row
                    const cells = updatedRows[i].getElementsByTagName('c');
                    for (let j = 0; j < cells.length; j++) {
                        const oldRef = cells[j].getAttribute('r');
                        const colPart = oldRef.match(/^([A-Z]+)/)[1];
                        cells[j].setAttribute('r', colPart + newRn);

                        // Update formulas in footer
                        const fEl = cells[j].getElementsByTagName('f')[0];
                        if (fEl && fEl.textContent) {
                            fEl.textContent = updateFormulaRangesGeneric(
                                fEl.textContent,
                                dataStartRowNum, origDataEndRowNum,
                                dataStartRowNum, newDataEnd,
                                origFooterStartRowNum, origFooterStartRowNum + rowShift
                            );
                        }
                    }
                }
            }
        }

        // === Step 6: Update merge cells ===
        updateMergeCells(doc, dataStartRowNum, origDataEndRowNum, newDataEnd, origFooterStartRowNum, rowShift);

        // === Step 7: Update dimension ref ===
        const dimNode = doc.getElementsByTagName('dimension')[0];
        if (dimNode) {
            const lastRow = (footerZone.rows.length > 0)
                ? footerZone.rows[footerZone.rows.length - 1].rowNum + rowShift
                : newDataEnd;
            const lastCol = colToRef(maxCol || 8);
            dimNode.setAttribute('ref', `A1:${lastCol}${lastRow}`);
        }

        // === Step 8: Update autoFilter range if it exists ===
        const autoFilterNodes = doc.getElementsByTagName('autoFilter');
        if (autoFilterNodes.length > 0) {
            const af = autoFilterNodes[0];
            const afRef = af.getAttribute('ref');
            if (afRef) {
                const match = afRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
                if (match) {
                    const [, c1, r1, c2, r2] = match;
                    const origR2 = parseInt(r2);
                    if (origR2 >= origDataEndRowNum) {
                        af.setAttribute('ref', `${c1}${r1}:${c2}${origR2 + rowShift}`);
                    }
                }
            }
        }

        // === Step 9: Update conditional formatting ranges ===
        const cfNodes = doc.getElementsByTagName('conditionalFormatting');
        for (let i = 0; i < cfNodes.length; i++) {
            const sqref = cfNodes[i].getAttribute('sqref');
            if (sqref) {
                cfNodes[i].setAttribute('sqref', adjustRangeRef(sqref, origDataEndRowNum, rowShift));
            }
        }

        // Serialize back
        const serializer = new XMLSerializer();
        let output = serializer.serializeToString(doc);

        // Clean up extra namespace declarations that XMLSerializer adds
        output = output.replace(/\s+xmlns:ns\d+="[^"]*"/g, '');
        output = output.replace(/ns\d+:/g, '');

        return output;
    }

    /**
     * Update a cell's value in-place within the DOM
     */
    function updateCellValue(doc, cellEl, newValue, getOrAddSS) {
        // Remove existing v element
        const existingV = cellEl.getElementsByTagName('v')[0];

        if (isNumeric(newValue)) {
            cellEl.removeAttribute('t');
            if (existingV) {
                existingV.textContent = String(newValue);
            } else {
                const vEl = doc.createElementNS(XLSX_NS, 'v');
                vEl.textContent = String(newValue);
                cellEl.appendChild(vEl);
            }
        } else {
            const ssIdx = getOrAddSS(String(newValue));
            cellEl.setAttribute('t', 's');
            if (existingV) {
                existingV.textContent = String(ssIdx);
            } else {
                const vEl = doc.createElementNS(XLSX_NS, 'v');
                vEl.textContent = String(ssIdx);
                cellEl.appendChild(vEl);
            }
        }
    }

    /**
     * Generic formula range updater — handles all columns, not just H
     */
    function updateFormulaRangesGeneric(formula, origDataStart, origDataEnd, newDataStart, newDataEnd, origFooterStart, newFooterStart) {
        // Update cell range references like SUM(H17:H45) → SUM(H17:H30)
        let result = formula.replace(
            /([A-Z]+)(\d+):([A-Z]+)(\d+)/g,
            (match, col1, row1, col2, row2) => {
                const r1 = parseInt(row1);
                const r2 = parseInt(row2);
                // If range falls within or covers the original data area
                if (r1 >= origDataStart && r1 <= origDataEnd + 20 && r2 >= origDataStart) {
                    const newR1 = (r1 >= origFooterStart) ? newFooterStart + (r1 - origFooterStart) : r1;
                    const newR2 = (r2 >= origFooterStart) ? newFooterStart + (r2 - origFooterStart)
                        : (r2 <= origDataEnd) ? newDataEnd : newDataEnd;
                    return `${col1}${newR1}:${col2}${newR2}`;
                }
                return match;
            }
        );

        // Update individual cell references in footer rows
        result = result.replace(
            /([A-Z]+)(\d+)(?![:\d])/g,
            (match, col, row) => {
                const r = parseInt(row);
                if (r >= origFooterStart) {
                    return `${col}${newFooterStart + (r - origFooterStart)}`;
                }
                return match;
            }
        );

        return result;
    }

    /**
     * Update merge cells in the DOM — adjust for data zone changes
     */
    function updateMergeCells(doc, dataStart, origDataEnd, newDataEnd, origFooterStart, rowShift) {
        const mergeCellsNode = doc.getElementsByTagName('mergeCells')[0];
        if (!mergeCellsNode) return;

        const mergeNodes = mergeCellsNode.getElementsByTagName('mergeCell');
        const toRemove = [];
        const toAdd = [];

        for (let i = mergeNodes.length - 1; i >= 0; i--) {
            const ref = mergeNodes[i].getAttribute('ref');
            const match = ref.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
            if (!match) continue;

            const [, col1, row1str, col2, row2str] = match;
            const r1 = parseInt(row1str);
            const r2 = parseInt(row2str);

            if (r1 >= dataStart && r2 <= origDataEnd) {
                // Data zone merge — remove (new data has its own structure)
                toRemove.push(mergeNodes[i]);
            } else if (r1 >= origFooterStart) {
                // Footer zone merge — adjust row numbers
                mergeNodes[i].setAttribute('ref',
                    `${col1}${r1 + rowShift}:${col2}${r2 + rowShift}`
                );
            } else if (r1 < dataStart && r2 >= origFooterStart) {
                // Cross-zone merge — adjust end row
                mergeNodes[i].setAttribute('ref',
                    `${col1}${r1}:${col2}${r2 + rowShift}`
                );
            }
        }

        toRemove.forEach(node => mergeCellsNode.removeChild(node));

        // Update count
        const remaining = mergeCellsNode.getElementsByTagName('mergeCell').length;
        if (remaining === 0) {
            mergeCellsNode.parentNode.removeChild(mergeCellsNode);
        } else {
            mergeCellsNode.setAttribute('count', String(remaining));
        }
    }

    /**
     * Adjust a sqref string (e.g. "A5:H50") when rows shift
     */
    function adjustRangeRef(sqref, origDataEnd, rowShift) {
        return sqref.replace(
            /([A-Z]+)(\d+):([A-Z]+)(\d+)/g,
            (match, col1, row1, col2, row2) => {
                const r2 = parseInt(row2);
                if (r2 >= origDataEnd) {
                    return `${col1}${row1}:${col2}${r2 + rowShift}`;
                }
                return match;
            }
        );
    }

    /**
     * Build shared strings XML preserving original rich text formatting
     */
    function buildSharedStringsXmlPreserved(rawSiElements) {
        const count = rawSiElements.length;
        // Clean up namespace attributes from serialized elements
        const cleanedElements = rawSiElements.map(si => {
            return si
                .replace(/\s+xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/g, '')
                .replace(/\s+xmlns:ns\d+="[^"]*"/g, '');
        });
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${count}">
${cleanedElements.join('\n')}
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
            headerRowCount: a.headerZone.rows.length + 1,
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
