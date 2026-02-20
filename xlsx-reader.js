// ============================================
// Custom XLSX Reader Engine
// Reads .xlsx files by parsing OpenXML directly
// Uses only JSZip (generic ZIP library) + DOMParser
// ============================================

const XLSXReader = (function () {
    'use strict';

    /**
     * Read an XLSX file (ArrayBuffer) and return structured data
     * @param {ArrayBuffer} buffer - The file content
     * @returns {Promise<{sheetNames: string[], sheets: Object}>}
     */
    async function read(buffer) {
        const zip = await JSZip.loadAsync(buffer);

        // 1. Parse shared strings
        const sharedStrings = await parseSharedStrings(zip);

        // 2. Parse workbook to get sheet names and their rIds
        const { sheets: sheetInfoList } = await parseWorkbook(zip);

        // 3. Parse workbook relationships to map rId -> file path
        const rels = await parseWorkbookRels(zip);

        // 4. Parse each sheet
        const result = { sheetNames: [], sheets: {} };

        for (const sheetInfo of sheetInfoList) {
            const filePath = rels[sheetInfo.rId];
            if (!filePath) continue;

            const sheetData = await parseSheet(zip, filePath, sharedStrings);
            result.sheetNames.push(sheetInfo.name);
            result.sheets[sheetInfo.name] = sheetData;
        }

        if (result.sheetNames.length === 0) {
            throw new Error('Không tìm thấy sheet nào trong file.');
        }

        return result;
    }

    /**
     * Parse xl/sharedStrings.xml
     */
    async function parseSharedStrings(zip) {
        const strings = [];
        const file = zip.file('xl/sharedStrings.xml');
        if (!file) return strings;

        const xml = await file.async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const siNodes = doc.getElementsByTagName('si');

        for (let i = 0; i < siNodes.length; i++) {
            // A shared string can have multiple <t> elements (rich text)
            const tNodes = siNodes[i].getElementsByTagName('t');
            let text = '';
            for (let j = 0; j < tNodes.length; j++) {
                text += tNodes[j].textContent || '';
            }
            strings.push(text);
        }

        return strings;
    }

    /**
     * Parse xl/workbook.xml to get sheet names and rIds
     */
    async function parseWorkbook(zip) {
        const file = zip.file('xl/workbook.xml');
        if (!file) throw new Error('Không tìm thấy workbook.xml');

        const xml = await file.async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');

        const sheets = [];
        const sheetNodes = doc.getElementsByTagName('sheet');
        for (let i = 0; i < sheetNodes.length; i++) {
            const node = sheetNodes[i];
            const name = node.getAttribute('name') || `Sheet${i + 1}`;
            // rId can be r:id or just id depending on namespace
            const rId = node.getAttribute('r:id') ||
                node.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ||
                `rId${i + 1}`;
            sheets.push({ name, rId });
        }

        return { sheets };
    }

    /**
     * Parse xl/_rels/workbook.xml.rels
     */
    async function parseWorkbookRels(zip) {
        const rels = {};
        const file = zip.file('xl/_rels/workbook.xml.rels');
        if (!file) return rels;

        const xml = await file.async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const relNodes = doc.getElementsByTagName('Relationship');

        for (let i = 0; i < relNodes.length; i++) {
            const node = relNodes[i];
            const id = node.getAttribute('Id');
            let target = node.getAttribute('Target');
            // Make path relative to xl/
            if (target && !target.startsWith('/')) {
                target = 'xl/' + target;
            } else if (target && target.startsWith('/')) {
                target = target.substring(1);
            }
            rels[id] = target;
        }

        return rels;
    }

    /**
     * Parse a worksheet XML file
     */
    async function parseSheet(zip, filePath, sharedStrings) {
        const file = zip.file(filePath);
        if (!file) throw new Error(`Không tìm thấy sheet: ${filePath}`);

        const xml = await file.async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml');

        // Get all rows
        const rowNodes = doc.getElementsByTagName('row');
        const rawRows = {}; // rowNum -> { colNum -> value }
        let maxCol = 0;

        for (let i = 0; i < rowNodes.length; i++) {
            const rowNode = rowNodes[i];
            const rowNum = parseInt(rowNode.getAttribute('r'), 10);
            const cells = rowNode.getElementsByTagName('c');
            const rowData = {};

            for (let j = 0; j < cells.length; j++) {
                const cell = cells[j];
                const ref = cell.getAttribute('r'); // e.g., "A1"
                const colNum = colRefToNum(ref);
                if (colNum > maxCol) maxCol = colNum;

                const value = getCellValue(cell, sharedStrings);
                rowData[colNum] = value;
            }

            rawRows[rowNum] = rowData;
        }

        // Convert to arrays
        // First row = headers
        const rowNums = Object.keys(rawRows).map(Number).sort((a, b) => a - b);
        if (rowNums.length === 0) {
            return { headers: ['Cột 1'], rows: [] };
        }

        const firstRowNum = rowNums[0];
        const headers = [];
        for (let c = 1; c <= maxCol; c++) {
            const val = rawRows[firstRowNum] ? (rawRows[firstRowNum][c] || '') : '';
            headers.push(val || `Cột ${c}`);
        }

        const rows = [];
        for (let ri = 1; ri < rowNums.length; ri++) {
            const rn = rowNums[ri];
            const row = [];
            for (let c = 1; c <= maxCol; c++) {
                row.push(rawRows[rn] ? (rawRows[rn][c] || '') : '');
            }
            rows.push(row);
        }

        return { headers, rows };
    }

    /**
     * Extract cell value from <c> element
     */
    function getCellValue(cellNode, sharedStrings) {
        const type = cellNode.getAttribute('t'); // s=shared string, n=number, b=boolean, etc.
        const vNode = cellNode.getElementsByTagName('v')[0];
        const isNode = cellNode.getElementsByTagName('is')[0]; // inline string

        if (isNode) {
            // Inline string
            const tNodes = isNode.getElementsByTagName('t');
            let text = '';
            for (let i = 0; i < tNodes.length; i++) {
                text += tNodes[i].textContent || '';
            }
            return text;
        }

        if (!vNode) return '';

        const rawValue = vNode.textContent || '';

        switch (type) {
            case 's': // Shared string
                const idx = parseInt(rawValue, 10);
                return (idx >= 0 && idx < sharedStrings.length) ? sharedStrings[idx] : rawValue;

            case 'b': // Boolean
                return rawValue === '1' ? 'TRUE' : 'FALSE';

            case 'e': // Error
                return rawValue;

            case 'str': // Formula string
                return rawValue;

            default: // Number or date
                return rawValue;
        }
    }

    /**
     * Convert cell reference like "AB12" to column number (1-indexed)
     * A=1, B=2, ..., Z=26, AA=27, etc.
     */
    function colRefToNum(ref) {
        const match = ref.match(/^([A-Z]+)/i);
        if (!match) return 1;
        const letters = match[1].toUpperCase();
        let num = 0;
        for (let i = 0; i < letters.length; i++) {
            num = num * 26 + (letters.charCodeAt(i) - 64);
        }
        return num;
    }

    /**
     * Read a CSV file
     */
    function readCSV(text) {
        const lines = text.split(/\r?\n/).filter(l => l.trim());
        if (lines.length === 0) throw new Error('File CSV rỗng.');

        // Detect delimiter
        const first = lines[0];
        const commas = (first.match(/,/g) || []).length;
        const tabs = (first.match(/\t/g) || []).length;
        const semis = (first.match(/;/g) || []).length;
        const delim = tabs > commas ? '\t' : semis > commas ? ';' : ',';

        const parseLine = (line) => {
            const result = [];
            let current = '';
            let inQuotes = false;
            for (let i = 0; i < line.length; i++) {
                const ch = line[i];
                if (ch === '"') {
                    if (inQuotes && line[i + 1] === '"') {
                        current += '"';
                        i++;
                    } else {
                        inQuotes = !inQuotes;
                    }
                } else if (ch === delim && !inQuotes) {
                    result.push(current.trim());
                    current = '';
                } else {
                    current += ch;
                }
            }
            result.push(current.trim());
            return result;
        };

        const headers = parseLine(lines[0]);
        const rows = lines.slice(1).map(l => {
            const p = parseLine(l);
            while (p.length < headers.length) p.push('');
            return p.slice(0, headers.length);
        });

        return {
            sheetNames: ['Sheet1'],
            sheets: { Sheet1: { headers, rows } },
        };
    }

    return { read, readCSV };
})();
