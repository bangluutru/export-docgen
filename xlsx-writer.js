// ============================================
// Custom XLSX Writer Engine
// Generates .xlsx files from scratch using OpenXML
// Uses only JSZip (generic ZIP) — NO Excel libraries
// ============================================

const XLSXWriter = (function () {
    'use strict';

    // ---- Template color definitions ----
    const TEMPLATES = {
        professional: {
            headerFill: '1E3A5F',
            headerFont: 'FFFFFF',
            evenFill: 'EEF2F7',
            oddFill: 'FFFFFF',
            borderColor: 'D0D5DD',
            titleColor: '1E3A5F',
            pdfHeadStyles: { fillColor: [30, 58, 95], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 10 },
            pdfAlternateRowStyles: { fillColor: [238, 242, 247] },
            pdfBodyStyles: { textColor: [51, 51, 51], fontSize: 9 },
            pdfTitleColor: [30, 58, 95],
        },
        modern: {
            headerFill: '6366F1',
            headerFont: 'FFFFFF',
            evenFill: 'F0EDFF',
            oddFill: 'FFFFFF',
            borderColor: 'E0DEFF',
            titleColor: '6366F1',
            pdfHeadStyles: { fillColor: [99, 102, 241], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 10 },
            pdfAlternateRowStyles: { fillColor: [240, 237, 255] },
            pdfBodyStyles: { textColor: [51, 51, 51], fontSize: 9 },
            pdfTitleColor: [99, 102, 241],
        },
        classic: {
            headerFill: '2D5016',
            headerFont: 'FFFFFF',
            evenFill: 'ECF5E8',
            oddFill: 'FFFFFF',
            borderColor: 'C8E6C9',
            titleColor: '2D5016',
            pdfHeadStyles: { fillColor: [45, 80, 22], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 10 },
            pdfAlternateRowStyles: { fillColor: [236, 245, 232] },
            pdfBodyStyles: { textColor: [51, 51, 51], fontSize: 9 },
            pdfTitleColor: [45, 80, 22],
        },
        minimal: {
            headerFill: '374151',
            headerFont: 'FFFFFF',
            evenFill: 'F3F4F6',
            oddFill: 'FFFFFF',
            borderColor: 'E5E7EB',
            titleColor: '374151',
            pdfHeadStyles: { fillColor: [55, 65, 81], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 10 },
            pdfAlternateRowStyles: { fillColor: [243, 244, 246] },
            pdfBodyStyles: { textColor: [55, 65, 81], fontSize: 9 },
            pdfTitleColor: [55, 65, 81],
        },
    };

    /**
     * Generate an XLSX file from data
     * @param {Object} options
     * @param {string[]} options.headers
     * @param {Array<Array<string|number>>} options.rows
     * @param {string} options.title
     * @param {string} options.templateName
     * @param {boolean} options.addSTT
     * @param {boolean} options.addDate
     * @param {boolean} options.autofit
     * @param {string} options.sheetName
     * @returns {Promise<Blob>}
     */
    async function generate(options) {
        const {
            headers: rawHeaders,
            rows: rawRows,
            title = 'BÁO CÁO DỮ LIỆU',
            templateName = 'professional',
            addSTT = true,
            addDate = true,
            autofit = true,
            sheetName = 'Sheet1',
        } = options;

        const tpl = TEMPLATES[templateName] || TEMPLATES.professional;
        const headers = addSTT ? ['STT', ...rawHeaders] : [...rawHeaders];
        const rows = rawRows.map((row, idx) => addSTT ? [idx + 1, ...row] : [...row]);
        const totalCols = headers.length;

        // Calculate column widths
        const colWidths = headers.map((h, i) => {
            let maxLen = String(h).length;
            rows.forEach(row => {
                const len = String(row[i] || '').length;
                if (len > maxLen) maxLen = len;
            });
            return autofit ? Math.min(Math.max(maxLen + 3, 10), 50) : 15;
        });

        // Build shared strings
        const sharedStrings = [];
        const ssMap = {};

        function addSharedString(str) {
            const s = String(str);
            if (ssMap[s] !== undefined) return ssMap[s];
            const idx = sharedStrings.length;
            sharedStrings.push(s);
            ssMap[s] = idx;
            return idx;
        }

        // Pre-register title and date strings
        const titleSSIdx = addSharedString(title);
        const dateStr = formatDateVN();
        const dateSSIdx = addDate ? addSharedString(dateStr) : -1;

        // Pre-register all header and data strings
        const headerSSIdxs = headers.map(h => addSharedString(String(h)));
        const dataSSIdxs = rows.map(row =>
            row.map(val => {
                const s = String(val);
                // Check if numeric
                if (s !== '' && !isNaN(parseFloat(s.replace(/,/g, ''))) && isFinite(s.replace(/,/g, ''))) {
                    return { type: 'n', value: parseFloat(s.replace(/,/g, '')), ssIdx: -1 };
                }
                return { type: 's', value: s, ssIdx: addSharedString(s) };
            })
        );

        // Build XML parts
        const zip = new JSZip();

        // [Content_Types].xml
        zip.file('[Content_Types].xml', buildContentTypes());

        // _rels/.rels
        zip.file('_rels/.rels', buildRootRels());

        // xl/workbook.xml
        zip.file('xl/workbook.xml', buildWorkbook(sheetName));

        // xl/_rels/workbook.xml.rels
        zip.file('xl/_rels/workbook.xml.rels', buildWorkbookRels());

        // xl/styles.xml — the formatting heart
        zip.file('xl/styles.xml', buildStyles(tpl));

        // xl/sharedStrings.xml
        zip.file('xl/sharedStrings.xml', buildSharedStrings(sharedStrings));

        // xl/worksheets/sheet1.xml — the data
        zip.file('xl/worksheets/sheet1.xml', buildSheet({
            headers,
            rows,
            headerSSIdxs,
            dataSSIdxs,
            titleSSIdx,
            dateSSIdx,
            addDate,
            totalCols,
            colWidths,
        }));

        const blob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
        });

        return blob;
    }

    // ---- XML Builders ----

    function buildContentTypes() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;
    }

    function buildRootRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
    }

    function buildWorkbook(sheetName) {
        const escapedName = escapeXml(sheetName);
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${escapedName}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
    }

    function buildWorkbookRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;
    }

    /**
     * Build styles.xml with formatted cell styles
     * 
     * Style indices (xf):
     * 0 = default
     * 1 = title (bold, large, colored, center)
     * 2 = date subtitle (italic, gray, center)
     * 3 = header (bold, white text, colored fill, center, border)
     * 4 = data even row (fill, border)
     * 5 = data odd row (border)
     * 6 = data even row - number (right-aligned)
     * 7 = data odd row - number (right-aligned)
     * 8 = data even row - center (for STT)
     * 9 = data odd row - center (for STT)
     */
    function buildStyles(tpl) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.##"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="10"/><name val="Arial"/></font>
    <font><b/><sz val="16"/><color rgb="FF${tpl.titleColor}"/><name val="Arial"/></font>
    <font><i/><sz val="10"/><color rgb="FF666666"/><name val="Arial"/></font>
    <font><b/><sz val="11"/><color rgb="FF${tpl.headerFont}"/><name val="Arial"/></font>
    <font><sz val="10"/><color rgb="FF333333"/><name val="Arial"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF${tpl.headerFill}"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF${tpl.evenFill}"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF${tpl.borderColor}"/></left>
      <right style="thin"><color rgb="FF${tpl.borderColor}"/></right>
      <top style="thin"><color rgb="FF${tpl.borderColor}"/></top>
      <bottom style="thin"><color rgb="FF${tpl.borderColor}"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="164" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="164" fontId="4" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>`;
    }

    function buildSharedStrings(strings) {
        const items = strings.map(s => `<si><t>${escapeXml(s)}</t></si>`).join('\n  ');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">
  ${items}
</sst>`;
    }

    function buildSheet(opts) {
        const { headers, rows, headerSSIdxs, dataSSIdxs, titleSSIdx, dateSSIdx, addDate, totalCols, colWidths } = opts;

        // Column definitions
        const colDefs = colWidths.map((w, i) =>
            `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`
        ).join('\n      ');

        const lastColLetter = numToColRef(totalCols);
        let xmlRows = '';
        let rowNum = 1;

        // Title row (merged)
        xmlRows += `    <row r="${rowNum}" ht="30" customHeight="1">
      <c r="A${rowNum}" s="1" t="s"><v>${titleSSIdx}</v></c>
    </row>\n`;
        rowNum++;

        // Date row (merged) if needed
        if (addDate) {
            xmlRows += `    <row r="${rowNum}" ht="20" customHeight="1">
      <c r="A${rowNum}" s="2" t="s"><v>${dateSSIdx}</v></c>
    </row>\n`;
            rowNum++;
        }

        // Spacer row
        xmlRows += `    <row r="${rowNum}" ht="6" customHeight="1"/>\n`;
        rowNum++;

        // Header row
        const headerStartRow = rowNum;
        let headerCells = '';
        headers.forEach((h, i) => {
            const colRef = numToColRef(i + 1);
            headerCells += `      <c r="${colRef}${rowNum}" s="3" t="s"><v>${headerSSIdxs[i]}</v></c>\n`;
        });
        xmlRows += `    <row r="${rowNum}" ht="24" customHeight="1">\n${headerCells}    </row>\n`;
        rowNum++;

        // Data rows
        dataSSIdxs.forEach((rowData, rIdx) => {
            const isEven = rIdx % 2 === 0;
            let cells = '';

            rowData.forEach((cellInfo, cIdx) => {
                const colRef = numToColRef(cIdx + 1);
                const isSTT = cIdx === 0 && headers[0] === 'STT';

                if (cellInfo.type === 'n') {
                    // Numeric
                    const styleIdx = isSTT
                        ? (isEven ? 8 : 9)   // center
                        : (isEven ? 6 : 7);   // right-aligned number
                    cells += `      <c r="${colRef}${rowNum}" s="${styleIdx}"><v>${cellInfo.value}</v></c>\n`;
                } else {
                    // String
                    const styleIdx = isSTT
                        ? (isEven ? 8 : 9)   // center
                        : (isEven ? 4 : 5);   // left-aligned text
                    cells += `      <c r="${colRef}${rowNum}" s="${styleIdx}" t="s"><v>${cellInfo.ssIdx}</v></c>\n`;
                }
            });

            xmlRows += `    <row r="${rowNum}" ht="20" customHeight="1">\n${cells}    </row>\n`;
            rowNum++;
        });

        // Merge cells for title and date
        let mergeCells = '';
        const mergeList = [];
        mergeList.push(`A1:${lastColLetter}1`); // Title
        if (addDate) {
            mergeList.push(`A2:${lastColLetter}2`); // Date
        }
        if (mergeList.length > 0) {
            mergeCells = `  <mergeCells count="${mergeList.length}">\n`;
            mergeList.forEach(m => {
                mergeCells += `    <mergeCell ref="${m}"/>\n`;
            });
            mergeCells += `  </mergeCells>\n`;
        }

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane ySplit="${headerStartRow}" topLeftCell="A${headerStartRow + 1}" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
      ${colDefs}
  </cols>
  <sheetData>
${xmlRows}  </sheetData>
${mergeCells}  <pageSetup orientation="portrait" paperSize="9"/>
</worksheet>`;
    }

    // ---- Helpers ----

    function numToColRef(num) {
        let result = '';
        while (num > 0) {
            num--;
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26);
        }
        return result;
    }

    function escapeXml(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    function formatDateVN() {
        const d = new Date();
        const pad = n => String(n).padStart(2, '0');
        return `Ngày ${pad(d.getDate())} tháng ${pad(d.getMonth() + 1)} năm ${d.getFullYear()}`;
    }

    /**
     * Get the template config (for PDF use)
     */
    function getTemplate(name) {
        return TEMPLATES[name] || TEMPLATES.professional;
    }

    return { generate, getTemplate, formatDateVN };
})();
