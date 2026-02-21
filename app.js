/**
 * App.js — Main application controller
 * 3-step workflow: Template → Data → Export
 */
(function () {
    'use strict';

    // ===== State =====
    let templateData = null;  // Result from TemplateEngine.analyzeTemplate
    let templateSummary = null;
    let workbookData = null;  // Result from XLSXReader.read or readCSV
    let columnMapping = null; // Map: templateColIndex -> dataColIndex

    // ===== DOM References =====
    const $ = id => document.getElementById(id);

    // Step 1: Template
    const templateDropZone = $('templateDropZone');
    const templateFileInput = $('templateFileInput');
    const templateInfoCard = $('templateInfoCard');
    const templateFileName = $('templateFileName');
    const templateFileMeta = $('templateFileMeta');
    const templateDetailsGrid = $('templateDetailsGrid');
    const templateColumns = $('templateColumns');
    const btnRemoveTemplate = $('btnRemoveTemplate');

    // Step 2: Data upload
    const stepUpload = $('stepUpload');
    const dropZone = $('dropZone');
    const fileInput = $('fileInput');
    const fileInfo = $('fileInfo');
    const fileName = $('fileName');
    const fileSize = $('fileSize');
    const btnRemoveFile = $('btnRemoveFile');

    // Step 3: Config & Export
    const stepConfig = $('stepConfig');
    const sheetSelect = $('sheetSelect');
    const columnMappingCard = $('columnMappingCard');
    const mappingGrid = $('mappingGrid');
    const builtinTemplateCard = $('builtinTemplateCard');
    const titleCard = $('titleCard');
    const optionsCard = $('optionsCard');
    const previewToolbar = $('previewToolbar');
    const previewContainer = $('previewContainer');
    const btnExportExcel = $('btnExportExcel');
    const btnExportPDF = $('btnExportPDF');

    // Shared
    const loadingOverlay = $('loadingOverlay');
    const loadingText = $('loadingText');

    // ===== Utility =====
    function formatSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }

    function showLoading(text) {
        loadingText.textContent = text || 'Đang xử lý...';
        loadingOverlay.style.display = 'flex';
    }

    function hideLoading() {
        loadingOverlay.style.display = 'none';
    }

    function showToast(message, type = 'success') {
        const container = $('toastContainer');
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.textContent = message;
        container.appendChild(toast);
        setTimeout(() => toast.classList.add('show'), 10);
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }

    function enableStep(stepEl) {
        stepEl.classList.remove('disabled-section');
    }

    function disableStep(stepEl) {
        stepEl.classList.add('disabled-section');
    }

    // ===== Step 1: Template Upload =====
    function setupTemplateUpload() {
        templateDropZone.addEventListener('click', () => templateFileInput.click());
        templateDropZone.addEventListener('dragover', e => {
            e.preventDefault();
            templateDropZone.classList.add('dragover');
        });
        templateDropZone.addEventListener('dragleave', () => templateDropZone.classList.remove('dragover'));
        templateDropZone.addEventListener('drop', e => {
            e.preventDefault();
            templateDropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                handleTemplateFile(e.dataTransfer.files[0]);
            }
        });
        templateFileInput.addEventListener('change', () => {
            if (templateFileInput.files.length > 0) {
                handleTemplateFile(templateFileInput.files[0]);
            }
        });
        btnRemoveTemplate.addEventListener('click', removeTemplate);
    }

    async function handleTemplateFile(file) {
        if (!file.name.endsWith('.xlsx')) {
            showToast('Chỉ hỗ trợ file .xlsx cho file mẫu', 'error');
            return;
        }

        showLoading('Đang phân tích file mẫu...');
        try {
            const buffer = await file.arrayBuffer();
            templateData = await TemplateEngine.analyzeTemplate(buffer);
            templateSummary = TemplateEngine.getTemplateSummary(templateData);

            // Show template info
            templateDropZone.style.display = 'none';
            templateInfoCard.style.display = 'block';
            templateFileName.textContent = file.name;
            templateFileMeta.textContent = `${formatSize(file.size)} • ${templateSummary.sheetCount} sheet`;

            // Show details grid
            templateDetailsGrid.innerHTML = `
                <div class="detail-item">
                    <span class="detail-label">Cột</span>
                    <span class="detail-value">${templateSummary.maxColumns}</span>
                </div>
                <div class="detail-item">
                    <span class="detail-label">Dòng header</span>
                    <span class="detail-value">${templateSummary.headerRowCount}</span>
                </div>
                <div class="detail-item">
                    <span class="detail-label">Dòng dữ liệu</span>
                    <span class="detail-value">${templateSummary.dataRowCount}</span>
                </div>
                <div class="detail-item">
                    <span class="detail-label">Dòng footer</span>
                    <span class="detail-value">${templateSummary.footerRowCount}</span>
                </div>
                <div class="detail-item">
                    <span class="detail-label">Merge cells</span>
                    <span class="detail-value">${templateSummary.mergeCount}</span>
                </div>
                <div class="detail-item">
                    <span class="detail-label">Sheets</span>
                    <span class="detail-value">${templateSummary.sheetNames.join(', ')}</span>
                </div>
            `;

            // Show column headers from template
            const headers = templateSummary.columnHeaders.filter(h => h.trim().length > 0);
            templateColumns.innerHTML = `
                <div class="template-col-label">Cột phát hiện trong mẫu:</div>
                <div class="template-col-tags">
                    ${headers.map(h => `<span class="col-tag">${h}</span>`).join('')}
                </div>
            `;

            // Enable Step 2
            enableStep(stepUpload);

            // Update UI: show column mapping, hide built-in templates
            columnMappingCard.style.display = 'block';
            builtinTemplateCard.style.display = 'none';
            // Hide title/options that don't apply to template mode
            titleCard.style.display = 'none';
            optionsCard.style.display = 'none';

            showToast('Đã phân tích mẫu thành công!', 'success');
        } catch (err) {
            console.error('Template analysis error:', err);
            showToast('Lỗi phân tích file mẫu: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    function removeTemplate() {
        templateData = null;
        templateSummary = null;
        columnMapping = null;
        templateDropZone.style.display = '';
        templateInfoCard.style.display = 'none';
        templateFileInput.value = '';

        // Hide column mapping, show built-in templates
        columnMappingCard.style.display = 'none';
        builtinTemplateCard.style.display = '';
        titleCard.style.display = '';
        optionsCard.style.display = '';

        // If no data loaded, disable step 2 — but actually step 2 is always enabled
        // Just reset if needed
        showToast('Đã xóa file mẫu', 'success');
    }

    // ===== Step 2: Data Upload =====
    function setupDataUpload() {
        dropZone.addEventListener('click', () => {
            if (stepUpload.classList.contains('disabled-section')) return;
            fileInput.click();
        });
        dropZone.addEventListener('dragover', e => {
            e.preventDefault();
            if (!stepUpload.classList.contains('disabled-section')) {
                dropZone.classList.add('dragover');
            }
        });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            if (stepUpload.classList.contains('disabled-section')) return;
            if (e.dataTransfer.files.length > 0) {
                handleDataFile(e.dataTransfer.files[0]);
            }
        });
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) {
                handleDataFile(fileInput.files[0]);
            }
        });
        btnRemoveFile.addEventListener('click', removeDataFile);
    }

    async function handleDataFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'csv'].includes(ext)) {
            showToast('Chỉ hỗ trợ file .xlsx và .csv', 'error');
            return;
        }
        if (file.size > 10 * 1024 * 1024) {
            showToast('File quá lớn (tối đa 10MB)', 'error');
            return;
        }

        showLoading('Đang đọc dữ liệu...');
        try {
            if (ext === 'csv') {
                const text = await file.text();
                workbookData = XLSXReader.readCSV(text);
            } else {
                const buffer = await file.arrayBuffer();
                workbookData = await XLSXReader.read(buffer);
            }

            // Show file info
            dropZone.style.display = 'none';
            fileInfo.style.display = 'flex';
            fileName.textContent = file.name;
            fileSize.textContent = formatSize(file.size);

            // Enable Step 3
            enableStep(stepConfig);

            // Populate sheet selector
            sheetSelect.innerHTML = '';
            workbookData.sheetNames.forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                sheetSelect.appendChild(opt);
            });

            // Build column mapping if template is loaded
            if (templateData) {
                buildColumnMapping();
            }

            // Render preview
            renderPreview();
            previewToolbar.style.display = 'flex';
            previewContainer.style.display = 'block';

            showToast('Đã đọc dữ liệu thành công!', 'success');
        } catch (err) {
            console.error('Data read error:', err);
            showToast('Lỗi đọc file: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    function removeDataFile() {
        workbookData = null;
        columnMapping = null;
        dropZone.style.display = '';
        fileInfo.style.display = 'none';
        fileInput.value = '';
        previewToolbar.style.display = 'none';
        previewContainer.style.display = 'none';
        disableStep(stepConfig);
        showToast('Đã xóa file dữ liệu', 'success');
    }

    // ===== Column Mapping =====
    function buildColumnMapping() {
        if (!templateData || !workbookData) return;

        // Use ALL template columns (including empty headers) to preserve column positions
        const allTplHeaders = templateSummary.columnHeaders;
        const selectedSheet = sheetSelect.value || workbookData.sheetNames[0];
        const dataSheet = workbookData.sheets[selectedSheet];
        if (!dataSheet) return;

        const dataHeaders = dataSheet.headers;

        // Auto-map by name similarity
        columnMapping = {};
        const mappingRows = [];

        allTplHeaders.forEach((tplHeader, tplIdx) => {
            const headerText = String(tplHeader).trim();

            if (headerText.length === 0) {
                // Empty template header — no mapping available
                columnMapping[tplIdx] = -1;
                return; // Skip UI row for empty headers
            }

            // Try to find a matching data column
            let bestMatch = -1;
            let bestScore = 0;

            dataHeaders.forEach((dh, di) => {
                const score = stringSimilarity(headerText.toLowerCase(), String(dh).toLowerCase());
                if (score > bestScore) {
                    bestScore = score;
                    bestMatch = di;
                }
            });

            // Use match if score > 0.3
            const matchIdx = bestScore > 0.3 ? bestMatch : -1;
            columnMapping[tplIdx] = matchIdx;

            const options = dataHeaders.map((dh, di) =>
                `<option value="${di}" ${di === matchIdx ? 'selected' : ''}>${dh}</option>`
            ).join('');

            mappingRows.push(`
                <div class="mapping-row">
                    <div class="mapping-template-col">
                        <span class="col-tag">${headerText}</span>
                    </div>
                    <div class="mapping-arrow">→</div>
                    <div class="mapping-data-col">
                        <select class="form-select mapping-select" data-tpl-idx="${tplIdx}">
                            <option value="-1">— Bỏ trống —</option>
                            ${options}
                        </select>
                    </div>
                </div>
            `);
        });

        mappingGrid.innerHTML = mappingRows.join('');

        // Listen for mapping changes
        mappingGrid.querySelectorAll('.mapping-select').forEach(sel => {
            sel.addEventListener('change', e => {
                const tplIdx = parseInt(e.target.dataset.tplIdx);
                columnMapping[tplIdx] = parseInt(e.target.value);
            });
        });
    }

    function stringSimilarity(a, b) {
        if (a === b) return 1;
        if (!a || !b) return 0;
        // Check containment
        if (a.includes(b) || b.includes(a)) return 0.8;
        // Simple Jaccard on trigrams
        const triA = new Set();
        const triB = new Set();
        for (let i = 0; i <= a.length - 2; i++) triA.add(a.substring(i, i + 2));
        for (let i = 0; i <= b.length - 2; i++) triB.add(b.substring(i, i + 2));
        let inter = 0;
        triA.forEach(t => { if (triB.has(t)) inter++; });
        return inter / (triA.size + triB.size - inter);
    }

    // ===== Preview =====
    function renderPreview() {
        const selectedSheet = sheetSelect.value || workbookData.sheetNames[0];
        const sheet = workbookData.sheets[selectedSheet];
        if (!sheet) return;

        const head = $('previewHead');
        const body = $('previewBody');

        // Header row
        let headerHtml = '<tr>';
        sheet.headers.forEach(h => {
            headerHtml += `<th>${escapeHtml(String(h))}</th>`;
        });
        headerHtml += '</tr>';
        head.innerHTML = headerHtml;

        // Body rows (max 100 for preview)
        const maxRows = Math.min(sheet.rows.length, 100);
        let bodyHtml = '';
        for (let i = 0; i < maxRows; i++) {
            bodyHtml += '<tr>';
            sheet.rows[i].forEach(cell => {
                bodyHtml += `<td>${escapeHtml(String(cell))}</td>`;
            });
            bodyHtml += '</tr>';
        }
        body.innerHTML = bodyHtml;

        $('rowCount').textContent = `${sheet.rows.length} dòng`;
        $('colCount').textContent = `${sheet.headers.length} cột`;
    }

    function escapeHtml(str) {
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // ===== Export: Excel =====
    async function exportExcel() {
        if (!workbookData) return;

        const selectedSheet = sheetSelect.value || workbookData.sheetNames[0];
        const sheet = workbookData.sheets[selectedSheet];
        if (!sheet) return;

        showLoading('Đang tạo file Excel...');
        try {
            let blob;

            if (templateData && columnMapping) {
                // Template-based export
                const mappedRows = mapDataToTemplate(sheet);
                blob = await TemplateEngine.generateFromTemplate(templateData, {
                    rows: mappedRows,
                    sheetName: selectedSheet,
                });
            } else {
                // Built-in template export
                blob = await XLSXWriter.generate({
                    headers: sheet.headers,
                    rows: sheet.rows,
                    title: $('reportTitle').value || 'BÁO CÁO DỮ LIỆU',
                    template: getSelectedTemplate(),
                    includeSTT: $('includeSTT').checked,
                    includeDate: $('includeDate').checked,
                    autofitColumns: $('autofitColumns').checked,
                });
            }

            const safeName = (selectedSheet || 'export').replace(/[\\/:*?"<>|]/g, '_');
            saveAs(blob, `${safeName}.xlsx`);
            showToast('Đã tải Excel thành công!', 'success');
        } catch (err) {
            console.error('Export error:', err);
            showToast('Lỗi xuất file: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    function mapDataToTemplate(sheet) {
        if (!columnMapping || !templateSummary) return sheet.rows;

        // Use ALL template columns (including empty headers) to preserve column positions
        const allTplHeaders = templateSummary.columnHeaders;
        const mappedRows = [];

        for (const row of sheet.rows) {
            const mappedRow = [];
            allTplHeaders.forEach((_, tplIdx) => {
                const dataIdx = columnMapping[tplIdx];
                if (dataIdx >= 0 && dataIdx < row.length) {
                    mappedRow.push(row[dataIdx]);
                } else {
                    mappedRow.push('');
                }
            });
            mappedRows.push(mappedRow);
        }

        return mappedRows;
    }

    function getSelectedTemplate() {
        const checked = document.querySelector('input[name="template"]:checked');
        return checked ? checked.value : 'professional';
    }

    // ===== Export: PDF =====
    async function exportPDF() {
        if (!workbookData) return;

        const selectedSheet = sheetSelect.value || workbookData.sheetNames[0];
        const sheet = workbookData.sheets[selectedSheet];
        if (!sheet) return;

        const jsPDFLib = window.jspdf || window.jsPDF;
        if (!jsPDFLib) {
            showToast('Lỗi: Không tải được thư viện jsPDF', 'error');
            return;
        }

        showLoading('Đang tạo PDF...');
        try {
            const pgSize = $('pageSize').value;
            const isLandscape = $('landscape').checked;
            const title = $('reportTitle')?.value || selectedSheet;

            const headers = templateData && columnMapping
                ? templateSummary.columnHeaders
                : sheet.headers;

            const rows = templateData && columnMapping
                ? mapDataToTemplate(sheet)
                : sheet.rows;

            // Use SVG-PDF renderer (style-aware, smaller output)
            if (typeof SVGPDFRenderer !== 'undefined') {
                const blob = await SVGPDFRenderer.renderToPDF({
                    headers: headers,
                    rows: rows,
                    templateData: templateData || null,
                    title: templateData ? '' : title, // Template has its own header zone
                    pageSize: pgSize,
                    landscape: isLandscape,
                });

                const safeName = (selectedSheet || 'export').replace(/[\\/:*?"<>|]/g, '_');
                saveAs(blob, `${safeName}.pdf`);

                const sizeKB = (blob.size / 1024).toFixed(0);
                showToast(`PDF tải thành công! (${sizeKB} KB)`, 'success');
            } else {
                // Fallback: old jsPDF autoTable approach
                showLoading('Đang tải font Unicode...');

                const doc = new jsPDFLib.jsPDF({
                    orientation: isLandscape ? 'landscape' : 'portrait',
                    unit: 'mm',
                    format: pgSize,
                });

                if (typeof FontLoader !== 'undefined') {
                    await FontLoader.registerFont(doc);
                    showLoading('Đang tạo PDF...');
                }

                const fontName = (typeof FontLoader !== 'undefined' && FontLoader.isLoaded()) ? 'NotoSans' : 'helvetica';
                doc.setFontSize(16);
                doc.setFont(fontName, 'bold');
                const pageWidth = doc.internal.pageSize.getWidth();
                doc.text(title, pageWidth / 2, 15, { align: 'center' });

                doc.autoTable({
                    head: [headers],
                    body: rows,
                    startY: 25,
                    styles: { font: fontName, fontSize: 9, cellPadding: 3 },
                    headStyles: { fillColor: [99, 102, 241], textColor: 255, fontStyle: 'bold', halign: 'center' },
                    alternateRowStyles: { fillColor: [245, 245, 255] },
                    margin: { top: 10, right: 10, bottom: 10, left: 10 },
                });

                const safeName = (selectedSheet || 'export').replace(/[\\/:*?"<>|]/g, '_');
                doc.save(`${safeName}.pdf`);
                showToast('Đã tải PDF thành công!', 'success');
            }
        } catch (err) {
            console.error('PDF export error:', err);
            showToast('Lỗi xuất PDF: ' + err.message, 'error');
        } finally {
            hideLoading();
        }
    }

    // ===== Template Options (built-in) =====
    function setupTemplateOptions() {
        document.querySelectorAll('.template-option').forEach(opt => {
            opt.addEventListener('click', () => {
                document.querySelectorAll('.template-option').forEach(o => o.classList.remove('selected'));
                opt.classList.add('selected');
            });
        });
    }

    // ===== Help Modal =====
    function setupHelp() {
        $('btnHelp').addEventListener('click', () => $('helpModal').style.display = 'flex');
        $('btnCloseHelp').addEventListener('click', () => $('helpModal').style.display = 'none');
        $('helpModal').addEventListener('click', e => {
            if (e.target === $('helpModal')) $('helpModal').style.display = 'none';
        });
    }

    // ===== Sheet change handler =====
    function setupSheetSelect() {
        sheetSelect.addEventListener('change', () => {
            if (templateData) buildColumnMapping();
            renderPreview();
        });
    }

    // ===== Initialize =====
    function init() {
        setupTemplateUpload();
        setupDataUpload();
        setupTemplateOptions();
        setupHelp();
        setupSheetSelect();

        btnExportExcel.addEventListener('click', exportExcel);
        btnExportPDF.addEventListener('click', exportPDF);

        // If no template is loaded, step 2 is still accessible (user can skip template)
        // Enable step 2 by default for non-template workflow
        enableStep(stepUpload);
    }

    // Wait for page load
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
