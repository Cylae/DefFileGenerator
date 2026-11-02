document.addEventListener('DOMContentLoaded', () => {
    // =========================================================================
    // Tab Navigation
    // =========================================================================
    const navTabs = document.querySelectorAll('.nav-tab');
    const tabPanes = document.querySelectorAll('.tab-pane');
    let libraryLoaded = false;

    navTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const targetId = tab.getAttribute('data-tab');
            navTabs.forEach(t => t.classList.remove('active'));
            tabPanes.forEach(p => {
                p.classList.remove('active');
                p.classList.add('hidden');
            });

            tab.classList.add('active');
            const activePane = document.getElementById(targetId);
            if (activePane) {
                activePane.classList.remove('hidden');
                activePane.classList.add('active');
            }

            if (targetId === 'tab-library' && !libraryLoaded) {
                loadLibrary();
            }
        });
    });

    function escapeHtml(str) {
        return String(str ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function triggerDownload(content, filename) {
        const blob = new Blob(['\ufeff', content], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    // =========================================================================
    // Helper: Wire Drag & Drop for a Zone
    // =========================================================================
    function setupDropZone(dropZoneEl, fileInputEl, onFileSelected) {
        ['dragenter', 'dragover'].forEach(eventName => {
            dropZoneEl.addEventListener(eventName, (e) => {
                e.preventDefault();
                dropZoneEl.classList.add('dragover');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZoneEl.addEventListener(eventName, (e) => {
                e.preventDefault();
                dropZoneEl.classList.remove('dragover');
            });
        });

        dropZoneEl.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length) onFileSelected(files[0]);
        });

        fileInputEl.addEventListener('change', (e) => {
            if (e.target.files.length) onFileSelected(e.target.files[0]);
        });
    }

    // =========================================================================
    // TAB 1: Convert Datasheet
    // =========================================================================
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileDetails = document.getElementById('fileDetails');
    const fileName = document.getElementById('fileName');
    const removeFileBtn = document.getElementById('removeFileBtn');
    const convertBtn = document.getElementById('convertBtn');
    const resultsSection = document.getElementById('resultsSection');
    const previewBody = document.getElementById('previewBody');
    const registerBadge = document.getElementById('registerBadge');
    const validBadge = document.getElementById('validBadge');
    const downloadBtn = document.getElementById('downloadBtn');
    const convertLabel = convertBtn.innerHTML;

    let selectedFile = null;
    let convertedCsvContent = null;
    let convertedFilename = 'definition.csv';
    let requestVersion = 0;
    let pendingController = null;

    function resetResults() {
        convertedCsvContent = null;
        resultsSection.classList.add('hidden');
        previewBody.innerHTML = '';
        downloadBtn.disabled = true;
    }

    function invalidateConversion() {
        requestVersion += 1;
        if (pendingController) pendingController.abort();
        pendingController = null;
        resetResults();
        convertBtn.disabled = !selectedFile;
        convertBtn.innerHTML = convertLabel;
    }

    document.getElementById('convertForm').addEventListener('input', invalidateConversion);

    setupDropZone(dropZone, fileInput, (file) => {
        selectedFile = file;
        invalidateConversion();
        fileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        dropZone.classList.add('hidden');
        fileDetails.classList.remove('hidden');
        convertBtn.disabled = false;
    });

    removeFileBtn.addEventListener('click', () => {
        selectedFile = null;
        invalidateConversion();
        fileInput.value = '';
        dropZone.classList.remove('hidden');
        fileDetails.classList.add('hidden');
        convertBtn.disabled = true;
    });

    convertBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        const version = ++requestVersion;
        const controller = new AbortController();
        pendingController = controller;
        resetResults();
        convertBtn.disabled = true;
        convertBtn.textContent = 'Processing...';

        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('manufacturer', document.getElementById('mfgInput').value || 'Manufacturer');
        formData.append('model', document.getElementById('modelInput').value || 'Model');
        formData.append('protocol', document.getElementById('protocolSelect').value);
        formData.append('category', document.getElementById('categorySelect').value);
        formData.append('address_offset', document.getElementById('offsetInput').value || '0');
        formData.append('forced_write', document.getElementById('forcedWriteInput').value || '');

        try {
            const response = await fetch('/api/convert', {
                method: 'POST',
                body: formData,
                signal: controller.signal
            });
            if (version !== requestVersion) return;

            if (!response.ok) {
                const errData = await response.json();
                if (version !== requestVersion) return;
                alert(`Conversion Error: ${errData.detail || 'Unknown error'}`);
                return;
            }

            const data = await response.json();
            if (version !== requestVersion) return;
            convertedCsvContent = data.csv_content;
            convertedFilename = data.filename;

            const skipped = Math.max(0, (data.extracted_count || data.register_count) - data.register_count);
            registerBadge.textContent = `${data.register_count} registers generated${skipped ? ` (${skipped} skipped)` : ''}`;
            validBadge.textContent = data.is_valid === true ? 'Valid' : 'Invalid';
            validBadge.classList.toggle('badge-success', data.is_valid === true);
            validBadge.classList.toggle('badge-error', data.is_valid !== true);
            downloadBtn.disabled = data.is_valid !== true;
            renderConvertPreview(data.preview);
            resultsSection.classList.remove('hidden');

        } catch (err) {
            if (version === requestVersion && err.name !== 'AbortError') {
                alert(`Network or Server Error: ${err.message}`);
            }
        } finally {
            if (version === requestVersion) {
                pendingController = null;
                convertBtn.disabled = !selectedFile;
                convertBtn.innerHTML = convertLabel;
            }
        }
    });

    function renderConvertPreview(rows) {
        previewBody.innerHTML = '';
        rows.forEach(r => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${escapeHtml(r.Name || '')}</td>
                <td><code>${escapeHtml(r.Tag || '')}</code></td>
                <td>${escapeHtml(r.RegisterType || 'Holding Register')}</td>
                <td><code>${escapeHtml(r.Address ?? '')}</code></td>
                <td>${escapeHtml(r.Type || '')}</td>
                <td>${escapeHtml(r.Factor ?? '1')}</td>
                <td>${escapeHtml(r.Offset ?? '0')}</td>
                <td>${escapeHtml(r.Unit || '')}</td>
                <td>${escapeHtml(r.Action || '1')}</td>
            `;
            previewBody.appendChild(tr);
        });
    }

    downloadBtn.addEventListener('click', () => {
        if (!convertedCsvContent || downloadBtn.disabled) return;
        triggerDownload(convertedCsvContent, convertedFilename);
    });

    // =========================================================================
    // TAB 2: SunPM Tagger
    // =========================================================================
    const tagDropZone = document.getElementById('tagDropZone');
    const tagFileInput = document.getElementById('tagFileInput');
    const tagFileDetails = document.getElementById('tagFileDetails');
    const tagFileName = document.getElementById('tagFileName');
    const tagRemoveFileBtn = document.getElementById('tagRemoveFileBtn');
    const tagProcessBtn = document.getElementById('tagProcessBtn');
    const tagResultsSection = document.getElementById('tagResultsSection');
    const tagPreviewBody = document.getElementById('tagPreviewBody');
    const tagCountBadge = document.getElementById('tagCountBadge');
    const tagRegisterBadge = document.getElementById('tagRegisterBadge');
    const tagDownloadBtn = document.getElementById('tagDownloadBtn');
    const tagCategorySelect = document.getElementById('tagCategorySelect');
    const tagOverwriteCheck = document.getElementById('tagOverwriteCheck');

    let tagSelectedFile = null;
    let taggedCsvContent = null;
    let taggedFilename = 'definition_tagged.csv';

    setupDropZone(tagDropZone, tagFileInput, (file) => {
        tagSelectedFile = file;
        tagFileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        tagDropZone.classList.add('hidden');
        tagFileDetails.classList.remove('hidden');
        tagProcessBtn.disabled = false;
        tagResultsSection.classList.add('hidden');
    });

    tagRemoveFileBtn.addEventListener('click', () => {
        tagSelectedFile = null;
        tagFileInput.value = '';
        tagDropZone.classList.remove('hidden');
        tagFileDetails.classList.add('hidden');
        tagProcessBtn.disabled = true;
        tagResultsSection.classList.add('hidden');
    });

    tagProcessBtn.addEventListener('click', async () => {
        if (!tagSelectedFile) return;

        tagProcessBtn.disabled = true;
        tagProcessBtn.textContent = 'Enriching tags...';

        const formData = new FormData();
        formData.append('file', tagSelectedFile);
        if (tagCategorySelect.value) {
            formData.append('category', tagCategorySelect.value);
        }
        formData.append('overwrite', tagOverwriteCheck.checked ? 'true' : 'false');

        try {
            const response = await fetch('/api/tag', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                alert(`Tagging Error: ${errData.detail || 'Failed to tag file'}`);
                return;
            }

            const data = await response.json();
            taggedCsvContent = data.csv_content;
            taggedFilename = data.filename || 'definition_tagged.csv';

            tagCountBadge.textContent = `${data.updated_count} tags updated`;
            tagRegisterBadge.textContent = `${data.total_registers} registers`;

            tagPreviewBody.innerHTML = '';
            (data.preview || []).forEach(r => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${escapeHtml(r.Name || '')}</td>
                    <td><code>${escapeHtml(r.Tag || '')}</code></td>
                    <td><code>${escapeHtml(r.Address ?? '')}</code></td>
                    <td>${escapeHtml(r.Type || '')}</td>
                    <td>${escapeHtml(r.Unit || '')}</td>
                    <td>${escapeHtml(r.Action || '1')}</td>
                `;
                tagPreviewBody.appendChild(tr);
            });

            tagResultsSection.classList.remove('hidden');
        } catch (err) {
            alert(`Network or Server Error: ${err.message}`);
        } finally {
            tagProcessBtn.disabled = false;
            tagProcessBtn.textContent = 'Apply Canonical SunPM Tags';
        }
    });

    tagDownloadBtn.addEventListener('click', () => {
        if (!taggedCsvContent) return;
        triggerDownload(taggedCsvContent, taggedFilename);
    });

    // =========================================================================
    // TAB 3: Validator & Auditor
    // =========================================================================
    const auditDropZone = document.getElementById('auditDropZone');
    const auditFileInput = document.getElementById('auditFileInput');
    const auditFileDetails = document.getElementById('auditFileDetails');
    const auditFileName = document.getElementById('auditFileName');
    const auditRemoveFileBtn = document.getElementById('auditRemoveFileBtn');
    const auditProcessBtn = document.getElementById('auditProcessBtn');
    const auditResultsSection = document.getElementById('auditResultsSection');
    const auditStatusBadge = document.getElementById('auditStatusBadge');
    const auditRegistersBadge = document.getElementById('auditRegistersBadge');
    const auditErrorsBadge = document.getElementById('auditErrorsBadge');
    const auditWarningsBadge = document.getElementById('auditWarningsBadge');
    const auditDownloadFixedBtn = document.getElementById('auditDownloadFixedBtn');
    const auditIssuesList = document.getElementById('auditIssuesList');
    const auditStrictOverlapCheck = document.getElementById('auditStrictOverlapCheck');
    const auditAutoFixCheck = document.getElementById('auditAutoFixCheck');

    let auditSelectedFile = null;
    let auditFixedCsvContent = null;
    let auditFixedFilename = 'definition_repaired.csv';

    setupDropZone(auditDropZone, auditFileInput, (file) => {
        auditSelectedFile = file;
        auditFileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        auditDropZone.classList.add('hidden');
        auditFileDetails.classList.remove('hidden');
        auditProcessBtn.disabled = false;
        auditResultsSection.classList.add('hidden');
    });

    auditRemoveFileBtn.addEventListener('click', () => {
        auditSelectedFile = null;
        auditFileInput.value = '';
        auditDropZone.classList.remove('hidden');
        auditFileDetails.classList.add('hidden');
        auditProcessBtn.disabled = true;
        auditResultsSection.classList.add('hidden');
    });

    auditProcessBtn.addEventListener('click', async () => {
        if (!auditSelectedFile) return;

        auditProcessBtn.disabled = true;
        auditProcessBtn.textContent = 'Auditing...';

        const formData = new FormData();
        formData.append('file', auditSelectedFile);
        formData.append('strict_overlap', auditStrictOverlapCheck.checked ? 'true' : 'false');
        formData.append('auto_fix', auditAutoFixCheck.checked ? 'true' : 'false');

        try {
            const response = await fetch('/api/audit', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                alert(`Audit Error: ${errData.detail || 'Audit failed'}`);
                return;
            }

            const data = await response.json();

            // Status Badge
            auditStatusBadge.textContent = data.is_valid ? 'Valid' : 'Issues Found';
            auditStatusBadge.classList.toggle('badge-success', data.is_valid);
            auditStatusBadge.classList.toggle('badge-error', !data.is_valid);

            auditRegistersBadge.textContent = `${data.register_count} registers`;

            if (data.error_count > 0) {
                auditErrorsBadge.textContent = `${data.error_count} errors`;
                auditErrorsBadge.classList.remove('hidden');
            } else {
                auditErrorsBadge.classList.add('hidden');
            }

            if (data.warning_count > 0) {
                auditWarningsBadge.textContent = `${data.warning_count} warnings`;
                auditWarningsBadge.classList.remove('hidden');
            } else {
                auditWarningsBadge.classList.add('hidden');
            }

            // Render Issues List
            auditIssuesList.innerHTML = '';
            if (data.issues && data.issues.length > 0) {
                data.issues.forEach(issue => {
                    const card = document.createElement('div');
                    const level = (issue.level || 'info').toLowerCase();
                    card.className = `audit-issue-card ${level}`;
                    card.innerHTML = `
                        <span class="issue-level">${escapeHtml(issue.level || 'INFO')}</span>
                        <span>${issue.field ? `<strong>${escapeHtml(issue.field)}:</strong> ` : ''}${escapeHtml(issue.message || '')}</span>
                    `;
                    auditIssuesList.appendChild(card);
                });
            } else {
                const cleanCard = document.createElement('div');
                cleanCard.className = 'audit-issue-card info';
                cleanCard.innerHTML = `<span class="issue-level">PASS</span><span>Definition passed all integrity checks. Zero syntax, overlap, or semantic issues detected.</span>`;
                auditIssuesList.appendChild(cleanCard);
            }

            // Fixed CSV Download
            if (data.fixed_csv_content) {
                auditFixedCsvContent = data.fixed_csv_content;
                auditFixedFilename = data.filename ? `repaired_${data.filename}` : 'definition_repaired.csv';
                auditDownloadFixedBtn.classList.remove('hidden');
            } else {
                auditFixedCsvContent = null;
                auditDownloadFixedBtn.classList.add('hidden');
            }

            auditResultsSection.classList.remove('hidden');
        } catch (err) {
            alert(`Network or Server Error: ${err.message}`);
        } finally {
            auditProcessBtn.disabled = false;
            auditProcessBtn.textContent = 'Audit Definition';
        }
    });

    auditDownloadFixedBtn.addEventListener('click', () => {
        if (!auditFixedCsvContent) return;
        triggerDownload(auditFixedCsvContent, auditFixedFilename);
    });

    // =========================================================================
    // TAB 4: Equipment Library
    // =========================================================================
    const libraryTableBody = document.getElementById('libraryTableBody');
    const libraryCountBadge = document.getElementById('libraryCountBadge');
    const librarySearchInput = document.getElementById('librarySearchInput');
    let libraryItems = [];

    async function loadLibrary() {
        try {
            const response = await fetch('/api/library');
            if (!response.ok) {
                libraryTableBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #ef4444;">Failed to load definitions catalog.</td></tr>`;
                return;
            }
            const data = await response.json();
            libraryItems = data.definitions || [];
            libraryLoaded = true;
            libraryCountBadge.textContent = `${data.count} models`;
            renderLibraryTable(libraryItems);
        } catch (err) {
            libraryTableBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #ef4444;">Network error: ${escapeHtml(err.message)}</td></tr>`;
        }
    }

    function renderLibraryTable(items) {
        libraryTableBody.innerHTML = '';
        if (items.length === 0) {
            libraryTableBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-secondary);">No matching definitions found.</td></tr>`;
            return;
        }

        items.forEach(item => {
            const tr = document.createElement('tr');
            const sizeKb = (item.size_bytes / 1024).toFixed(1);
            tr.innerHTML = `
                <td><span class="badge badge-info">${escapeHtml(item.category)}</span></td>
                <td><strong>${escapeHtml(item.manufacturer)}</strong></td>
                <td><code>${escapeHtml(item.model)}</code></td>
                <td>${sizeKb} KB</td>
                <td>
                    <a href="/api/library/${encodeURIComponent(item.filename)}" class="btn btn-primary btn-sm" download>
                        Download CSV
                    </a>
                </td>
            `;
            libraryTableBody.appendChild(tr);
        });
    }

    librarySearchInput.addEventListener('input', () => {
        const query = librarySearchInput.value.trim().toLowerCase();
        if (!query) {
            renderLibraryTable(libraryItems);
            return;
        }

        const filtered = libraryItems.filter(item => {
            return (item.manufacturer && item.manufacturer.toLowerCase().includes(query)) ||
                   (item.model && item.model.toLowerCase().includes(query)) ||
                   (item.category && item.category.toLowerCase().includes(query)) ||
                   (item.filename && item.filename.toLowerCase().includes(query));
        });
        renderLibraryTable(filtered);
    });
});
