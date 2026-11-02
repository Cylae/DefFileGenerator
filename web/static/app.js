document.addEventListener('DOMContentLoaded', () => {
    // =========================================================================
    // Tab Navigation & Keyboard Accessibility
    // =========================================================================
    let libraryLoaded = false;
    const tabBindings = [
        { btnId: 'tabBtnConvert', paneId: 'tab-convert' },
        { btnId: 'tabBtnTag', paneId: 'tab-tag' },
        { btnId: 'tabBtnAudit', paneId: 'tab-audit' },
        {
            btnId: 'tabBtnLibrary',
            paneId: 'tab-library',
            onActivate: () => {
                if (!libraryLoaded) loadLibrary();
            }
        }
    ];

    function activateTab(index) {
        if (index < 0 || index >= tabBindings.length) return;
        const target = tabBindings[index];
        const btn = document.getElementById(target.btnId);
        if (!btn) return;
        tabBindings.forEach(other => {
            const b = document.getElementById(other.btnId);
            const p = document.getElementById(other.paneId);
            if (b) {
                if (b.classList) b.classList.remove('active');
                if (b.setAttribute) {
                    b.setAttribute('aria-selected', 'false');
                    b.setAttribute('tabindex', '-1');
                }
            }
            if (p && p.classList) {
                p.classList.remove('active');
                p.classList.add('hidden');
            }
        });
        if (btn.classList) btn.classList.add('active');
        if (btn.setAttribute) {
            btn.setAttribute('aria-selected', 'true');
            btn.setAttribute('tabindex', '0');
        }
        if (btn.focus) btn.focus();
        const pane = document.getElementById(target.paneId);
        if (pane && pane.classList) {
            pane.classList.remove('hidden');
            pane.classList.add('active');
        }
        if (target.onActivate) target.onActivate();
    }

    tabBindings.forEach(({ btnId }, idx) => {
        const btn = document.getElementById(btnId);
        if (!btn || !btn.addEventListener) return;
        btn.addEventListener('click', () => activateTab(idx));
    });

    const navTabs = document.getElementById('navTabs');
    if (navTabs && navTabs.addEventListener) {
        navTabs.addEventListener('keydown', (e) => {
            const activeIdx = tabBindings.findIndex(t => {
                const b = document.getElementById(t.btnId);
                return b && b.classList && b.classList.contains('active');
            });
            if (activeIdx === -1) return;
            if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
                if (e.preventDefault) e.preventDefault();
                activateTab((activeIdx + 1) % tabBindings.length);
            } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
                if (e.preventDefault) e.preventDefault();
                activateTab((activeIdx - 1 + tabBindings.length) % tabBindings.length);
            } else if (e.key === 'Home') {
                if (e.preventDefault) e.preventDefault();
                activateTab(0);
            } else if (e.key === 'End') {
                if (e.preventDefault) e.preventDefault();
                activateTab(tabBindings.length - 1);
            }
        });
    }

    // =========================================================================
    // Utilities: Toast, Clipboard, and HTML Escaping
    // =========================================================================
    function showToast(message, type = 'info') {
        const container = document.getElementById('toastContainer');
        if (container && container.appendChild) {
            const toast = document.createElement('div');
            toast.className = `toast toast-${type}`;
            toast.textContent = message;
            container.appendChild(toast);
            setTimeout(() => {
                if (container.removeChild) {
                    container.removeChild(toast);
                }
            }, 3500);
        }
        if (type === 'error' && typeof alert === 'function' && typeof window === 'undefined') {
            alert(message);
        }
    }

    async function copyToClipboard(text, successMsg = 'Copied to clipboard!') {
        if (!text) return;
        try {
            if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
                await navigator.clipboard.writeText(text);
            }
            showToast(successMsg, 'success');
        } catch {
            showToast('Unable to copy to clipboard', 'error');
        }
    }

    function setupTableFilter(inputEl, tbodyEl) {
        if (!inputEl || !inputEl.addEventListener || !tbodyEl) return;
        inputEl.addEventListener('input', () => {
            const q = inputEl.value.trim().toLowerCase();
            const rows = tbodyEl.children || [];
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                if (!row) continue;
                const text = (row.textContent || row.innerText || row.innerHTML || '').toLowerCase();
                const visible = !q || text.includes(q);
                if (row.style) {
                    row.style.display = visible ? '' : 'none';
                }
                if (row.classList) {
                    row.classList.toggle('hidden', !visible);
                }
            }
        });
    }

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
        if (!dropZoneEl || !dropZoneEl.addEventListener) return;
        ['dragenter', 'dragover'].forEach(eventName => {
            dropZoneEl.addEventListener(eventName, (e) => {
                if (e && e.preventDefault) e.preventDefault();
                if (dropZoneEl.classList) dropZoneEl.classList.add('dragover');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZoneEl.addEventListener(eventName, (e) => {
                if (e && e.preventDefault) e.preventDefault();
                if (dropZoneEl.classList) dropZoneEl.classList.remove('dragover');
            });
        });

        dropZoneEl.addEventListener('drop', (e) => {
            const files = e && e.dataTransfer ? e.dataTransfer.files : [];
            if (files && files.length) onFileSelected(files[0]);
        });

        if (fileInputEl && fileInputEl.addEventListener) {
            fileInputEl.addEventListener('change', (e) => {
                const files = e && e.target ? e.target.files : [];
                if (files && files.length) onFileSelected(files[0]);
            });
        }
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
    const copyBtn = document.getElementById('copyBtn');
    const previewFilterInput = document.getElementById('previewFilterInput');
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
        if (copyBtn) copyBtn.disabled = true;
        if (previewFilterInput) previewFilterInput.value = '';
    }

    function invalidateConversion() {
        requestVersion += 1;
        if (pendingController) pendingController.abort();
        pendingController = null;
        resetResults();
        convertBtn.disabled = !selectedFile;
        convertBtn.innerHTML = convertLabel;
    }

    const convertForm = document.getElementById('convertForm');
    if (convertForm && convertForm.addEventListener) {
        convertForm.addEventListener('input', invalidateConversion);
    }

    setupDropZone(dropZone, fileInput, (file) => {
        selectedFile = file;
        invalidateConversion();
        fileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        dropZone.classList.add('hidden');
        fileDetails.classList.remove('hidden');
        convertBtn.disabled = false;
    });

    if (removeFileBtn && removeFileBtn.addEventListener) {
        removeFileBtn.addEventListener('click', () => {
            selectedFile = null;
            invalidateConversion();
            fileInput.value = '';
            dropZone.classList.remove('hidden');
            fileDetails.classList.add('hidden');
            convertBtn.disabled = true;
        });
    }

    if (convertBtn && convertBtn.addEventListener) {
        convertBtn.addEventListener('click', async () => {
            if (!selectedFile) return;

            const version = ++requestVersion;
            const controller = new AbortController();
            pendingController = controller;
            resetResults();
            convertBtn.disabled = true;
            convertBtn.textContent = 'Processing...';

            const mfgInput = document.getElementById('mfgInput');
            const modelInput = document.getElementById('modelInput');
            const protocolSelect = document.getElementById('protocolSelect');
            const categorySelect = document.getElementById('categorySelect');
            const offsetInput = document.getElementById('offsetInput');
            const forcedWriteInput = document.getElementById('forcedWriteInput');

            const formData = new FormData();
            formData.append('file', selectedFile);
            formData.append('manufacturer', (mfgInput && mfgInput.value) || 'Manufacturer');
            formData.append('model', (modelInput && modelInput.value) || 'Model');
            formData.append('protocol', (protocolSelect && protocolSelect.value) || 'modbusRTU');
            formData.append('category', (categorySelect && categorySelect.value) || 'Inverter');
            formData.append('address_offset', (offsetInput && offsetInput.value) || '0');
            formData.append('forced_write', (forcedWriteInput && forcedWriteInput.value) || '');

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
                    showToast(`Conversion Error: ${errData.detail || 'Unknown error'}`, 'error');
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
                if (copyBtn) copyBtn.disabled = data.is_valid !== true;
                renderConvertPreview(data.preview || []);
                resultsSection.classList.remove('hidden');

            } catch (err) {
                if (version === requestVersion && err.name !== 'AbortError') {
                    showToast(`Network or Server Error: ${err.message}`, 'error');
                }
            } finally {
                if (version === requestVersion) {
                    pendingController = null;
                    convertBtn.disabled = !selectedFile;
                    convertBtn.innerHTML = convertLabel;
                }
            }
        });
    }

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

    if (downloadBtn && downloadBtn.addEventListener) {
        downloadBtn.addEventListener('click', () => {
            if (!convertedCsvContent || downloadBtn.disabled) return;
            triggerDownload(convertedCsvContent, convertedFilename);
        });
    }

    if (copyBtn && copyBtn.addEventListener) {
        copyBtn.addEventListener('click', () => {
            if (!convertedCsvContent || copyBtn.disabled) return;
            copyToClipboard(convertedCsvContent, 'Definition CSV copied to clipboard!');
        });
    }
    setupTableFilter(previewFilterInput, previewBody);

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
    const tagCopyBtn = document.getElementById('tagCopyBtn');
    const tagFilterInput = document.getElementById('tagFilterInput');
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

    if (tagRemoveFileBtn && tagRemoveFileBtn.addEventListener) {
        tagRemoveFileBtn.addEventListener('click', () => {
            tagSelectedFile = null;
            if (tagFileInput) tagFileInput.value = '';
            tagDropZone.classList.remove('hidden');
            tagFileDetails.classList.add('hidden');
            tagProcessBtn.disabled = true;
            tagResultsSection.classList.add('hidden');
        });
    }

    if (tagProcessBtn && tagProcessBtn.addEventListener) {
        tagProcessBtn.addEventListener('click', async () => {
            if (!tagSelectedFile) return;

            tagProcessBtn.disabled = true;
            tagProcessBtn.textContent = 'Enriching tags...';

            const formData = new FormData();
            formData.append('file', tagSelectedFile);
            if (tagCategorySelect && tagCategorySelect.value) {
                formData.append('category', tagCategorySelect.value);
            }
            formData.append('overwrite', (tagOverwriteCheck && tagOverwriteCheck.checked) ? 'true' : 'false');

            try {
                const response = await fetch('/api/tag', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const errData = await response.json();
                    showToast(`Tagging Error: ${errData.detail || 'Failed to tag file'}`, 'error');
                    return;
                }

                const data = await response.json();
                taggedCsvContent = data.csv_content;
                taggedFilename = data.filename || 'definition_tagged.csv';

                const updatedCount = data.tags_updated ?? data.updated_count ?? 0;
                const regCount = data.register_count ?? data.total_registers ?? 0;
                if (tagCountBadge) tagCountBadge.textContent = `${updatedCount} tags updated`;
                if (tagRegisterBadge) tagRegisterBadge.textContent = `${regCount} registers`;

                if (tagPreviewBody) {
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
                }

                if (tagCopyBtn) tagCopyBtn.disabled = false;
                if (tagResultsSection) tagResultsSection.classList.remove('hidden');
            } catch (err) {
                showToast(`Network or Server Error: ${err.message}`, 'error');
            } finally {
                tagProcessBtn.disabled = false;
                tagProcessBtn.textContent = 'Apply Canonical SunPM Tags';
            }
        });
    }

    if (tagDownloadBtn && tagDownloadBtn.addEventListener) {
        tagDownloadBtn.addEventListener('click', () => {
            if (!taggedCsvContent) return;
            triggerDownload(taggedCsvContent, taggedFilename);
        });
    }

    if (tagCopyBtn && tagCopyBtn.addEventListener) {
        tagCopyBtn.addEventListener('click', () => {
            if (!taggedCsvContent) return;
            copyToClipboard(taggedCsvContent, 'Enriched CSV copied to clipboard!');
        });
    }
    setupTableFilter(tagFilterInput, tagPreviewBody);

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
    const auditCopyFixedBtn = document.getElementById('auditCopyFixedBtn');
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

    if (auditRemoveFileBtn && auditRemoveFileBtn.addEventListener) {
        auditRemoveFileBtn.addEventListener('click', () => {
            auditSelectedFile = null;
            if (auditFileInput) auditFileInput.value = '';
            auditDropZone.classList.remove('hidden');
            auditFileDetails.classList.add('hidden');
            auditProcessBtn.disabled = true;
            auditResultsSection.classList.add('hidden');
        });
    }

    if (auditProcessBtn && auditProcessBtn.addEventListener) {
        auditProcessBtn.addEventListener('click', async () => {
            if (!auditSelectedFile) return;

            auditProcessBtn.disabled = true;
            auditProcessBtn.textContent = 'Auditing...';

            const formData = new FormData();
            formData.append('file', auditSelectedFile);
            formData.append('strict_overlap', (auditStrictOverlapCheck && auditStrictOverlapCheck.checked) ? 'true' : 'false');
            formData.append('auto_fix', (auditAutoFixCheck && auditAutoFixCheck.checked) ? 'true' : 'false');

            try {
                const response = await fetch('/api/audit', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const errData = await response.json();
                    showToast(`Audit Error: ${errData.detail || 'Audit failed'}`, 'error');
                    return;
                }

                const data = await response.json();

                const isValid = Boolean(data.is_valid ?? data.valid);
                if (auditStatusBadge) {
                    auditStatusBadge.textContent = isValid ? 'Valid' : 'Issues Found';
                    auditStatusBadge.classList.toggle('badge-success', isValid);
                    auditStatusBadge.classList.toggle('badge-error', !isValid);
                }

                if (auditRegistersBadge) {
                    auditRegistersBadge.textContent = `${data.register_count ?? 0} registers`;
                }

                const errCount = data.error_count ?? (Array.isArray(data.errors) ? data.errors.length : 0);
                if (auditErrorsBadge) {
                    if (errCount > 0) {
                        auditErrorsBadge.textContent = `${errCount} errors`;
                        auditErrorsBadge.classList.remove('hidden');
                    } else {
                        auditErrorsBadge.classList.add('hidden');
                    }
                }

                const warnCount = data.warning_count ?? (Array.isArray(data.warnings) ? data.warnings.length : 0);
                if (auditWarningsBadge) {
                    if (warnCount > 0) {
                        auditWarningsBadge.textContent = `${warnCount} warnings`;
                        auditWarningsBadge.classList.remove('hidden');
                    } else {
                        auditWarningsBadge.classList.add('hidden');
                    }
                }

                const rawIssues = Array.isArray(data.issues) ? data.issues : [
                    ...(Array.isArray(data.errors) ? data.errors.map(e => ({ ...e, level: 'error' })) : []),
                    ...(Array.isArray(data.warnings) ? data.warnings.map(w => ({ ...w, level: 'warning' })) : []),
                    ...(Array.isArray(data.infos) ? data.infos.map(i => ({ ...i, level: 'info' })) : [])
                ];

                if (auditIssuesList) {
                    auditIssuesList.innerHTML = '';
                    if (rawIssues.length > 0) {
                        rawIssues.forEach(issue => {
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
                }

                const fixedContent = data.fixed_csv_content ?? data.csv_content;
                if (fixedContent) {
                    auditFixedCsvContent = fixedContent;
                    auditFixedFilename = data.filename ? `repaired_${data.filename}` : 'definition_repaired.csv';
                    if (auditDownloadFixedBtn) auditDownloadFixedBtn.classList.remove('hidden');
                    if (auditCopyFixedBtn) auditCopyFixedBtn.classList.remove('hidden');
                } else {
                    auditFixedCsvContent = null;
                    if (auditDownloadFixedBtn) auditDownloadFixedBtn.classList.add('hidden');
                    if (auditCopyFixedBtn) auditCopyFixedBtn.classList.add('hidden');
                }

                if (auditResultsSection) auditResultsSection.classList.remove('hidden');
            } catch (err) {
                showToast(`Network or Server Error: ${err.message}`, 'error');
            } finally {
                auditProcessBtn.disabled = false;
                auditProcessBtn.textContent = 'Audit Definition';
            }
        });
    }

    if (auditDownloadFixedBtn && auditDownloadFixedBtn.addEventListener) {
        auditDownloadFixedBtn.addEventListener('click', () => {
            if (!auditFixedCsvContent) return;
            triggerDownload(auditFixedCsvContent, auditFixedFilename);
        });
    }

    if (auditCopyFixedBtn && auditCopyFixedBtn.addEventListener) {
        auditCopyFixedBtn.addEventListener('click', () => {
            if (!auditFixedCsvContent) return;
            copyToClipboard(auditFixedCsvContent, 'Repaired CSV copied to clipboard!');
        });
    }

    // =========================================================================
    // TAB 4: Equipment Library
    // =========================================================================
    const libraryTableBody = document.getElementById('libraryTableBody');
    const libraryCountBadge = document.getElementById('libraryCountBadge');
    const librarySearchInput = document.getElementById('librarySearchInput');
    let libraryItems = [];

    async function loadLibrary() {
        if (!libraryTableBody) return;
        try {
            const response = await fetch('/api/library');
            if (!response.ok) {
                libraryTableBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #ef4444;">Failed to load definitions catalog.</td></tr>`;
                return;
            }
            const data = await response.json();
            libraryItems = Array.isArray(data) ? data : (data.definitions || []);
            libraryLoaded = true;
            const modelCount = data.count ?? libraryItems.length;
            if (libraryCountBadge) libraryCountBadge.textContent = `${modelCount} models`;
            renderLibraryTable(libraryItems);
        } catch (err) {
            libraryTableBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #ef4444;">Network error: ${escapeHtml(err.message)}</td></tr>`;
        }
    }

    function renderLibraryTable(items) {
        if (!libraryTableBody) return;
        libraryTableBody.innerHTML = '';
        if (items.length === 0) {
            libraryTableBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-secondary);">No matching definitions found.</td></tr>`;
            return;
        }

        items.forEach(item => {
            const tr = document.createElement('tr');
            const sizeKb = item.size_kb != null ? item.size_kb : (item.size_bytes ? (item.size_bytes / 1024).toFixed(1) : '0.0');
            tr.innerHTML = `
                <td><span class="badge badge-info">${escapeHtml(item.category || '')}</span></td>
                <td><strong>${escapeHtml(item.manufacturer || '')}</strong></td>
                <td><code>${escapeHtml(item.model || '')}</code></td>
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

    if (librarySearchInput && librarySearchInput.addEventListener) {
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
    }
});
