document.addEventListener('DOMContentLoaded', () => {
    // =========================================================================
    // Tab Navigation & Keyboard Accessibility
    // =========================================================================
    let libraryLoaded = false;
    let auditTrailLoaded = false;
    let versionsLoaded = false;
    let duplicatesLoaded = false;
    let currentLibrarySource = 'production';
    let currentCategoryFilter = '';

    const tabBindings = [
        { btnId: 'tabBtnConvert', paneId: 'tab-convert' },
        { btnId: 'tabBtnDiscovery', paneId: 'tab-discovery' },
        { btnId: 'tabBtnTag', paneId: 'tab-tag' },
        { btnId: 'tabBtnAudit', paneId: 'tab-audit' },
        {
            btnId: 'tabBtnLibrary',
            paneId: 'tab-library',
            onActivate: () => {
                if (!libraryLoaded) loadLibrary();
            }
        },
        {
            btnId: 'tabBtnVersions',
            paneId: 'tab-versions',
            onActivate: () => {
                if (!versionsLoaded) loadVersions();
            }
        },
        {
            btnId: 'tabBtnDuplicates',
            paneId: 'tab-duplicates',
            onActivate: () => {
                if (!duplicatesLoaded) loadDuplicates();
            }
        },
        {
            btnId: 'tabBtnAuditTrail',
            paneId: 'tab-audit-trail',
            onActivate: () => {
                if (!auditTrailLoaded) loadAuditTrail();
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
                    try { container.removeChild(toast); } catch {}
                }
            }, 3500);
        }
        if (type === 'error' && typeof alert === 'function' && typeof window === 'undefined') {
            alert(message);
        }
    }

    async function copyToClipboard(text, successMsg = 'Copié dans le presse-papier !') {
        if (!text) return;
        try {
            if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
                await navigator.clipboard.writeText(text);
            }
            showToast(successMsg, 'success');
        } catch {
            showToast('Impossible de copier dans le presse-papier', 'error');
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
                const text = (row.textContent || row.innerText || '').toLowerCase();
                const visible = !q || text.includes(q);
                if (row.style) row.style.display = visible ? '' : 'none';
                if (row.classList) row.classList.toggle('hidden', !visible);
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

    // Load Knowledge Base Stats on startup
    async function loadKnowledgeStats() {
        const pill = document.getElementById('knowledgePill');
        if (!pill) return;
        try {
            const res = await fetch('/api/knowledge-stats');
            if (res.ok) {
                const data = await res.json();
                if (data && data.total_mappings) {
                    pill.textContent = `🧠 ${data.total_mappings.toLocaleString()} variables apprises`;
                }
            }
        } catch {}
    }
    loadKnowledgeStats();

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

    let selectedFile = null;
    let convertedCsvContent = null;
    let convertedFilename = 'definition.csv';
    let requestVersion = 0;
    let pendingController = null;

    function resetResults() {
        resultsSection.classList.add('hidden');
        downloadBtn.disabled = true;
        if (copyBtn) copyBtn.disabled = true;
        convertedCsvContent = null;
        const errBox = document.getElementById('convertErrorBox');
        if (errBox) errBox.classList.add('hidden');
    }

    function invalidateConversion() {
        requestVersion++;
        if (pendingController) {
            pendingController.abort();
            pendingController = null;
        }
        resetResults();
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
        resultsSection.classList.add('hidden');
    });

    if (removeFileBtn && removeFileBtn.addEventListener) {
        removeFileBtn.addEventListener('click', () => {
            selectedFile = null;
            invalidateConversion();
            if (fileInput) fileInput.value = '';
            dropZone.classList.remove('hidden');
            fileDetails.classList.add('hidden');
            convertBtn.disabled = true;
        });
    }

    if (convertBtn && convertBtn.addEventListener) {
        convertBtn.addEventListener('click', async () => {
            if (!selectedFile) return;

            const mfg = (document.getElementById('mfgInput').value || '').trim() || 'Manufacturer';
            const model = (document.getElementById('modelInput').value || '').trim() || 'Model';
            const protocol = document.getElementById('protocolSelect').value || 'Modbus';
            const category = document.getElementById('categorySelect').value || 'Inverter';
            const offset = document.getElementById('offsetInput').value || '0';
            const forcedWrite = (document.getElementById('forcedWriteInput').value || '').trim();
            const autoTagCheck = document.getElementById('autoTagCheck');
            const doTag = autoTagCheck ? autoTagCheck.checked : true;

            const convertErrorBox = document.getElementById('convertErrorBox');
            const convertErrorMessage = document.getElementById('convertErrorMessage');
            if (convertErrorBox) convertErrorBox.classList.add('hidden');

            const formData = new FormData();
            formData.append('file', selectedFile);
            formData.append('manufacturer', mfg);
            formData.append('model', model);
            formData.append('protocol', protocol);
            formData.append('category', category);
            formData.append('address_offset', offset);
            formData.append('tag', doTag);
            if (forcedWrite) formData.append('forced_write', forcedWrite);

            const convertLabel = convertBtn.innerHTML;
            convertBtn.disabled = true;
            convertBtn.textContent = 'Extraction & Conversion en cours...';

            if (pendingController) pendingController.abort();
            pendingController = new AbortController();
            const version = ++requestVersion;

            try {
                const response = await fetch('/api/convert', {
                    method: 'POST',
                    body: formData,
                    signal: pendingController.signal
                });

                if (version !== requestVersion) return;

                if (!response.ok) {
                    let errorDetail = 'Error during conversion';
                    try {
                        const err = await response.json();
                        errorDetail = err.detail || errorDetail;
                    } catch (_) {
                        errorDetail = `HTTP ${response.status} Error: ${response.statusText}`;
                    }
                    showToast(errorDetail, 'error');
                    if (convertErrorBox && convertErrorMessage) {
                        convertErrorMessage.textContent = errorDetail;
                        convertErrorBox.classList.remove('hidden');
                    }
                    resultsSection.classList.add('hidden');
                    downloadBtn.disabled = true;
                    if (copyBtn) copyBtn.disabled = true;
                    return;
                }

                const data = await response.json();
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
                    const msg = `Network or server error: ${err.message}`;
                    showToast(msg, 'error');
                    if (convertErrorBox && convertErrorMessage) {
                        convertErrorMessage.textContent = msg;
                        convertErrorBox.classList.remove('hidden');
                    }
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
        if (!previewBody) return;
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
            copyToClipboard(convertedCsvContent, 'Définition CSV copiée dans le presse-papier !');
        });
    }
    setupTableFilter(previewFilterInput, previewBody);

    // =========================================================================
    // TAB: Autonomous Modbus Discovery (Phase B)
    // =========================================================================
    const discoveryQueryInput = document.getElementById('discoveryQueryInput');
    const discoveryTargetSelect = document.getElementById('discoveryTargetSelect');
    const discoveryRunBtn = document.getElementById('discoveryRunBtn');
    const discoveryEnableOcr = document.getElementById('discoveryEnableOcr');
    const discoveryManualFileInput = document.getElementById('discoveryManualFileInput');
    const discoveryStatusBox = document.getElementById('discoveryStatusBox');
    const discoveryStatusTitle = document.getElementById('discoveryStatusTitle');
    const discoveryStatusBadge = document.getElementById('discoveryStatusBadge');
    const discoveryStatusDetail = document.getElementById('discoveryStatusDetail');
    const discoveryResultsSection = document.getElementById('discoveryResultsSection');
    const discoveryRegistersBadge = document.getElementById('discoveryRegistersBadge');
    const discoveryPillarsBadge = document.getElementById('discoveryPillarsBadge');
    const discoveryCopyBtn = document.getElementById('discoveryCopyBtn');
    const discoveryDownloadBtn = document.getElementById('discoveryDownloadBtn');
    const discoveryEvidenceBtn = document.getElementById('discoveryEvidenceBtn');
    const discoveryDocMeta = document.getElementById('discoveryDocMeta');
    const discoveryPreviewBody = document.getElementById('discoveryPreviewBody');

    let currentDiscoveryCsv = '';
    let currentDiscoveryFilename = 'discovered_modbus_definition.csv';
    let currentDiscoveryEvidence = '';

    function updateDiscoveryStatus(title, badgeText, badgeClass, detailText) {
        if (!discoveryStatusBox) return;
        if (discoveryStatusBox.classList) discoveryStatusBox.classList.remove('hidden');
        if (discoveryStatusTitle) discoveryStatusTitle.textContent = title;
        if (discoveryStatusBadge) {
            discoveryStatusBadge.textContent = badgeText;
            discoveryStatusBadge.className = `badge ${badgeClass}`;
        }
        if (discoveryStatusDetail) discoveryStatusDetail.textContent = detailText;
    }

    function renderDiscoveryResult(data) {
        if (!data) return;
        currentDiscoveryCsv = data.csv_content || '';
        currentDiscoveryEvidence = data.evidence_report || '';

        const mfg = data.identity?.manufacturer || data.model?.manufacturer || 'Generic';
        const mdl = data.identity?.exact_model || data.model?.model || 'Equipment';
        const cleanMfg = mfg.replace(/[^a-zA-Z0-9_-]/g, '_');
        const cleanMdl = mdl.replace(/[^a-zA-Z0-9_-]/g, '_');
        const cat = data.model?.category || '';
        const prefix = (cat === 'Meter') ? 'MET' : (cat === 'Battery') ? 'BAT' : (cat === 'Logger') ? 'COM' : 'INV';
        currentDiscoveryFilename = `${prefix}_${cleanMfg}_${cleanMdl}.csv`;

        const regCount = data.model?.total_registers || data.model?.registers?.length || 0;
        if (discoveryRegistersBadge) {
            discoveryRegistersBadge.textContent = `${regCount} registres`;
        }

        const isCompliant = data.validation?.status === 'VALIDATED_OFFLINE';
        if (discoveryPillarsBadge) {
            if (isCompliant) {
                discoveryPillarsBadge.textContent = '5 Piliers Conformes';
                discoveryPillarsBadge.className = 'badge badge-success';
            } else {
                discoveryPillarsBadge.textContent = data.validation?.status || 'Validation Partielle';
                discoveryPillarsBadge.className = 'badge badge-warning';
            }
        }

        if (discoveryDocMeta) {
            const docTitle = data.document?.title || data.fetch?.final_url || 'Fichier Source Local';
            const sha256 = data.fetch?.sha256 ? data.fetch.sha256.substring(0, 16) + '...' : '-';
            const confidence = data.identity?.confidence_score ? `${Math.round(data.identity.confidence_score * 100)}%` : '100%';
            discoveryDocMeta.innerHTML = `
                <div><strong>Source certifiée :</strong> ${escapeHtml(docTitle)}</div>
                <div style="display: flex; gap: 16px; margin-top: 4px; color: var(--text-secondary);">
                    <span><strong>Modèle :</strong> ${escapeHtml(mfg)} ${escapeHtml(mdl)}</span>
                    <span><strong>Confiance :</strong> ${escapeHtml(confidence)}</span>
                    <span><strong>SHA-256 :</strong> <code>${escapeHtml(sha256)}</code></span>
                </div>
            `;
        }

        if (discoveryPreviewBody) {
            discoveryPreviewBody.innerHTML = '';
            const registers = data.model?.registers || [];
            const displayLimit = Math.min(registers.length, 300);
            for (let i = 0; i < displayLimit; i++) {
                const reg = registers[i];
                const tr = document.createElement('tr');
                const tagDisplay = reg.canonical_sunpm_tag ? `<code>${escapeHtml(reg.canonical_sunpm_tag)}</code>` : '<span style="color:var(--text-secondary);">-</span>';
                tr.innerHTML = `
                    <td><code>${escapeHtml(reg.address)}</code></td>
                    <td>${escapeHtml(reg.name)}</td>
                    <td>${tagDisplay}</td>
                    <td>${escapeHtml(reg.data_type || 'U16')}</td>
                    <td>${escapeHtml(reg.factor ?? 1.0)}</td>
                    <td>${escapeHtml(reg.unit || '-')}</td>
                    <td>${escapeHtml(reg.action_code || 4)}</td>
                `;
                discoveryPreviewBody.appendChild(tr);
            }
        }

        if (discoveryResultsSection && discoveryResultsSection.classList) {
            discoveryResultsSection.classList.remove('hidden');
        }
    }

    const discoveryQuickPills = document.getElementById('discoveryQuickPills');
    if (discoveryQuickPills && discoveryQuickPills.addEventListener) {
        discoveryQuickPills.addEventListener('click', (e) => {
            const btn = e.target.closest ? e.target.closest('.pill-btn') : null;
            if (!btn) return;
            const q = btn.getAttribute ? btn.getAttribute('data-query') : null;
            if (q && discoveryQueryInput) {
                discoveryQueryInput.value = q;
                if (discoveryQueryInput.focus) discoveryQueryInput.focus();
            }
        });
    }

    if (discoveryQueryInput && discoveryQueryInput.addEventListener) {
        discoveryQueryInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                if (e.preventDefault) e.preventDefault();
                if (discoveryRunBtn && discoveryRunBtn.click) discoveryRunBtn.click();
            }
        });
    }

    if (discoveryRunBtn && discoveryRunBtn.addEventListener) {
        discoveryRunBtn.addEventListener('click', async () => {
            const query = discoveryQueryInput ? discoveryQueryInput.value.trim() : '';
            if (!query) {
                showToast('Veuillez saisir une référence équipement.', 'error');
                if (discoveryQueryInput && discoveryQueryInput.focus) discoveryQueryInput.focus();
                return;
            }

            const targetFormat = discoveryTargetSelect ? discoveryTargetSelect.value : 'webdyn_sunpm';
            const ocrFallback = discoveryEnableOcr ? discoveryEnableOcr.checked : true;

            discoveryRunBtn.disabled = true;
            if (discoveryResultsSection && discoveryResultsSection.classList) {
                discoveryResultsSection.classList.add('hidden');
            }
            updateDiscoveryStatus('Discovery in progress', 'In progress...', 'badge-info', `Searching official vendor documentation for "${query}"...`);

            try {
                const res = await fetch('/api/discovery/lookup', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        query: query,
                        target_format: targetFormat,
                        ocr_fallback: ocrFallback
                    })
                });

                if (!res.ok) {
                    const err = await res.json().catch(() => ({ detail: 'HTTP Error ' + res.status }));
                    updateDiscoveryStatus('Discovery failed', 'Error', 'badge-error', err.detail || 'Error during discovery query.');
                    showToast(err.detail || 'Error during discovery', 'error');
                    return;
                }

                const data = await res.json();
                if (data.success) {
                    updateDiscoveryStatus(
                        'Discovery & Certification complete',
                        'Certified',
                        'badge-success',
                        `Definition successfully generated for ${data.identity?.manufacturer || ''} ${data.identity?.exact_model || ''} (${data.model?.total_registers || 0} certified registers).`
                    );
                    renderDiscoveryResult(data);
                    showToast('Modbus definition discovered and certified successfully!', 'success');
                } else {
                    const errMsg = data.error_message || 'Could not extract certified register table for this reference.';
                    updateDiscoveryStatus('Attention required', data.status || 'Unresolved', 'badge-warning', errMsg);
                    if (data.model && data.model.registers && data.model.registers.length > 0) {
                        renderDiscoveryResult(data);
                    }
                    showToast(errMsg, 'error');
                }
            } catch (err) {
                updateDiscoveryStatus('Connection error', 'Error', 'badge-error', `Network error: ${err.message}`);
                showToast(`Network error: ${err.message}`, 'error');
            } finally {
                discoveryRunBtn.disabled = false;
            }
        });
    }

    if (discoveryManualFileInput && discoveryManualFileInput.addEventListener) {
        discoveryManualFileInput.addEventListener('change', async (e) => {
            const files = e.target ? e.target.files : null;
            if (!files || files.length === 0) return;
            const file = files[0];

            const targetFormat = discoveryTargetSelect ? discoveryTargetSelect.value : 'webdyn_sunpm';
            const ocrFallback = discoveryEnableOcr ? discoveryEnableOcr.checked : true;

            const formData = new FormData();
            formData.append('file', file);
            formData.append('target_format', targetFormat);
            formData.append('ocr_fallback', ocrFallback ? 'true' : 'false');

            if (discoveryRunBtn) discoveryRunBtn.disabled = true;
            if (discoveryResultsSection && discoveryResultsSection.classList) {
                discoveryResultsSection.classList.add('hidden');
            }
            updateDiscoveryStatus('Local extraction in progress', 'In progress...', 'badge-info', `Analyzing and normalizing fallback file "${file.name}"...`);

            try {
                const res = await fetch('/api/discovery/upload', {
                    method: 'POST',
                    body: formData
                });

                if (!res.ok) {
                    const err = await res.json().catch(() => ({ detail: 'HTTP Error ' + res.status }));
                    updateDiscoveryStatus('Extraction failed', 'Error', 'badge-error', err.detail || 'Error processing file.');
                    showToast(err.detail || 'Error during processing', 'error');
                    return;
                }

                const data = await res.json();
                if (data.success) {
                    updateDiscoveryStatus(
                        'Extraction & Certification complete',
                        'Certified',
                        'badge-success',
                        `Definition extracted successfully from ${file.name} (${data.model?.total_registers || 0} registers).`
                    );
                    renderDiscoveryResult(data);
                    showToast('Definition extracted and certified successfully!', 'success');
                } else {
                    const errMsg = data.error_message || 'Validation failed on provided document.';
                    updateDiscoveryStatus('Attention required', data.status || 'Unresolved', 'badge-warning', errMsg);
                    if (data.model && data.model.registers && data.model.registers.length > 0) {
                        renderDiscoveryResult(data);
                    }
                    showToast(errMsg, 'error');
                }
            } catch (err) {
                updateDiscoveryStatus('Processing error', 'Error', 'badge-error', `Network error: ${err.message}`);
                showToast(`Network error: ${err.message}`, 'error');
            } finally {
                if (discoveryRunBtn) discoveryRunBtn.disabled = false;
                if (discoveryManualFileInput) discoveryManualFileInput.value = '';
            }
        });
    }

    if (discoveryCopyBtn && discoveryCopyBtn.addEventListener) {
        discoveryCopyBtn.addEventListener('click', () => {
            if (!currentDiscoveryCsv) return;
            copyToClipboard(currentDiscoveryCsv, 'Définition CSV copiée dans le presse-papier !');
        });
    }

    if (discoveryDownloadBtn && discoveryDownloadBtn.addEventListener) {
        discoveryDownloadBtn.addEventListener('click', () => {
            if (!currentDiscoveryCsv) return;
            triggerDownload(currentDiscoveryCsv, currentDiscoveryFilename);
        });
    }

    if (discoveryEvidenceBtn && discoveryEvidenceBtn.addEventListener) {
        discoveryEvidenceBtn.addEventListener('click', () => {
            if (!currentDiscoveryEvidence) {
                showToast('Aucun rapport de preuve disponible.', 'info');
                return;
            }
            const blob = new Blob([currentDiscoveryEvidence], { type: 'text/markdown;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', `${currentDiscoveryFilename.replace('.csv', '')}_EVIDENCE.md`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            showToast('Rapport de traçabilité cryptographique téléchargé !', 'success');
        });
    }

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
        const errBox = document.getElementById('tagErrorBox');
        if (errBox) errBox.classList.add('hidden');
    });

    if (tagRemoveFileBtn && tagRemoveFileBtn.addEventListener) {
        tagRemoveFileBtn.addEventListener('click', () => {
            tagSelectedFile = null;
            if (tagFileInput) tagFileInput.value = '';
            tagDropZone.classList.remove('hidden');
            tagFileDetails.classList.add('hidden');
            tagProcessBtn.disabled = true;
            tagResultsSection.classList.add('hidden');
            const errBox = document.getElementById('tagErrorBox');
            if (errBox) errBox.classList.add('hidden');
            taggedCsvContent = null;
        });
    }

    if (tagProcessBtn && tagProcessBtn.addEventListener) {
        tagProcessBtn.addEventListener('click', async () => {
            if (!tagSelectedFile) return;

            const category = tagCategorySelect ? tagCategorySelect.value : '';
            const overwrite = tagOverwriteCheck ? tagOverwriteCheck.checked : false;

            const formData = new FormData();
            formData.append('file', tagSelectedFile);
            if (category) formData.append('category', category);
            formData.append('overwrite', overwrite ? 'true' : 'false');
            formData.append('normalize_legacy', 'true');

            const tagErrorBox = document.getElementById('tagErrorBox');
            const tagErrorMessage = document.getElementById('tagErrorMessage');
            if (tagErrorBox) tagErrorBox.classList.add('hidden');

            tagProcessBtn.disabled = true;
            tagProcessBtn.textContent = 'Tagging SunPM registers...';

            try {
                const response = await fetch('/api/tag', {
                    method: 'POST',
                    body: formData,
                });

                if (!response.ok) {
                    let errDetail = 'Error during tagging';
                    try {
                        const err = await response.json();
                        errDetail = err.detail || errDetail;
                    } catch (_) {
                        errDetail = `HTTP Error ${response.status}: ${response.statusText}`;
                    }
                    showToast(errDetail, 'error');
                    if (tagErrorBox && tagErrorMessage) {
                        tagErrorMessage.textContent = errDetail;
                        tagErrorBox.classList.remove('hidden');
                    }
                    return;
                }

                const data = await response.json();
                taggedCsvContent = data.csv_content;
                taggedFilename = data.filename;

                tagCountBadge.textContent = `${data.tags_updated ?? data.updated_count ?? 0} tags updated`;
                tagRegisterBadge.textContent = `${data.register_count ?? data.total_registers ?? 0} registers`;

                renderTagPreview(data.preview || []);
                tagResultsSection.classList.remove('hidden');
                showToast(`Tagging complete: ${data.tags_updated ?? 0} tags updated!`, 'success');
            } catch (err) {
                const msg = `Network error: ${err.message}`;
                showToast(msg, 'error');
                if (tagErrorBox && tagErrorMessage) {
                    tagErrorMessage.textContent = msg;
                    tagErrorBox.classList.remove('hidden');
                }
            } finally {
                tagProcessBtn.disabled = false;
                tagProcessBtn.textContent = 'Apply Canonical SunPM Tags';
            }
        });
    }

    function renderTagPreview(rows) {
        if (!tagPreviewBody) return;
        tagPreviewBody.innerHTML = '';
        rows.forEach(r => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${escapeHtml(r.Name || '')}</td>
                <td><code>${escapeHtml(r.Tag || '')}</code></td>
                <td><code>${escapeHtml(r.Address ?? '')}</code></td>
                <td>${escapeHtml(r.Type || '')}</td>
                <td>${escapeHtml(r.Unit || '')}</td>
                <td>${escapeHtml(r.Action || '4')}</td>
            `;
            tagPreviewBody.appendChild(tr);
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
            copyToClipboard(taggedCsvContent, 'Tagged definition copied!');
        });
    }
    setupTableFilter(tagFilterInput, tagPreviewBody);

    // =========================================================================
    // TAB 3: Validator & Auditor (5 Piliers Apollo)
    // =========================================================================
    const auditDropZone = document.getElementById('auditDropZone');
    const auditFileInput = document.getElementById('auditFileInput');
    const auditFileDetails = document.getElementById('auditFileDetails');
    const auditFileName = document.getElementById('auditFileName');
    const auditRemoveFileBtn = document.getElementById('auditRemoveFileBtn');
    const auditProcessBtn = document.getElementById('auditProcessBtn');
    const auditPillarsBtn = document.getElementById('auditPillarsBtn');
    const auditResultsSection = document.getElementById('auditResultsSection');
    const pillarsResultsSection = document.getElementById('pillarsResultsSection');
    const pillarsGrid = document.getElementById('pillarsGrid');
    const pillarsIssuesList = document.getElementById('pillarsIssuesList');
    const pillarsStatusBadge = document.getElementById('pillarsStatusBadge');
    const auditStatusBadge = document.getElementById('auditStatusBadge');
    const auditRegistersBadge = document.getElementById('auditRegistersBadge');
    const auditErrorsBadge = document.getElementById('auditErrorsBadge');
    const auditWarningsBadge = document.getElementById('auditWarningsBadge');
    const auditIssuesList = document.getElementById('auditIssuesList');
    const auditDownloadFixedBtn = document.getElementById('auditDownloadFixedBtn');
    const auditCopyFixedBtn = document.getElementById('auditCopyFixedBtn');
    const auditStrictOverlapCheck = document.getElementById('auditStrictOverlapCheck');
    const auditAutoFixCheck = document.getElementById('auditAutoFixCheck');
    const auditErrorBox = document.getElementById('auditErrorBox');
    const auditErrorMessage = document.getElementById('auditErrorMessage');

    let auditSelectedFile = null;
    let auditFixedCsvContent = null;
    let auditFixedFilename = 'definition_repaired.csv';

    setupDropZone(auditDropZone, auditFileInput, (file) => {
        auditSelectedFile = file;
        auditFileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        auditDropZone.classList.add('hidden');
        auditFileDetails.classList.remove('hidden');
        auditProcessBtn.disabled = false;
        if (auditPillarsBtn) auditPillarsBtn.disabled = false;
        if (auditResultsSection) auditResultsSection.classList.add('hidden');
        if (pillarsResultsSection) pillarsResultsSection.classList.add('hidden');
        if (auditErrorBox) auditErrorBox.classList.add('hidden');
    });

    if (auditRemoveFileBtn && auditRemoveFileBtn.addEventListener) {
        auditRemoveFileBtn.addEventListener('click', () => {
            auditSelectedFile = null;
            if (auditFileInput) auditFileInput.value = '';
            auditDropZone.classList.remove('hidden');
            auditFileDetails.classList.add('hidden');
            auditProcessBtn.disabled = true;
            if (auditPillarsBtn) auditPillarsBtn.disabled = true;
            if (auditResultsSection) auditResultsSection.classList.add('hidden');
            if (pillarsResultsSection) pillarsResultsSection.classList.add('hidden');
            if (auditErrorBox) auditErrorBox.classList.add('hidden');
            auditFixedCsvContent = null;
        });
    }

    // 5 PILLARS CHECK HANDLER
    if (auditPillarsBtn && auditPillarsBtn.addEventListener) {
        auditPillarsBtn.addEventListener('click', async () => {
            if (!auditSelectedFile) return;

            const formData = new FormData();
            formData.append('file', auditSelectedFile);

            if (auditErrorBox) auditErrorBox.classList.add('hidden');
            auditPillarsBtn.disabled = true;
            auditPillarsBtn.textContent = 'Verifying 5 Pillars...';

            try {
                const response = await fetch('/api/pillars', {
                    method: 'POST',
                    body: formData,
                });

                if (!response.ok) {
                    let errDetail = 'Error verifying 5 Pillars';
                    try {
                        const err = await response.json();
                        errDetail = err.detail || errDetail;
                    } catch (_) {
                        errDetail = `HTTP Error ${response.status}: ${response.statusText}`;
                    }
                    showToast(errDetail, 'error');
                    if (auditErrorBox && auditErrorMessage) {
                        auditErrorMessage.textContent = errDetail;
                        auditErrorBox.classList.remove('hidden');
                    }
                    return;
                }

                const data = await response.json();
                renderPillarsReport(data);
                if (pillarsResultsSection) pillarsResultsSection.classList.remove('hidden');
            } catch (err) {
                const msg = `Network error: ${err.message}`;
                showToast(msg, 'error');
                if (auditErrorBox && auditErrorMessage) {
                    auditErrorMessage.textContent = msg;
                    auditErrorBox.classList.remove('hidden');
                }
            } finally {
                auditPillarsBtn.disabled = false;
                auditPillarsBtn.textContent = 'Verify Apollo 5 Pillars';
            }
        });
    }

    function renderPillarsReport(data) {
        if (!pillarsGrid) return;
        pillarsGrid.innerHTML = '';

        const pillarsMeta = [
            { id: 'naming', title: '1. Standard Naming', desc: '<PREFIX>_<Manufacturer>_<Model>.csv' },
            { id: 'addressing', title: '2. Modbus Addressing', desc: 'Exact decimal 0..65535, STRING suffixes' },
            { id: 'function_codes', title: '3. Function Codes', desc: 'Modbus function codes 1, 2, 3, 4' },
            { id: 'action_codes', title: '4. Action Codes', desc: 'Webdyn action codes 0, 1, 2, 4, 8, 9, 10' },
            { id: 'tags', title: '5. Canonical Tags', desc: 'SunPM tags without physical duplicates' },
            { id: 'encoding_schema', title: '6. Schema & UTF-8', desc: 'Strict UTF-8 without BOM, 11 columns' },
        ];

        pillarsMeta.forEach(p => {
            const pass = data.pillars && data.pillars[p.id] !== false;
            const card = document.createElement('div');
            card.className = `pillar-card ${pass ? 'pass' : 'fail'}`;
            card.innerHTML = `
                <div class="pillar-title">${escapeHtml(p.title)}</div>
                <div class="pillar-desc">${escapeHtml(p.desc)}</div>
                <span class="pillar-badge">${pass ? '✓ COMPLIANT' : '✗ NON-COMPLIANT'}</span>
            `;
            pillarsGrid.appendChild(card);
        });

        if (pillarsStatusBadge) {
            pillarsStatusBadge.textContent = data.is_valid ? '100% 5-Pillar Compliant (PASS)' : 'Non-Compliant (Failures detected)';
            pillarsStatusBadge.className = `badge ${data.is_valid ? 'badge-success' : 'badge-error'}`;
        }

        if (pillarsIssuesList) {
            pillarsIssuesList.innerHTML = '';
            if (data.issues && data.issues.length) {
                data.issues.forEach(iss => {
                    const div = document.createElement('div');
                    div.className = 'audit-issue-card error';
                    div.innerHTML = `<span class="issue-level">PILLAR</span> <span>${escapeHtml(iss)}</span>`;
                    pillarsIssuesList.appendChild(div);
                });
            } else {
                const div = document.createElement('div');
                div.className = 'audit-issue-card info';
                div.innerHTML = `<span class="issue-level">INFO</span> <span>All Apollo 5-Pillar rules are strictly satisfied.</span>`;
                pillarsIssuesList.appendChild(div);
            }
        }
    }

    // AUDIT & OVERLAP PROCESS HANDLER
    if (auditProcessBtn && auditProcessBtn.addEventListener) {
        auditProcessBtn.addEventListener('click', async () => {
            if (!auditSelectedFile) return;

            const strictOverlap = auditStrictOverlapCheck ? auditStrictOverlapCheck.checked : true;
            const autoFix = auditAutoFixCheck ? auditAutoFixCheck.checked : false;

            const formData = new FormData();
            formData.append('file', auditSelectedFile);
            formData.append('strict_overlap', strictOverlap ? 'true' : 'false');
            formData.append('auto_fix', autoFix ? 'true' : 'false');

            if (auditErrorBox) auditErrorBox.classList.add('hidden');
            auditProcessBtn.disabled = true;
            auditProcessBtn.textContent = 'Auditing in progress...';

            try {
                const response = await fetch('/api/audit', {
                    method: 'POST',
                    body: formData,
                });

                if (!response.ok) {
                    let errDetail = 'Error during audit';
                    try {
                        const err = await response.json();
                        errDetail = err.detail || errDetail;
                    } catch (_) {
                        errDetail = `HTTP Error ${response.status}: ${response.statusText}`;
                    }
                    showToast(errDetail, 'error');
                    if (auditErrorBox && auditErrorMessage) {
                        auditErrorMessage.textContent = errDetail;
                        auditErrorBox.classList.remove('hidden');
                    }
                    return;
                }

                const data = await response.json();
                auditStatusBadge.textContent = data.valid ? 'Valid' : 'Issues Found';
                auditStatusBadge.className = `badge ${data.valid ? 'badge-success' : 'badge-error'}`;
                auditRegistersBadge.textContent = `${data.register_count || 0} registers`;

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

                auditIssuesList.innerHTML = '';
                const issues = data.issues || [];
                if (issues.length === 0) {
                    auditIssuesList.innerHTML = `<div class="audit-issue-card info"><span class="issue-level">INFO</span><span>No syntax issues or address collisions detected.</span></div>`;
                } else {
                    issues.forEach(iss => {
                        const level = iss.level || 'info';
                        const card = document.createElement('div');
                        card.className = `audit-issue-card ${level}`;
                        const lineText = iss.line ? `Line ${iss.line}: ` : '';
                        card.innerHTML = `<span class="issue-level">${level}</span> <span><strong>${escapeHtml(lineText)}</strong>${escapeHtml(iss.message || '')}</span>`;
                        auditIssuesList.appendChild(card);
                    });
                }

                if (data.fixed_csv_content) {
                    auditFixedCsvContent = data.fixed_csv_content;
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
                const msg = `Server error: ${err.message}`;
                showToast(msg, 'error');
                if (auditErrorBox && auditErrorMessage) {
                    auditErrorMessage.textContent = msg;
                    auditErrorBox.classList.remove('hidden');
                }
            } finally {
                auditProcessBtn.disabled = false;
                auditProcessBtn.textContent = 'Run Syntax & Overlap Audit';
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
            copyToClipboard(auditFixedCsvContent, 'Repaired definition copied!');
        });
    }

    // =========================================================================
    // TAB 4: Production & Equipment Library
    // =========================================================================
    const libraryTableBody = document.getElementById('libraryTableBody');
    const libraryCountBadge = document.getElementById('libraryCountBadge');
    const librarySearchInput = document.getElementById('librarySearchInput');
    const sourceProductionBtn = document.getElementById('sourceProductionBtn');
    const sourceDefinitionsBtn = document.getElementById('sourceDefinitionsBtn');
    const categoryPills = document.getElementById('categoryPills');
    let libraryItems = [];

    async function loadLibrary(source = currentLibrarySource) {
        if (!libraryTableBody) return;
        currentLibrarySource = source;

        if (sourceProductionBtn && sourceDefinitionsBtn) {
            sourceProductionBtn.className = `btn btn-sm ${source === 'production' ? 'btn-primary' : 'btn-secondary'}`;
            sourceDefinitionsBtn.className = `btn btn-sm ${source === 'definitions' ? 'btn-primary' : 'btn-secondary'}`;
        }

        try {
            const url = `/api/library?source=${encodeURIComponent(source)}${currentCategoryFilter ? `&category=${encodeURIComponent(currentCategoryFilter)}` : ''}`;
            const response = await fetch(url);
            if (!response.ok) {
                libraryTableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: #ef4444;">Failed to load catalog.</td></tr>`;
                return;
            }
            const data = await response.json();
            libraryItems = Array.isArray(data) ? data : (data.definitions || []);
            libraryLoaded = true;
            const modelCount = data.count ?? libraryItems.length;
            if (libraryCountBadge) {
                libraryCountBadge.textContent = `${modelCount} models`;
            }
            renderLibraryTable(libraryItems);
        } catch (err) {
            libraryTableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: #ef4444;">Network error: ${escapeHtml(err.message)}</td></tr>`;
        }
    }

    function renderLibraryTable(items) {
        if (!libraryTableBody) return;
        libraryTableBody.innerHTML = '';
        if (items.length === 0) {
            libraryTableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: var(--text-secondary);">No matching definitions found.</td></tr>`;
            return;
        }

        items.forEach(item => {
            const tr = document.createElement('tr');
            const sizeKb = item.size_kb != null ? item.size_kb : (item.size_bytes ? (item.size_bytes / 1024).toFixed(1) : '0.0');
            const regCount = item.registers || '-';
            tr.innerHTML = `
                <td><span class="badge badge-info">${escapeHtml(item.prefix || item.category || '')}</span></td>
                <td><strong>${escapeHtml(item.manufacturer || '')}</strong></td>
                <td><code>${escapeHtml(item.model || '')}</code></td>
                <td><span class="badge" style="background: rgba(255,255,255,0.08);">${regCount} regs</span></td>
                <td><span class="badge badge-success" style="font-size: 0.72rem;">✓ 5 Pillars PASS</span></td>
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

    if (sourceProductionBtn && sourceProductionBtn.addEventListener) {
        sourceProductionBtn.addEventListener('click', () => loadLibrary('production'));
    }
    if (sourceDefinitionsBtn && sourceDefinitionsBtn.addEventListener) {
        sourceDefinitionsBtn.addEventListener('click', () => loadLibrary('definitions'));
    }

    if (categoryPills && categoryPills.addEventListener) {
        categoryPills.addEventListener('click', (e) => {
            const btn = e.target.closest('.pill-btn');
            if (!btn) return;
            categoryPills.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentCategoryFilter = btn.getAttribute('data-cat') || '';
            loadLibrary(currentLibrarySource);
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
                       (item.prefix && item.prefix.toLowerCase().includes(query)) ||
                       (item.filename && item.filename.toLowerCase().includes(query));
            });
            renderLibraryTable(filtered);
        });
    }

    // =========================================================================
    // TAB 5: Immutable Audit Trail Viewer
    // =========================================================================
    const auditTrailBody = document.getElementById('auditTrailBody');
    const refreshAuditTrailBtn = document.getElementById('refreshAuditTrailBtn');
    const auditTrailSearchInput = document.getElementById('auditTrailSearchInput');
    let auditTrailItems = [];

    async function loadAuditTrail() {
        if (!auditTrailBody) return;
        try {
            const response = await fetch('/api/audit-trail?limit=100');
            if (!response.ok) {
                auditTrailBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: #ef4444;">Failed to load audit trail.</td></tr>`;
                return;
            }
            auditTrailItems = await response.json();
            auditTrailLoaded = true;
            renderAuditTrailTable(auditTrailItems);
        } catch (err) {
            auditTrailBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: #ef4444;">Network error: ${escapeHtml(err.message)}</td></tr>`;
        }
    }

    function renderAuditTrailTable(items) {
        if (!auditTrailBody) return;
        auditTrailBody.innerHTML = '';
        if (!items || items.length === 0) {
            auditTrailBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: var(--text-secondary);">No audit events recorded.</td></tr>`;
            return;
        }

        items.forEach(ev => {
            const tr = document.createElement('tr');
            const ts = ev.timestamp ? ev.timestamp.replace('T', ' ').substring(0, 19) : '-';
            const defs = (ev.definitions || []).map(d => `<code>${escapeHtml(d)}</code>`).join('<br>');
            const statusClass = (ev.validation_status || '').includes('PASS') ? 'badge-success' : 'badge-error';

            tr.innerHTML = `
                <td style="white-space: nowrap; font-family: monospace; font-size: 0.8rem;">${escapeHtml(ts)}</td>
                <td><span class="badge badge-info">${escapeHtml(ev.action || '')}</span></td>
                <td><strong>${escapeHtml(ev.manufacturer || '')}</strong></td>
                <td>${escapeHtml(ev.equipment || '')}</td>
                <td>${defs || '-'}</td>
                <td><span class="badge ${statusClass}">${escapeHtml(ev.validation_status || 'DONE')}</span></td>
                <td style="max-width: 320px; font-size: 0.8rem; color: var(--text-secondary);">${escapeHtml(ev.details || '')}</td>
            `;
            auditTrailBody.appendChild(tr);
        });
    }

    if (refreshAuditTrailBtn && refreshAuditTrailBtn.addEventListener) {
        refreshAuditTrailBtn.addEventListener('click', () => loadAuditTrail());
    }

    if (auditTrailSearchInput && auditTrailSearchInput.addEventListener) {
        auditTrailSearchInput.addEventListener('input', () => {
            const q = auditTrailSearchInput.value.trim().toLowerCase();
            if (!q) {
                renderAuditTrailTable(auditTrailItems);
                return;
            }
            const filtered = auditTrailItems.filter(ev => {
                return (ev.action && ev.action.toLowerCase().includes(q)) ||
                       (ev.manufacturer && ev.manufacturer.toLowerCase().includes(q)) ||
                       (ev.equipment && ev.equipment.toLowerCase().includes(q)) ||
                       (ev.details && ev.details.toLowerCase().includes(q));
            });
            renderAuditTrailTable(filtered);
        });
    }

    // =========================================================================
    // MODULE: Versions & Commits (/api/versions)
    // =========================================================================
    const versionsTableBody = document.getElementById('versionsTableBody');
    const versionsCountBadge = document.getElementById('versionsCountBadge');
    const refreshVersionsBtn = document.getElementById('refreshVersionsBtn');
    const versionsSearchInput = document.getElementById('versionsSearchInput');

    const versionDetailSection = document.getElementById('versionDetailSection');
    const detailCommitHash = document.getElementById('detailCommitHash');
    const detailMetaSummary = document.getElementById('detailMetaSummary');
    const detailRestoreBtn = document.getElementById('detailRestoreBtn');
    const closeDetailBtn = document.getElementById('closeDetailBtn');

    const viewDiffTabBtn = document.getElementById('viewDiffTabBtn');
    const viewManifestTabBtn = document.getElementById('viewManifestTabBtn');
    const diffViewPane = document.getElementById('diffViewPane');
    const manifestViewPane = document.getElementById('manifestViewPane');

    const diffParentBadge = document.getElementById('diffParentBadge');
    const diffAddedBadge = document.getElementById('diffAddedBadge');
    const diffModifiedBadge = document.getElementById('diffModifiedBadge');
    const diffRemovedBadge = document.getElementById('diffRemovedBadge');
    const diffFileSelect = document.getElementById('diffFileSelect');
    const diffContainer = document.getElementById('diffContainer');

    const manifestFilesCount = document.getElementById('manifestFilesCount');
    const manifestFilterInput = document.getElementById('manifestFilterInput');
    const manifestTableBody = document.getElementById('manifestTableBody');

    // Restore modal elements
    const restoreModal = document.getElementById('restoreModal');
    const closeRestoreModalBtn = document.getElementById('closeRestoreModalBtn');
    const cancelRestoreBtn = document.getElementById('cancelRestoreBtn');
    const confirmRestoreBtn = document.getElementById('confirmRestoreBtn');
    const modalCommitId = document.getElementById('modalCommitId');
    const modalAuthor = document.getElementById('modalAuthor');
    const modalTimestamp = document.getElementById('modalTimestamp');
    const modalMessage = document.getElementById('modalMessage');
    const restoreTargetDir = document.getElementById('restoreTargetDir');
    const restoreSpecificFile = document.getElementById('restoreSpecificFile');
    const restorePreviewCountBadge = document.getElementById('restorePreviewCountBadge');
    const restorePreviewList = document.getElementById('restorePreviewList');

    let versionsList = [];
    let currentCommitData = null;
    let currentDiffData = null;
    let pendingRestoreCommit = null;

    async function loadVersions() {
        if (!versionsTableBody) return;
        try {
            versionsTableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: var(--text-secondary);">Loading commit history...</td></tr>`;
            const res = await fetch('/api/versions?limit=50');
            if (!res.ok) {
                throw new Error(`HTTP ${res.status}: ${res.statusText}`);
            }
            versionsList = await res.json();
            versionsLoaded = true;
            if (versionsCountBadge) {
                versionsCountBadge.textContent = `${versionsList.length} snapshot${versionsList.length > 1 ? 's' : ''}`;
            }
            renderVersionsTable(versionsList);
        } catch (err) {
            versionsTableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: #ef4444;">Error loading commits: ${escapeHtml(err.message)}</td></tr>`;
            showToast(`Commit history error: ${err.message}`, 'error');
        }
    }

    function renderVersionsTable(commits) {
        if (!versionsTableBody) return;
        versionsTableBody.innerHTML = '';
        if (!commits || commits.length === 0) {
            versionsTableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: var(--text-secondary);">No commits recorded in definitions_history.</td></tr>`;
            return;
        }

        commits.forEach(commit => {
            const tr = document.createElement('tr');
            const shortId = (commit.commit_id || '').substring(0, 10);
            const ts = (commit.timestamp || '').replace('T', ' ').substring(0, 19);
            const tagsHtml = (commit.tags || []).map(t => `<span class="tag-pill">🏷️ ${escapeHtml(t)}</span>`).join('') || '<span style="color: var(--text-secondary);">-</span>';
            const fileCount = commit.files_count ?? 0;
            const regCountHtml = commit.total_registers != null
                ? `<span class="badge" style="background: rgba(255,255,255,0.08);">${commit.total_registers} regs</span>`
                : (commit.changes ? `<span class="badge" style="background: rgba(255,255,255,0.08);">+${commit.changes.added || 0} ~${commit.changes.modified || 0} -${commit.changes.deleted || 0}</span>` : '');

            tr.innerHTML = `
                <td><code style="color: var(--accent); cursor: pointer;" title="Inspect this commit">${escapeHtml(shortId)}</code></td>
                <td style="white-space: nowrap; font-size: 0.8rem; font-family: monospace;">${escapeHtml(ts)}</td>
                <td><strong>${escapeHtml(commit.author || 'Unknown')}</strong></td>
                <td>${escapeHtml(commit.message || 'No message')}</td>
                <td><span class="badge badge-info">${fileCount} files</span> ${regCountHtml}</td>
                <td>${tagsHtml}</td>
                <td style="text-align: right; white-space: nowrap;">
                    <button class="btn btn-secondary btn-sm btn-inspect" data-id="${escapeHtml(commit.commit_id)}">🔍 Diff & Details</button>
                    <button class="btn btn-primary btn-sm btn-restore" data-id="${escapeHtml(commit.commit_id)}">🔄 Restore</button>
                </td>
            `;

            const codeEl = tr.querySelector('code');
            if (codeEl) codeEl.addEventListener('click', () => inspectCommit(commit.commit_id));
            const inspectBtn = tr.querySelector('.btn-inspect');
            if (inspectBtn) inspectBtn.addEventListener('click', () => inspectCommit(commit.commit_id));
            const restoreBtn = tr.querySelector('.btn-restore');
            if (restoreBtn) restoreBtn.addEventListener('click', () => openRestoreModal(commit));
            versionsTableBody.appendChild(tr);
        });
    }

    async function inspectCommit(commitId) {
        if (!versionDetailSection) return;
        try {
            showToast(`Loading commit ${commitId.substring(0, 8)}...`, 'info');
            versionDetailSection.classList.remove('hidden');
            if (detailCommitHash) detailCommitHash.textContent = commitId.substring(0, 12);
            if (detailMetaSummary) detailMetaSummary.innerHTML = 'Loading metadata and diff...';

            const [commitRes, diffRes] = await Promise.all([
                fetch(`/api/versions/${encodeURIComponent(commitId)}`),
                fetch(`/api/versions/${encodeURIComponent(commitId)}/diff`)
            ]);

            if (!commitRes.ok) throw new Error(`Failed to load commit: ${commitRes.status}`);
            currentCommitData = await commitRes.json();
            currentDiffData = diffRes.ok ? await diffRes.json() : null;

            const ts = (currentCommitData.timestamp || '').replace('T', ' ').substring(0, 19);
            const parentShort = currentCommitData.parent_id ? currentCommitData.parent_id.substring(0, 10) : 'none (initial)';
            if (detailMetaSummary) {
                detailMetaSummary.innerHTML = `
                    <strong>Author:</strong> ${escapeHtml(currentCommitData.author)} | 
                    <strong>Date:</strong> ${escapeHtml(ts)} UTC | 
                    <strong>Parent:</strong> <code>${escapeHtml(parentShort)}</code> | 
                    <strong>Message:</strong> <em>${escapeHtml(currentCommitData.message)}</em>
                `;
            }

            // Populate diff files dropdown
            if (diffFileSelect) {
                diffFileSelect.innerHTML = `<option value="">All modified files</option>`;
                if (currentDiffData) {
                    const allDiffFiles = [
                        ...(currentDiffData.added_files || []),
                        ...(currentDiffData.modified_files || []),
                        ...(currentDiffData.removed_files || [])
                    ].sort();
                    allDiffFiles.forEach(f => {
                        const opt = document.createElement('option');
                        opt.value = f;
                        opt.textContent = f;
                        diffFileSelect.appendChild(opt);
                    });
                }
            }

            renderCommitDiff(currentDiffData, diffFileSelect ? diffFileSelect.value : '');
            renderCommitManifest(currentCommitData.manifest || {});

            versionDetailSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } catch (err) {
            showToast(`Inspection error: ${err.message}`, 'error');
            if (detailMetaSummary) detailMetaSummary.innerHTML = `<span style="color: #ef4444;">Error: ${escapeHtml(err.message)}</span>`;
        }
    }

    function renderCommitDiff(diff, filterFile = '') {
        if (!diffContainer) return;
        diffContainer.innerHTML = '';

        if (!diff) {
            diffContainer.innerHTML = `<div style="text-align: center; color: var(--text-secondary); padding: 20px;">No diff data available.</div>`;
            return;
        }

        const addedFiles = diff.added_files || [];
        const modifiedFiles = diff.modified_files || [];
        const removedFiles = diff.removed_files || [];

        if (diffAddedBadge) diffAddedBadge.textContent = `+${addedFiles.length} added`;
        if (diffModifiedBadge) diffModifiedBadge.textContent = `~${modifiedFiles.length} modified`;
        if (diffRemovedBadge) diffRemovedBadge.textContent = `-${removedFiles.length} removed`;
        if (diffParentBadge) {
            diffParentBadge.textContent = diff.commit_b ? `Diff vs ${diff.commit_b.substring(0, 8)}` : `Initial Snapshot`;
        }

        if (addedFiles.length === 0 && modifiedFiles.length === 0 && removedFiles.length === 0) {
            diffContainer.innerHTML = `<div style="text-align: center; color: #10b981; padding: 20px; background: rgba(16,185,129,0.08); border-radius: 8px;">No file or register differences detected relative to parent.</div>`;
            return;
        }

        // Render added files
        addedFiles.filter(f => !filterFile || f === filterFile).forEach(fn => {
            const card = document.createElement('div');
            card.className = 'diff-file-card';
            card.innerHTML = `
                <div class="diff-file-header diff-row-added">
                    <span style="display: flex; align-items: center; gap: 8px;">
                        <span class="diff-badge-add">+ ADDED</span>
                        <code>${escapeHtml(fn)}</code>
                    </span>
                    <span style="font-size: 0.8rem; color: #34d399;">New certified definition file</span>
                </div>
            `;
            diffContainer.appendChild(card);
        });

        // Render removed files
        removedFiles.filter(f => !filterFile || f === filterFile).forEach(fn => {
            const card = document.createElement('div');
            card.className = 'diff-file-card';
            card.innerHTML = `
                <div class="diff-file-header diff-row-removed">
                    <span style="display: flex; align-items: center; gap: 8px;">
                        <span class="diff-badge-del">- REMOVED</span>
                        <code>${escapeHtml(fn)}</code>
                    </span>
                    <span style="font-size: 0.8rem; color: #f87171;">File removed or archived</span>
                </div>
            `;
            diffContainer.appendChild(card);
        });

        // Render modified files with detailed register diffs
        const fileDiffs = diff.file_diffs || {};
        modifiedFiles.filter(f => !filterFile || f === filterFile).forEach(fn => {
            const fd = fileDiffs[fn];
            const card = document.createElement('div');
            card.className = 'diff-file-card';

            const addCount = fd ? (fd.added_registers_count || 0) : 0;
            const delCount = fd ? (fd.removed_registers_count || 0) : 0;
            const modCount = fd ? (fd.modified_registers_count || 0) : 0;

            let bodyHtml = '';
            if (fd) {
                let parts = [];
                if (fd.header_changed) {
                    parts.push(`<div style="padding: 8px 14px; background: rgba(245,158,11,0.1); color: #fbbf24; font-size: 0.8rem; border-bottom: 1px solid var(--border);">⚠️ CSV header changed: <code>${escapeHtml(fd.old_header)}</code> ➔ <code>${escapeHtml(fd.new_header)}</code></div>`);
                }

                if (addCount > 0 && fd.added_addresses && fd.added_addresses.length) {
                    const addPills = fd.added_addresses.map(a => `<span class="diff-badge-add">+${escapeHtml(a)}</span>`).join(' ');
                    parts.push(`<div style="padding: 8px 14px; font-size: 0.8rem;"><strong>Registers added:</strong> ${addPills}</div>`);
                }

                if (delCount > 0 && fd.removed_addresses && fd.removed_addresses.length) {
                    const delPills = fd.removed_addresses.map(a => `<span class="diff-badge-del">-${escapeHtml(a)}</span>`).join(' ');
                    parts.push(`<div style="padding: 8px 14px; font-size: 0.8rem;"><strong>Registers removed:</strong> ${delPills}</div>`);
                }

                if (modCount > 0 && fd.modified_registers && fd.modified_registers.length) {
                    let regRows = '';
                    fd.modified_registers.forEach(m => {
                        const changes = m.changes || {};
                        const pills = Object.entries(changes).map(([k, [oldVal, newVal]]) => {
                            return `<span class="diff-field-pill">${escapeHtml(k)}: <span class="val-old">${escapeHtml(oldVal || 'empty')}</span> ➔ <span class="val-new">${escapeHtml(newVal || 'empty')}</span></span>`;
                        }).join(' ');

                        regRows += `
                            <tr class="diff-row-modified">
                                <td style="width: 80px; font-family: monospace; font-weight: 600;">${escapeHtml(m.address)}</td>
                                <td style="font-weight: 500;">${escapeHtml(m.name || '')}</td>
                                <td>${pills}</td>
                            </tr>
                        `;
                    });

                    parts.push(`
                        <div class="table-wrapper" style="max-height: 250px; border-left: none; border-right: none; border-bottom: none;">
                            <table>
                                <thead>
                                    <tr>
                                        <th>Address</th>
                                        <th>Name</th>
                                        <th>Detected Modifications</th>
                                    </tr>
                                </thead>
                                <tbody>${regRows}</tbody>
                            </table>
                        </div>
                    `);
                }

                bodyHtml = parts.join('');
            }

            card.innerHTML = `
                <div class="diff-file-header diff-row-modified">
                    <span style="display: flex; align-items: center; gap: 8px;">
                        <span class="diff-badge-mod">~ MODIFIED</span>
                        <code>${escapeHtml(fn)}</code>
                    </span>
                    <span style="font-size: 0.8rem; display: flex; gap: 6px;">
                        <span class="badge badge-success">+${addCount}</span>
                        <span class="badge badge-warning">~${modCount}</span>
                        <span class="badge badge-error">-${delCount}</span>
                    </span>
                </div>
                ${bodyHtml}
            `;
            diffContainer.appendChild(card);
        });
    }

    function renderCommitManifest(manifest) {
        if (!manifestTableBody) return;
        manifestTableBody.innerHTML = '';
        const files = Object.values(manifest || {});
        if (manifestFilesCount) manifestFilesCount.textContent = files.length;

        if (files.length === 0) {
            manifestTableBody.innerHTML = `<tr><td colspan="8" style="text-align: center; color: var(--text-secondary);">No files in this manifest.</td></tr>`;
            return;
        }

        files.forEach(f => {
            const tr = document.createElement('tr');
            const shortSha = (f.sha256 || '').substring(0, 8);
            const sizeKb = ((f.size_bytes || 0) / 1024).toFixed(1);

            tr.innerHTML = `
                <td><code>${escapeHtml(f.file_name)}</code></td>
                <td><span class="badge badge-info">${escapeHtml(f.category || 'OTH')}</span></td>
                <td><strong>${escapeHtml(f.manufacturer || '-')}</strong></td>
                <td>${escapeHtml(f.model || '-')}</td>
                <td>${escapeHtml(f.register_count ?? 0)}</td>
                <td>${escapeHtml(f.tag_count ?? 0)}</td>
                <td style="font-size: 0.8rem; color: var(--text-secondary);">${sizeKb} KB</td>
                <td><code style="font-size: 0.75rem;">${escapeHtml(shortSha)}</code></td>
            `;
            manifestTableBody.appendChild(tr);
        });
    }

    if (viewDiffTabBtn && viewManifestTabBtn && diffViewPane && manifestViewPane) {
        viewDiffTabBtn.addEventListener('click', () => {
            viewDiffTabBtn.classList.add('active');
            viewManifestTabBtn.classList.remove('active');
            diffViewPane.classList.remove('hidden');
            manifestViewPane.classList.add('hidden');
        });
        viewManifestTabBtn.addEventListener('click', () => {
            viewManifestTabBtn.classList.add('active');
            viewDiffTabBtn.classList.remove('active');
            manifestViewPane.classList.remove('hidden');
            diffViewPane.classList.add('hidden');
        });
    }

    if (diffFileSelect) {
        diffFileSelect.addEventListener('change', () => {
            if (currentDiffData) {
                renderCommitDiff(currentDiffData, diffFileSelect.value);
            }
        });
    }

    if (closeDetailBtn && versionDetailSection) {
        closeDetailBtn.addEventListener('click', () => {
            versionDetailSection.classList.add('hidden');
        });
    }

    if (refreshVersionsBtn) {
        refreshVersionsBtn.addEventListener('click', () => loadVersions());
    }

    if (versionsSearchInput) {
        versionsSearchInput.addEventListener('input', () => {
            const q = versionsSearchInput.value.trim().toLowerCase();
            if (!q) {
                renderVersionsTable(versionsList);
                return;
            }
            const filtered = versionsList.filter(c => {
                const tagsStr = (c.tags || []).join(' ').toLowerCase();
                return (c.commit_id && c.commit_id.toLowerCase().includes(q)) ||
                       (c.message && c.message.toLowerCase().includes(q)) ||
                       (c.author && c.author.toLowerCase().includes(q)) ||
                       tagsStr.includes(q);
            });
            renderVersionsTable(filtered);
        });
    }

    if (manifestFilterInput && manifestTableBody) {
        setupTableFilter(manifestFilterInput, manifestTableBody);
    }

    // Modal restore helpers
    function openRestoreModal(commit) {
        if (!restoreModal) return;
        pendingRestoreCommit = commit;
        if (modalCommitId) modalCommitId.textContent = commit.commit_id;
        if (modalAuthor) modalAuthor.textContent = commit.author || 'Unknown';
        if (modalTimestamp) modalTimestamp.textContent = (commit.timestamp || '').replace('T', ' ').substring(0, 19);
        if (modalMessage) modalMessage.textContent = commit.message || 'No message';

        if (restorePreviewList) {
            const files = Object.keys(commit.manifest || {});
            const total = files.length || commit.files_count || 0;
            if (restorePreviewCountBadge) restorePreviewCountBadge.textContent = `${total} files`;
            if (files.length > 0) {
                const sample = files.slice(0, 15).map(f => `• ${escapeHtml(f)}`).join('<br>');
                const extra = files.length > 15 ? `<br><em>... and ${files.length - 15} other files</em>` : '';
                restorePreviewList.innerHTML = `${sample}${extra}`;
            } else {
                restorePreviewList.innerHTML = `<em>This snapshot will restore ${total} certified definition files.</em>`;
            }
        }

        restoreModal.classList.remove('hidden');
    }

    function closeRestoreModal() {
        if (restoreModal) restoreModal.classList.add('hidden');
        pendingRestoreCommit = null;
    }

    if (closeRestoreModalBtn) closeRestoreModalBtn.addEventListener('click', closeRestoreModal);
    if (cancelRestoreBtn) cancelRestoreBtn.addEventListener('click', closeRestoreModal);
    if (detailRestoreBtn) {
        detailRestoreBtn.addEventListener('click', () => {
            if (currentCommitData) openRestoreModal(currentCommitData);
        });
    }

    if (confirmRestoreBtn) {
        confirmRestoreBtn.addEventListener('click', async () => {
            if (!pendingRestoreCommit) return;
            const targetDir = (restoreTargetDir ? restoreTargetDir.value.trim() : '') || 'production';
            const fileName = restoreSpecificFile && restoreSpecificFile.value.trim() ? restoreSpecificFile.value.trim() : null;

            confirmRestoreBtn.disabled = true;
            confirmRestoreBtn.textContent = 'Restoring in progress...';

            try {
                const res = await fetch('/api/versions/restore', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        commit_id: pendingRestoreCommit.commit_id,
                        target_dir: targetDir,
                        file_name: fileName
                    })
                });

                if (!res.ok) {
                    const errData = await res.json().catch(() => ({}));
                    throw new Error(errData.detail || `HTTP Error ${res.status}`);
                }

                const data = await res.json();
                showToast(`Restore successful! ${data.restored_count} file(s) restored into ${data.target_dir}.`, 'success');
                closeRestoreModal();

                // Refresh library if open
                if (libraryLoaded && typeof loadLibrary === 'function') {
                    loadLibrary(currentLibrarySource);
                }
            } catch (err) {
                showToast(`Restore failed: ${err.message}`, 'error');
            } finally {
                confirmRestoreBtn.disabled = false;
                confirmRestoreBtn.textContent = 'Confirm Restore';
            }
        });
    }

    // =========================================================================
    // MODULE: Doublons & Qualité (/api/duplicates)
    // =========================================================================
    const dupPathInput = document.getElementById('dupPathInput');
    const scanDuplicatesBtn = document.getElementById('scanDuplicatesBtn');
    const dupGlobalStatusBadge = document.getElementById('dupGlobalStatusBadge');
    const dupSearchInput = document.getElementById('dupSearchInput');

    const kpiScannedFiles = document.getElementById('kpiScannedFiles');
    const kpiTotalDuplicates = document.getElementById('kpiTotalDuplicates');
    const kpiCategoryConflicts = document.getElementById('kpiCategoryConflicts');
    const kpiAliasSynonyms = document.getElementById('kpiAliasSynonyms');
    const kpiHardwareFamilies = document.getElementById('kpiHardwareFamilies');

    const badgeCategoryConflictCount = document.getElementById('badgeCategoryConflictCount');
    const badgeAliasCount = document.getElementById('badgeAliasCount');
    const badgeHardwareFamilyCount = document.getElementById('badgeHardwareFamilyCount');

    const categoryConflictBody = document.getElementById('categoryConflictBody');
    const aliasBody = document.getElementById('aliasBody');
    const hardwareFamilyBody = document.getElementById('hardwareFamilyBody');

    let duplicatesReport = null;

    async function loadDuplicates(path = 'production') {
        if (!categoryConflictBody) return;
        try {
            if (scanDuplicatesBtn) {
                scanDuplicatesBtn.disabled = true;
                scanDuplicatesBtn.textContent = '⏳ Scanning...';
            }
            if (dupGlobalStatusBadge) {
                dupGlobalStatusBadge.className = 'badge badge-info';
                dupGlobalStatusBadge.textContent = 'Scanning in progress...';
            }

            const target = path || (dupPathInput ? dupPathInput.value.trim() : 'production') || 'production';
            const res = await fetch(`/api/duplicates?path=${encodeURIComponent(target)}`);
            if (!res.ok) {
                const errData = await res.json().catch(() => ({}));
                throw new Error(errData.detail || `HTTP ${res.status}`);
            }

            duplicatesReport = await res.json();
            duplicatesLoaded = true;
            renderDuplicates(duplicatesReport);
            showToast('Duplicates scan completed successfully', 'success');
        } catch (err) {
            showToast(`Error scanning duplicates: ${err.message}`, 'error');
            if (dupGlobalStatusBadge) {
                dupGlobalStatusBadge.className = 'badge badge-error';
                dupGlobalStatusBadge.textContent = 'Scan failed';
            }
        } finally {
            if (scanDuplicatesBtn) {
                scanDuplicatesBtn.disabled = false;
                scanDuplicatesBtn.textContent = '🔄 Scan Duplicates';
            }
        }
    }

    function renderDuplicates(report) {
        if (!report) return;

        // KPIs
        const scanned = report.total_files_scanned ?? 0;
        const totalDupes = report.total_duplicate_files ?? 0;
        const catClusters = report.category_conflict_clusters || [];
        const aliasClusters = report.alias_clusters || [];
        const bodyClusters = report.exact_body_clusters || [];

        if (kpiScannedFiles) kpiScannedFiles.textContent = scanned;
        if (kpiTotalDuplicates) kpiTotalDuplicates.textContent = totalDupes;
        if (kpiCategoryConflicts) kpiCategoryConflicts.textContent = catClusters.length;
        if (kpiAliasSynonyms) kpiAliasSynonyms.textContent = aliasClusters.length;
        if (kpiHardwareFamilies) kpiHardwareFamilies.textContent = bodyClusters.length;

        if (badgeCategoryConflictCount) {
            badgeCategoryConflictCount.textContent = `${catClusters.length} conflict${catClusters.length > 1 ? 's' : ''}`;
            badgeCategoryConflictCount.className = `badge ${catClusters.length > 0 ? 'badge-error' : 'badge-success'}`;
        }
        if (badgeAliasCount) {
            badgeAliasCount.textContent = `${aliasClusters.length} alias${aliasClusters.length > 1 ? 'es' : ''}`;
            badgeAliasCount.className = `badge ${aliasClusters.length > 0 ? 'badge-warning' : 'badge-success'}`;
        }
        if (badgeHardwareFamilyCount) {
            badgeHardwareFamilyCount.textContent = `${bodyClusters.length} famil${bodyClusters.length > 1 ? 'ies' : 'y'}`;
            badgeHardwareFamilyCount.className = 'badge badge-success';
        }

        if (dupGlobalStatusBadge) {
            if (catClusters.length === 0 && aliasClusters.length === 0) {
                dupGlobalStatusBadge.className = 'badge badge-success';
                dupGlobalStatusBadge.textContent = '✅ Compliant (0 OTH_ conflicts)';
            } else {
                dupGlobalStatusBadge.className = 'badge badge-warning';
                dupGlobalStatusBadge.textContent = `⚠️ ${totalDupes} Duplicates Detected`;
            }
        }

        renderCategoryConflicts(catClusters);
        renderAliasClusters(aliasClusters);
        renderHardwareFamilies(bodyClusters);
    }

    function renderCategoryConflicts(clusters) {
        if (!categoryConflictBody) return;
        categoryConflictBody.innerHTML = '';
        if (clusters.length === 0) {
            categoryConflictBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #10b981; padding: 16px;">✅ No OTH_ category conflicts detected. All definition files are typed according to the 5 Pillars.</td></tr>`;
            return;
        }

        clusters.forEach(c => {
            const tr = document.createElement('tr');
            const dupes = (c.duplicate_files || []).map(d => `<code style="color: #f87171;">${escapeHtml(d)}</code>`).join('<br>');
            const model = (c.models && c.models.length) ? c.models[0] : '-';

            tr.innerHTML = `
                <td><code style="color: #34d399; font-weight: 600;">${escapeHtml(c.canonical_file)}</code></td>
                <td>${dupes}</td>
                <td><strong>${escapeHtml(model)}</strong></td>
                <td><span class="badge badge-info">${escapeHtml(c.register_count ?? 0)} regs</span></td>
                <td><span class="badge badge-error">Delete redundant OTH_</span></td>
            `;
            categoryConflictBody.appendChild(tr);
        });
    }

    function renderAliasClusters(clusters) {
        if (!aliasBody) return;
        aliasBody.innerHTML = '';
        if (clusters.length === 0) {
            aliasBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #10b981; padding: 16px;">✅ No unresolved aliases detected in the catalog.</td></tr>`;
            return;
        }

        clusters.forEach(c => {
            const tr = document.createElement('tr');
            const dupes = (c.duplicate_files || []).map(d => `<code style="color: #fbbf24;">${escapeHtml(d)}</code>`).join('<br>');

            tr.innerHTML = `
                <td><code style="color: #38bdf8; font-weight: 600;">${escapeHtml(c.canonical_file)}</code></td>
                <td>${dupes}</td>
                <td style="font-size: 0.85rem; color: var(--text-secondary);">${escapeHtml(c.description || 'Naming variant')}</td>
                <td><span class="badge badge-info">${escapeHtml(c.register_count ?? 0)} regs</span></td>
                <td><span class="badge badge-warning">Register in alias catalog</span></td>
            `;
            aliasBody.appendChild(tr);
        });
    }

    function renderHardwareFamilies(clusters) {
        if (!hardwareFamilyBody) return;
        hardwareFamilyBody.innerHTML = '';
        if (clusters.length === 0) {
            hardwareFamilyBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-secondary); padding: 16px;">No multi-model families recorded.</td></tr>`;
            return;
        }

        clusters.forEach(c => {
            const tr = document.createElement('tr');
            const models = (c.models || []).map(m => `<span class="model-pill">${escapeHtml(m)}</span>`).join(' ') || '-';
            const dupes = (c.duplicate_files || []).map(d => `<code>${escapeHtml(d)}</code>`).join('<br>') || '-';

            tr.innerHTML = `
                <td><code style="color: #10b981; font-weight: 600;">${escapeHtml(c.canonical_file)}</code></td>
                <td style="max-width: 320px;">${models}</td>
                <td>${dupes}</td>
                <td><span class="badge badge-info">${escapeHtml(c.register_count ?? 0)} regs</span></td>
                <td><span class="badge badge-success">Valid Hardware Family</span></td>
            `;
            hardwareFamilyBody.appendChild(tr);
        });
    }

    if (scanDuplicatesBtn) {
        scanDuplicatesBtn.addEventListener('click', () => {
            const p = dupPathInput ? dupPathInput.value.trim() : 'production';
            loadDuplicates(p);
        });
    }

    if (dupSearchInput) {
        dupSearchInput.addEventListener('input', () => {
            const q = dupSearchInput.value.trim().toLowerCase();
            const filterTable = (tbody) => {
                if (!tbody) return;
                const rows = tbody.children || [];
                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row) continue;
                    const text = (row.textContent || '').toLowerCase();
                    row.style.display = (!q || text.includes(q)) ? '' : 'none';
                }
            };
            filterTable(categoryConflictBody);
            filterTable(aliasBody);
            filterTable(hardwareFamilyBody);
        });
    }
});
