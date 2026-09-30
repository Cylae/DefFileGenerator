document.addEventListener('DOMContentLoaded', () => {
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

    // Drag & Drop
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        });
    });

    dropZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length) handleFile(files[0]);
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) handleFile(e.target.files[0]);
    });

    function handleFile(file) {
        selectedFile = file;
        invalidateConversion();
        fileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        dropZone.classList.add('hidden');
        fileDetails.classList.remove('hidden');
        convertBtn.disabled = false;
    }

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
            renderPreview(data.preview);
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

    function renderPreview(rows) {
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

    function escapeHtml(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    downloadBtn.addEventListener('click', () => {
        if (!convertedCsvContent || downloadBtn.disabled) return;
        const blob = new Blob(['\ufeff', convertedCsvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', convertedFilename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    });
});
