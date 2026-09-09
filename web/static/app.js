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
    const downloadBtn = document.getElementById('downloadBtn');

    let selectedFile = null;
    let convertedCsvContent = null;
    let convertedFilename = 'definition.csv';

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
        fileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        dropZone.classList.add('hidden');
        fileDetails.classList.remove('hidden');
        convertBtn.disabled = false;
    }

    removeFileBtn.addEventListener('click', () => {
        selectedFile = null;
        fileInput.value = '';
        dropZone.classList.remove('hidden');
        fileDetails.classList.add('hidden');
        convertBtn.disabled = true;
        resultsSection.classList.add('hidden');
    });

    convertBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

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
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                alert(`Conversion Error: ${errData.detail || 'Unknown error'}`);
                return;
            }

            const data = await response.json();
            convertedCsvContent = data.csv_content;
            convertedFilename = data.filename;

            registerBadge.textContent = `${data.register_count} registers`;
            renderPreview(data.preview);
            resultsSection.classList.remove('hidden');

        } catch (err) {
            alert(`Network or Server Error: ${err.message}`);
        } finally {
            convertBtn.disabled = false;
            convertBtn.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg> Extract & Generate Definition`;
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
                <td><code>${escapeHtml(r.Address || '')}</code></td>
                <td>${escapeHtml(r.Type || '')}</td>
                <td>${escapeHtml(r.Factor || '1')}</td>
                <td>${escapeHtml(r.Offset || '0')}</td>
                <td>${escapeHtml(r.Unit || '')}</td>
                <td>${escapeHtml(r.Action || '1')}</td>
            `;
            previewBody.appendChild(tr);
        });
    }

    function escapeHtml(str) {
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    downloadBtn.addEventListener('click', () => {
        if (!convertedCsvContent) return;
        const blob = new Blob([convertedCsvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', convertedFilename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
});
