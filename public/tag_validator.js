document.addEventListener('DOMContentLoaded', () => {
    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');
    const uploadText = document.getElementById('uploadText');
    const runBtn = document.getElementById('runBtn');
    const logBox = document.getElementById('logBox');
    const statusContainer = document.getElementById('statusContainer');
    const resultsTableContainer = document.getElementById('resultsTableContainer');
    const resultsBody = document.getElementById('resultsBody');
    const emptyState = document.getElementById('emptyState');
    const downloadBtn = document.getElementById('downloadBtn');

    let isRunning = false;
    let pollInterval = null;

    // --- File Upload ---
    uploadZone.onclick = () => fileInput.click();
    
    uploadZone.ondragover = (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    };

    uploadZone.ondragleave = () => uploadZone.classList.remove('dragover');

    uploadZone.ondrop = (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    };

    fileInput.onchange = (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    };

    async function handleFile(file) {
        uploadText.innerHTML = `📄 <b>${file.name}</b><br><small>Uploading...</small>`;
        
        const formData = new FormData();
        formData.append('file', file);

        try {
            const resp = await fetch('/api/tag-validator/upload', {
                method: 'POST',
                body: formData
            });
            const data = await resp.json();
            if (data.success) {
                uploadText.innerHTML = `✅ <b>${file.name}</b> ready!`;
                runBtn.disabled = false;
            } else {
                alert('Upload failed: ' + data.error);
            }
        } catch (err) {
            console.error(err);
            alert('Upload error');
        }
    }

    // --- Execution ---
    runBtn.onclick = async () => {
        if (isRunning) return;
        
        isRunning = true;
        runBtn.disabled = true;
        runBtn.innerHTML = 'Running Validation <span class="loading-dots"></span>';
        statusContainer.classList.remove('hidden');
        
        try {
            const resp = await fetch('/api/tag-validator/run', { method: 'POST' });
            const data = await resp.json();
            if (data.success) {
                startPolling();
            } else {
                alert('Run failed: ' + data.error);
                isRunning = false;
                runBtn.disabled = false;
                runBtn.innerText = 'Run Validation';
            }
        } catch (err) {
            console.error(err);
            isRunning = false;
        }
    };

    function startPolling() {
        if (pollInterval) clearInterval(pollInterval);
        pollInterval = setInterval(async () => {
            try {
                const resp = await fetch('/api/tag-validator/status');
                const data = await resp.json();
                
                updateLogs(data.logs);
                
                if (!data.running && isRunning) {
                    clearInterval(pollInterval);
                    finishRun();
                }
            } catch (err) {
                console.error('Polling error', err);
            }
        }, 1000);
    }

    function updateLogs(logs) {
        logBox.innerHTML = logs.map(line => {
            const cls = line.includes('ERROR') ? 'log-line error' : 'log-line';
            return `<div class="${cls}">${line}</div>`;
        }).join('');
        logBox.scrollTop = logBox.scrollHeight;
    }

    async function finishRun() {
        isRunning = false;
        runBtn.disabled = false;
        runBtn.innerText = 'Run Again';
        
        // Fetch results
        const resp = await fetch('/api/tag-validator/results');
        const data = await resp.json();
        
        if (data.results && data.results.length > 0) {
            renderResults(data.results);
            downloadBtn.classList.remove('hidden');
        }
    }

    function renderResults(results) {
        emptyState.classList.add('hidden');
        resultsTableContainer.classList.remove('hidden');
        
        resultsBody.innerHTML = results.map(row => `
            <tr>
                <td style="font-size: 0.85rem; color: #cbd5e1; max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
                    ${row.URL}
                </td>
                <td>
                    <span class="badge ${row.Tealium_Loaded === 'PASS' ? 'badge-pass' : 'badge-fail'}">
                        ${row.Tealium_Loaded}
                    </span>
                </td>
                <td>
                    <span class="badge ${row.Tealium_View_Fired === 'PASS' ? 'badge-pass' : 'badge-fail'}">
                        ${row.Tealium_View_Fired}
                    </span>
                </td>
                <td>
                    <span class="badge ${row.GA4_Fired === 'PASS' ? 'badge-pass' : 'badge-fail'}">
                        ${row.GA4_Fired}
                    </span>
                </td>
            </tr>
        `).join('');
    }

    downloadBtn.onclick = () => {
        window.location.href = '/api/tag-validator/download';
    };
});
