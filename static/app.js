/* Peptone Analysis Data Arrangement Tool — Frontend Logic */

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// State
let uploadedFiles = [];       // { file, type, name }
let sampleRows = [];          // { file_name, display_name, sheet_name, raw_name }
let resultFileId = null;
let resultFilename = null;

// DOM refs
const dropZone = $('#drop-zone');
const fileInput = $('#file-input');
const fileList = $('#file-list');
const sampleCard = $('#sample-card');
const sensangCard = $('#sensang-card');
const sampleTbody = $('#sample-tbody');
const sensangTbody = $('#sensang-tbody');
const btnProcess = $('#btn-process');
const btnDownload = $('#btn-download');
const btnAddSample = $('#btn-add-sample');
const statusArea = $('#status-area');
const statusLoading = $('#status-loading');
const statusSuccess = $('#status-success');
const statusError = $('#status-error');
const statusMessage = $('#status-message');
const errorMessage = $('#error-message');
const summaryPlaceholder = $('#summary-placeholder');
const summaryContent = $('#summary-content');
const summaryGrid = $('#summary-grid');

// ── File Upload ──
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    handleFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', () => {
    handleFiles(fileInput.files);
    fileInput.value = '';
});

function classifyFile(name) {
    return name.includes('의뢰품검사상세') ? 'lab' : 'summary';
}

function handleFiles(fileListInput) {
    for (const f of fileListInput) {
        const ext = f.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls'].includes(ext)) continue;
        if (uploadedFiles.some(u => u.name === f.name)) continue;
        uploadedFiles.push({ file: f, type: classifyFile(f.name), name: f.name });
    }
    renderFileList();
    prescanFiles();
}

function renderFileList() {
    fileList.innerHTML = '';
    uploadedFiles.forEach((uf, idx) => {
        const div = document.createElement('div');
        div.className = 'file-item';
        const typeLabel = uf.type === 'lab' ? '의뢰품검사상세' : '엑셀정리파일';
        const typeClass = uf.type === 'lab' ? 'lab' : 'summary';
        const sizeKB = (uf.file.size / 1024).toFixed(1);
        div.innerHTML = `
            <div class="file-item-icon">
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
                    <rect x="3" y="2" width="14" height="16" rx="2" stroke="currentColor" stroke-width="1.5" fill="none"/>
                    <path d="M7 7H13M7 10H11" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
                </svg>
            </div>
            <div class="file-item-info">
                <div class="file-item-name">${uf.name}</div>
                <div class="file-item-meta">${sizeKB} KB</div>
            </div>
            <span class="file-type-tag ${typeClass}">${typeLabel}</span>
            <button class="file-remove" data-idx="${idx}" title="제거">&times;</button>
        `;
        fileList.appendChild(div);
    });

    fileList.querySelectorAll('.file-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.idx);
            uploadedFiles.splice(idx, 1);
            renderFileList();
            if (uploadedFiles.length > 0) prescanFiles();
            else { sampleCard.style.display = 'none'; sensangCard.style.display = 'none'; updateProcessBtn(); }
        });
    });
}

async function prescanFiles() {
    if (uploadedFiles.length === 0) return;
    const formData = new FormData();
    uploadedFiles.forEach(uf => formData.append('files', uf.file));

    try {
        const resp = await fetch('/api/prescan', { method: 'POST', body: formData });
        const data = await resp.json();
        if (data.success) {
            const allSamples = [...new Set([...data.lab_samples, ...data.summary_samples])];
            sampleRows = allSamples.map(s => ({
                file_name: s,
                display_name: s,
                sheet_name: s.replace(/[^a-zA-Z0-9가-힣\-_ ]/g, '').substring(0, 20),
                raw_name: '',
            }));
            renderSampleTable();
            renderSensangTable();
            sampleCard.style.display = '';
            sensangCard.style.display = '';
            updateSummary(data);
        }
    } catch (err) {
        console.error('Prescan error:', err);
    }
    updateProcessBtn();
}

// ── Sample Table ──
function renderSampleTable() {
    sampleTbody.innerHTML = '';
    sampleRows.forEach((sr, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" value="${sr.file_name}" data-idx="${idx}" data-field="file_name"></td>
            <td><input type="text" value="${sr.display_name}" data-idx="${idx}" data-field="display_name"></td>
            <td><input type="text" value="${sr.sheet_name}" data-idx="${idx}" data-field="sheet_name"></td>
            <td><input type="text" value="${sr.raw_name}" data-idx="${idx}" data-field="raw_name" placeholder="없으면 비워두세요"></td>
            <td><button class="btn-remove-row" data-idx="${idx}">&times;</button></td>
        `;
        sampleTbody.appendChild(tr);
    });

    sampleTbody.querySelectorAll('input').forEach(inp => {
        inp.addEventListener('input', (e) => {
            const idx = parseInt(e.target.dataset.idx);
            const field = e.target.dataset.field;
            sampleRows[idx][field] = e.target.value;
            if (field === 'display_name') renderSensangTable();
            updateProcessBtn();
        });
    });

    sampleTbody.querySelectorAll('.btn-remove-row').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.idx);
            sampleRows.splice(idx, 1);
            renderSampleTable();
            renderSensangTable();
            updateProcessBtn();
        });
    });
}

btnAddSample.addEventListener('click', () => {
    sampleRows.push({ file_name: '', display_name: '', sheet_name: '', raw_name: '' });
    renderSampleTable();
    renderSensangTable();
});

// ── Sensang Table ──
function renderSensangTable() {
    sensangTbody.innerHTML = '';
    sampleRows.forEach((sr, idx) => {
        const name = sr.display_name || sr.file_name || `시료 ${idx + 1}`;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight:500;font-size:12px;white-space:nowrap;">${name}</td>
            <td><input type="number" step="0.01" data-sample="${name}" data-field="pH"></td>
            <td><input type="number" step="0.001" data-sample="${name}" data-field="탁도"></td>
            <td><input type="number" step="0.01" data-sample="${name}" data-field="색도_L"></td>
            <td><input type="number" step="0.01" data-sample="${name}" data-field="색도_a"></td>
            <td><input type="number" step="0.01" data-sample="${name}" data-field="색도_b"></td>
        `;
        sensangTbody.appendChild(tr);
    });
}

function collectSensangData() {
    const data = {};
    sensangTbody.querySelectorAll('input').forEach(inp => {
        const sample = inp.dataset.sample;
        const field = inp.dataset.field;
        if (!data[sample]) data[sample] = {};
        data[sample][field] = inp.value;
    });
    return data;
}

// ── Process ──
function updateProcessBtn() {
    const hasFiles = uploadedFiles.length > 0;
    const hasSamples = sampleRows.length > 0 && sampleRows.some(s => s.display_name.trim());
    btnProcess.disabled = !(hasFiles && hasSamples);
}

btnProcess.addEventListener('click', async () => {
    btnProcess.disabled = true;
    statusArea.style.display = '';
    statusLoading.style.display = 'flex';
    statusSuccess.style.display = 'none';
    statusError.style.display = 'none';
    btnDownload.style.display = 'none';

    const formData = new FormData();
    uploadedFiles.forEach(uf => formData.append('files', uf.file));

    const sampleConfig = sampleRows.filter(s => s.display_name.trim()).map(s => ({
        display_name: s.display_name.trim(),
        sheet_name: s.sheet_name.trim() || s.display_name.trim(),
        raw_material_name: s.raw_name.trim(),
        file_names: [s.file_name.trim(), s.display_name.trim()].filter(Boolean),
    }));

    formData.append('sample_config_json', JSON.stringify(sampleConfig));
    formData.append('sensang_data_json', JSON.stringify(collectSensangData()));
    formData.append('batch_date', $('#batch-date').value || '');

    try {
        const resp = await fetch('/api/process', { method: 'POST', body: formData });
        const data = await resp.json();

        statusLoading.style.display = 'none';

        if (data.success) {
            resultFileId = data.file_id;
            resultFilename = data.filename;
            statusSuccess.style.display = 'flex';
            statusMessage.textContent = data.message;
            btnDownload.style.display = '';
            btnDownload.disabled = false;
            updateSummaryResult(data.summary);
        } else {
            statusError.style.display = 'flex';
            errorMessage.textContent = data.detail || '처리 중 오류가 발생했습니다.';
        }
    } catch (err) {
        statusLoading.style.display = 'none';
        statusError.style.display = 'flex';
        errorMessage.textContent = `네트워크 오류: ${err.message}`;
    }

    updateProcessBtn();
});

btnDownload.addEventListener('click', () => {
    if (!resultFileId) return;
    const url = `/api/download/${resultFileId}?filename=${encodeURIComponent(resultFilename || 'output.xlsx')}`;
    const a = document.createElement('a');
    a.href = url;
    a.download = resultFilename || 'output.xlsx';
    a.click();
});

// ── Summary ──
function updateSummary(prescanData) {
    summaryPlaceholder.style.display = 'none';
    summaryContent.style.display = '';
    const allSamples = [...new Set([...prescanData.lab_samples, ...prescanData.summary_samples])];
    summaryGrid.innerHTML = `
        <div class="summary-item">
            <span class="summary-item-label">업로드 파일 수</span>
            <span class="summary-item-value">${prescanData.files.length}개</span>
        </div>
        <div class="summary-item">
            <span class="summary-item-label">의뢰품검사상세 파일</span>
            <span class="summary-item-value">${prescanData.files.filter(f => f.type === 'lab').length}개</span>
        </div>
        <div class="summary-item">
            <span class="summary-item-label">엑셀정리파일</span>
            <span class="summary-item-value">${prescanData.files.filter(f => f.type === 'summary').length}개</span>
        </div>
        <div class="summary-item">
            <span class="summary-item-label">감지된 시료</span>
            <span class="summary-item-value highlight">${allSamples.length}개</span>
        </div>
        <div class="summary-item">
            <span class="summary-item-label">시료 목록</span>
            <span class="summary-item-value" style="font-size:11px;text-align:right;max-width:200px;">${allSamples.join(', ')}</span>
        </div>
    `;
}

function updateSummaryResult(summary) {
    if (!summary) return;
    summaryGrid.innerHTML += `
        <div class="summary-item" style="border-top:1px solid var(--border);padding-top:12px;margin-top:4px;">
            <span class="summary-item-label">생성된 시료 시트</span>
            <span class="summary-item-value highlight">${summary.sample_count}개</span>
        </div>
        <div class="summary-item">
            <span class="summary-item-label">Lab 데이터 레코드</span>
            <span class="summary-item-value">${summary.lab_records}건</span>
        </div>
        <div class="summary-item">
            <span class="summary-item-label">정리파일 레코드</span>
            <span class="summary-item-value">${summary.summary_records}건</span>
        </div>
    `;
}
