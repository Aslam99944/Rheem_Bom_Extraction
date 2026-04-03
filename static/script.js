/**
 * BOM Extraction POC — Frontend Logic
 * Handles file upload, API calls, result rendering, and Excel download.
 */

// --- Elements ---
const dropZone      = document.getElementById('dropZone');
const fileInput      = document.getElementById('fileInput');
const fileInfo       = document.getElementById('fileInfo');
const fileName       = document.getElementById('fileName');
const fileSize       = document.getElementById('fileSize');
const removeFileBtn  = document.getElementById('removeFile');
const extractBtn     = document.getElementById('extractBtn');
const uploadSection  = document.getElementById('uploadSection');
const loadingSection = document.getElementById('loadingSection');
const resultsSection = document.getElementById('resultsSection');
const bomTableBody   = document.getElementById('bomTableBody');
const downloadBtn    = document.getElementById('downloadBtn');
const newUploadBtn   = document.getElementById('newUploadBtn');
const toggleOcrBtn   = document.getElementById('toggleOcr');
const ocrTextEl      = document.getElementById('ocrText');

// Pipeline step elements
const pSteps = [
    document.getElementById('step1'),
    document.getElementById('step2'),
    document.getElementById('step3'),
    document.getElementById('step4'),
];

let selectedFile = null;
let excelFilename = null;

// --- Helpers ---
function formatBytes(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

function setPipelineStep(activeIdx) {
    pSteps.forEach((el, i) => {
        el.classList.remove('active', 'done');
        if (i < activeIdx)  el.classList.add('done');
        if (i === activeIdx) el.classList.add('active');
    });
}

function animateLoadingSteps() {
    const ls1 = document.getElementById('ls1');
    const ls2 = document.getElementById('ls2');
    const ls3 = document.getElementById('ls3');

    ls1.className = 'loading-step active';
    ls2.className = 'loading-step';
    ls3.className = 'loading-step';

    setTimeout(() => {
        ls1.className = 'loading-step done';
        ls1.textContent = '✅ Document analysis complete';
        ls2.className = 'loading-step active';
    }, 3000);

    setTimeout(() => {
        ls2.className = 'loading-step done';
        ls2.textContent = '✅ BOM field extraction complete';
        ls3.className = 'loading-step active';
    }, 8000);
}

// --- Drag & Drop ---
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) {
        handleFile(e.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', () => {
    if (fileInput.files.length) {
        handleFile(fileInput.files[0]);
    }
});

function handleFile(file) {
    const allowed = ['.pdf', '.png', '.jpg', '.jpeg', '.tiff', '.tif'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!allowed.includes(ext)) {
        alert('Unsupported file type. Please use: ' + allowed.join(', '));
        return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatBytes(file.size);
    fileInfo.style.display = 'block';
    extractBtn.disabled = false;
    setPipelineStep(0);
}

removeFileBtn.addEventListener('click', () => {
    selectedFile = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    extractBtn.disabled = true;
    pSteps.forEach(el => el.classList.remove('active', 'done'));
});

// --- Extract ---
extractBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    // Show loading
    uploadSection.style.display = 'none';
    loadingSection.style.display = 'block';
    resultsSection.style.display = 'none';
    setPipelineStep(1);
    animateLoadingSteps();

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
        const resp = await fetch('/upload', { method: 'POST', body: formData });
        if (!resp.ok) {
            const err = await resp.json();
            throw new Error(err.detail || 'Extraction failed');
        }
        const data = await resp.json();
        showResults(data);
    } catch (err) {
        alert('Error: ' + err.message);
        loadingSection.style.display = 'none';
        uploadSection.style.display = 'block';
    }
});

// --- Show Results ---
function showResults(data) {
    loadingSection.style.display = 'none';
    resultsSection.style.display = 'block';
    setPipelineStep(3);

    // All pipeline steps done
    pSteps.forEach(el => {
        el.classList.remove('active');
        el.classList.add('done');
    });

    // Token metrics
    const ti = data.token_info || {};
    document.getElementById('inputTokens').textContent  = (ti.input_tokens || 0).toLocaleString();
    document.getElementById('outputTokens').textContent = (ti.output_tokens || 0).toLocaleString();
    document.getElementById('totalTokens').textContent  = (ti.total_tokens || 0).toLocaleString();
    document.getElementById('estCost').textContent       = '$' + (ti.estimated_cost_usd || 0).toFixed(6);
    document.getElementById('rowCount').textContent      = data.row_count || 0;
    document.getElementById('procTime').textContent      = (data.processing_time_seconds || 0) + 's';

    // OCR text
    if (data.ocr_text) {
        ocrTextEl.textContent = data.ocr_text;
    } else {
        ocrTextEl.textContent = '(no OCR text available)';
    }

    // Table
    bomTableBody.innerHTML = '';
    const items = data.bom_items || [];
    items.forEach(item => {
        const tr = document.createElement('tr');
        ['item', 'part_number', 'manufacturer', 'description', 'qty', 'uom', 'commodity', 'type', 'notes'].forEach(key => {
            const td = document.createElement('td');
            td.textContent = item[key] || '';
            tr.appendChild(td);
        });
        bomTableBody.appendChild(tr);
    });

    // Excel download
    excelFilename = data.excel_filename;
}

// --- OCR Toggle ---
toggleOcrBtn.addEventListener('click', () => {
    const showing = ocrTextEl.style.display !== 'none';
    ocrTextEl.style.display = showing ? 'none' : 'block';
    toggleOcrBtn.textContent = showing ? 'Show' : 'Hide';
});

// --- Download Excel ---
downloadBtn.addEventListener('click', () => {
    if (!excelFilename) return;
    window.open('/download/' + encodeURIComponent(excelFilename), '_blank');
});

// --- New Upload ---
newUploadBtn.addEventListener('click', () => {
    resultsSection.style.display = 'none';
    uploadSection.style.display = 'block';
    selectedFile = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    extractBtn.disabled = true;
    bomTableBody.innerHTML = '';
    ocrTextEl.textContent = '';
    ocrTextEl.style.display = 'none';
    toggleOcrBtn.textContent = 'Show';
    pSteps.forEach(el => el.classList.remove('active', 'done'));
});
