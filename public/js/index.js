// DOM Elements
const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const processBtn = document.getElementById('processBtn');
const loadingDiv = document.getElementById('loading');
const errorDiv = document.getElementById('error');
const successDiv = document.getElementById('success');
const fileCountSpan = document.getElementById('fileCount');
const recordCountSpan = document.getElementById('recordCount');
const previewSection = document.getElementById('previewSection');
const previewContent = document.getElementById('previewContent');

// State
let selectedFiles = [];
let fileSummaries = [];

// Upload area click handler
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

// File input change handler
fileInput.addEventListener('change', (e) => {
    selectedFiles = Array.from(e.target.files);
    fileCountSpan.textContent = `${selectedFiles.length} selected`;
    
    if (selectedFiles.length > 0) {
        processBtn.disabled = false;
        showFilePreview(selectedFiles);
    } else {
        processBtn.disabled = true;
        previewSection.style.display = 'none';
    }
});

// Drag and drop handlers
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', async (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
        selectedFiles = files;
        fileCountSpan.textContent = `${selectedFiles.length} selected`;
        processBtn.disabled = false;
        showFilePreview(selectedFiles);
    }
});

// Show preview of selected files
function showFilePreview(files) {
    const fileList = files.map(file => `
        <div class="file-item">
            <span class="file-name">📄 ${escapeHtml(file.name)}</span>
            <span class="file-size">(${(file.size / 1024).toFixed(1)} KB)</span>
        </div>
    `).join('');
    
    previewContent.innerHTML = `
        <div class="file-summary">
            <p><strong>Selected ${files.length} file(s):</strong></p>
            <div class="file-list">${fileList}</div>
            <p class="small-text mt-2">Click "Process & Export" to merge all files into one Excel document.</p>
        </div>
    `;
    previewSection.style.display = 'block';
}

// Process button handler
processBtn.addEventListener('click', async () => {
    if (selectedFiles.length === 0) {
        showError('Please select at least one file');
        return;
    }
    
    loadingDiv.style.display = 'block';
    errorDiv.style.display = 'none';
    successDiv.style.display = 'none';
    processBtn.disabled = true;
    
    const formData = new FormData();
    selectedFiles.forEach(file => {
        formData.append('userFiles', file);
    });
    
    try {
        const response = await fetch('/api/merge-multiple', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Processing failed');
        }
        
        // Get filename from server
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = 'merged_certificates.xlsx'; // fallback

        if (contentDisposition) {
            const match = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
            if (match && match[1]) {
                filename = match[1].replace(/['"]/g, '');
                try {
                    filename = decodeURIComponent(filename);
                } catch (e) {}
            }
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (error) {
        showError(error.message);
    } finally {
        loadingDiv.style.display = 'none';
        processBtn.disabled = false;
    }
});

function showError(message) {
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
    setTimeout(() => {
        errorDiv.style.display = 'none';
    }, 5000);
}

function showSuccess(message) {
    successDiv.textContent = message;
    successDiv.style.display = 'block';
    setTimeout(() => {
        successDiv.style.display = 'none';
    }, 3000);
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}