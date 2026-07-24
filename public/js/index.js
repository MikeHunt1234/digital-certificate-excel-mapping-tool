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

// Mode toggle buttons
const modernBtn = document.getElementById('modernModeBtn');
const legacyBtn = document.getElementById('legacyModeBtn');
const enrichBtn = document.getElementById('enrichModeBtn');
const modernContent = document.getElementById('modernContent');
const legacyContent = document.getElementById('legacyContent');
const enrichContent = document.getElementById('enrichContent');

// State
let selectedFiles = [];
let fileSummaries = [];

// ============ MODE SWITCHING ============

// Modern mode
if (modernBtn) {
    modernBtn.addEventListener('click', function() {
        setActiveMode('modern');
    });
}

// Legacy mode
if (legacyBtn) {
    legacyBtn.addEventListener('click', function() {
        setActiveMode('legacy');
    });
}

// Enrich mode
if (enrichBtn) {
    enrichBtn.addEventListener('click', function() {
        setActiveMode('enrich');
    });
}

/**
 * Set active mode
 */
function setActiveMode(mode) {
    // Hide all content
    if (modernContent) modernContent.style.display = 'none';
    if (legacyContent) legacyContent.style.display = 'none';
    if (enrichContent) enrichContent.style.display = 'none';
    
    // Remove all active classes
    if (modernBtn) modernBtn.classList.remove('active');
    if (legacyBtn) legacyBtn.classList.remove('active');
    if (enrichBtn) enrichBtn.classList.remove('active');
    
    // Show selected
    if (mode === 'modern') {
        if (modernContent) modernContent.style.display = 'block';
        if (modernBtn) modernBtn.classList.add('active');
    } else if (mode === 'legacy') {
        if (legacyContent) legacyContent.style.display = 'block';
        if (legacyBtn) legacyBtn.classList.add('active');
        // Initialize legacy mode if needed
        if (typeof loadLegacyContent === 'function') {
            loadLegacyContent();
        }
    } else if (mode === 'enrich') {
        if (enrichContent) enrichContent.style.display = 'block';
        if (enrichBtn) enrichBtn.classList.add('active');
        // Initialize enrichment mode
        if (typeof initEnrichmentMode === 'function') {
            initEnrichmentMode();
        }
    }
}

// ============ MODERN MODE - Upload & Process ============

// Upload area click handler
if (uploadArea) {
    uploadArea.addEventListener('click', () => {
        if (fileInput) fileInput.click();
    });
}

// File input change handler
if (fileInput) {
    fileInput.addEventListener('change', (e) => {
        selectedFiles = Array.from(e.target.files);
        if (fileCountSpan) {
            fileCountSpan.textContent = `${selectedFiles.length} selected`;
        }
        
        if (selectedFiles.length > 0) {
            if (processBtn) processBtn.disabled = false;
            showFilePreview(selectedFiles);
        } else {
            if (processBtn) processBtn.disabled = true;
            if (previewSection) previewSection.style.display = 'none';
        }
    });
}

// Drag and drop handlers
if (uploadArea) {
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
            if (fileCountSpan) {
                fileCountSpan.textContent = `${selectedFiles.length} selected`;
            }
            if (processBtn) processBtn.disabled = false;
            showFilePreview(selectedFiles);
        }
    });
}

// Show preview of selected files
function showFilePreview(files) {
    if (!previewContent || !previewSection) return;
    
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
if (processBtn) {
    processBtn.addEventListener('click', async () => {
        if (selectedFiles.length === 0) {
            showError('Please select at least one file');
            return;
        }
        
        if (loadingDiv) loadingDiv.style.display = 'block';
        if (errorDiv) errorDiv.style.display = 'none';
        if (successDiv) successDiv.style.display = 'none';
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
            
            if (successDiv) {
                successDiv.textContent = '✅ Files processed successfully!';
                successDiv.style.display = 'block';
                setTimeout(() => {
                    successDiv.style.display = 'none';
                }, 3000);
            }
        } catch (error) {
            showError(error.message);
        } finally {
            if (loadingDiv) loadingDiv.style.display = 'none';
            processBtn.disabled = false;
        }
    });
}

// ============ UTILITY FUNCTIONS ============

function showError(message) {
    if (errorDiv) {
        errorDiv.textContent = message;
        errorDiv.style.display = 'block';
        setTimeout(() => {
            errorDiv.style.display = 'none';
        }, 5000);
    }
}

function showSuccess(message) {
    if (successDiv) {
        successDiv.textContent = message;
        successDiv.style.display = 'block';
        setTimeout(() => {
            successDiv.style.display = 'none';
        }, 3000);
    }
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ============ INITIALIZATION ============

// Set default mode to modern
setActiveMode('modern');