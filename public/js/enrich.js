// Enrich Mode JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // Mode toggle - will be handled by the main index.js
});

let enrichTargetFile = null;
let enrichSourceFiles = []; // Changed from single file to array

/**
 * Initialize enrichment mode - called from index.js when enrich mode is activated
 */
function initEnrichmentMode() {
    // Only initialize if not already initialized
    if (document.getElementById('enrichContentPlaceholder') && 
        document.getElementById('enrichContentPlaceholder').innerHTML === '') {
        loadEnrichContent();
    }
    
    setupEnrichEvents();
}

/**
 * Load enrichment HTML content into the placeholder
 */
function loadEnrichContent() {
    const placeholder = document.getElementById('enrichContentPlaceholder');
    if (!placeholder) return;
    
    placeholder.innerHTML = `
        <div class="step">
            <h3>🔗 Data Enrichment</h3>
            <p class="step-description">Fill missing fields (Giới tính, Dân tộc, Nơi sinh) in your certificate file using participant registry data</p>
            
            <div class="enrich-upload-section">
                <!-- Target File -->
                <div class="upload-group">
                    <label class="upload-label">📄 Target File (to be enriched)</label>
                    <p class="upload-hint">Certificate file with missing fields (Template format)</p>
                    <div class="upload-area enrich-upload-area" id="targetUploadArea">
                        <div class="upload-icon">📄</div>
                        <p>Drop target Excel file here or click to select</p>
                        <p class="small-text">.xlsx, .xls supported</p>
                        <input type="file" id="targetFileInput" accept=".xlsx,.xls">
                    </div>
                    <div id="targetFileStatus" class="file-status"></div>
                </div>

                <!-- Source Files - MULTIPLE -->
                <div class="upload-group">
                    <label class="upload-label">📋 Source Files (data source)</label>
                    <p class="upload-hint">Participant registry files containing Giới tính, Dân tộc, Nơi sinh (can select multiple)</p>
                    <div class="upload-area enrich-upload-area" id="sourceUploadArea">
                        <div class="upload-icon">📋</div>
                        <p>Drop Excel file(s) here or click to select</p>
                        <p class="small-text">.xlsx, .xls supported | Multiple files can be selected</p>
                        <input type="file" id="sourceFileInput" accept=".xlsx,.xls" multiple>
                    </div>
                    <div id="sourceFilesStatus" class="file-status"></div>
                    <div id="sourceFilesList" class="file-list-container"></div>
                </div>
            </div>
            <div id="enrichStatusBar" class="status-bar">
                <div class="status-item">
                    <span class="status-label">Target File:</span>
                    <span id="targetFileStatusText" class="status-value">Not selected</span>
                </div>
                <div class="status-item">
                    <span class="status-label">Source Files:</span>
                    <span id="sourceFilesCount" class="status-value">0 selected</span>
                </div>
                <div class="status-item">
                    <span class="status-label">Status:</span>
                    <span id="enrichReadyStatus" class="status-value">⏳ Waiting for files</span>
                </div>
            </div>

            <div id="enrichProcessSection">
                <button id="enrichBtn" class="btn-primary" disabled>🔗 Enrich & Export</button>
            </div>
        </div>
    `;
}

/**
 * Setup enrichment event listeners
 */
function setupEnrichEvents() {
    // Target file upload
    const targetArea = document.getElementById('targetUploadArea');
    const targetInput = document.getElementById('targetFileInput');
    
    // Source files upload (multiple)
    const sourceArea = document.getElementById('sourceUploadArea');
    const sourceInput = document.getElementById('sourceFileInput');
    
    // Setup target upload
    if (targetArea && targetInput) {
        setupEnrichUpload(targetArea, targetInput, 'target');
    }
    
    // Setup source upload (multiple)
    if (sourceArea && sourceInput) {
        setupMultipleEnrichUpload(sourceArea, sourceInput);
    }
    
    // Enrich button
    const enrichBtn = document.getElementById('enrichBtn');
    if (enrichBtn) {
        enrichBtn.addEventListener('click', handleEnrich);
    }
}

/**
 * Setup single file upload for target
 */
function setupEnrichUpload(area, input, type) {
    // Click to open file dialog
    area.addEventListener('click', () => input.click());
    
    // File selection
    input.addEventListener('change', (e) => {
        if (input.files.length > 0) {
            handleEnrichFileSelect(input.files[0], type);
        }
    });
    
    // Drag and drop
    area.addEventListener('dragover', (e) => {
        e.preventDefault();
        area.classList.add('drag-over');
    });
    
    area.addEventListener('dragleave', () => {
        area.classList.remove('drag-over');
    });
    
    area.addEventListener('drop', (e) => {
        e.preventDefault();
        area.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
            const file = e.dataTransfer.files[0];
            const dt = new DataTransfer();
            dt.items.add(file);
            input.files = dt.files;
            handleEnrichFileSelect(file, type);
        }
    });
}

/**
 * Setup multiple file upload for source files
 */
function setupMultipleEnrichUpload(area, input) {
    // Click to open file dialog
    area.addEventListener('click', () => input.click());
    
    // File selection
    input.addEventListener('change', (e) => {
        if (input.files.length > 0) {
            handleMultipleSourceFiles(input.files);
        }
    });
    
    // Drag and drop
    area.addEventListener('dragover', (e) => {
        e.preventDefault();
        area.classList.add('drag-over');
    });
    
    area.addEventListener('dragleave', () => {
        area.classList.remove('drag-over');
    });
    
    area.addEventListener('drop', (e) => {
        e.preventDefault();
        area.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
            // Update the input's files
            const dt = new DataTransfer();
            Array.from(e.dataTransfer.files).forEach(f => dt.items.add(f));
            input.files = dt.files;
            handleMultipleSourceFiles(e.dataTransfer.files);
        }
    });
}

/**
 * Handle multiple source files selection
 */
function handleMultipleSourceFiles(files) {
    const validFiles = [];
    let invalidFiles = [];
    
    // Validate each file
    Array.from(files).forEach(file => {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls'].includes(ext)) {
            invalidFiles.push(file.name);
            return;
        }
        if (file.size > 50 * 1024 * 1024) {
            invalidFiles.push(`${file.name} (exceeds 50MB)`);
            return;
        }
        validFiles.push(file);
    });
    
    if (invalidFiles.length > 0) {
        showError(`Invalid files skipped: ${invalidFiles.join(', ')}`);
    }
    
    if (validFiles.length === 0) {
        showError('Please select valid Excel files');
        return;
    }
    
    enrichSourceFiles = validFiles;
    updateSourceFilesUI();
    updateEnrichButtonState();
    
    // Store files in input for form data
    const dt = new DataTransfer();
    validFiles.forEach(f => dt.items.add(f));
    document.getElementById('sourceFileInput').files = dt.files;
}

/**
 * Update source files UI with list
 */
function updateSourceFilesUI() {
    const statusContainer = document.getElementById('sourceFilesStatus');
    const listContainer = document.getElementById('sourceFilesList');
    const countSpan = document.getElementById('sourceFilesCount');
    
    if (!statusContainer || !listContainer) return;
    
    if (enrichSourceFiles.length === 0) {
        statusContainer.innerHTML = '';
        listContainer.innerHTML = '';
        if (countSpan) countSpan.textContent = '0 selected';
        return;
    }
    
    // Update status
    const totalSize = enrichSourceFiles.reduce((sum, f) => sum + f.size, 0);
    const totalSizeKB = (totalSize / 1024).toFixed(1);
    statusContainer.innerHTML = `
        <span class="status-icon success">✅</span>
        <span class="file-name">${enrichSourceFiles.length} file(s) selected</span>
        <span class="file-size">(${totalSizeKB} KB total)</span>
    `;
    
    if (countSpan) countSpan.textContent = `${enrichSourceFiles.length} selected`;
    
    // Build file list
    let listHTML = `<div class="file-list-header">📁 Selected source files:</div><ul class="file-list">`;
    enrichSourceFiles.forEach((file, index) => {
        const size = (file.size / 1024).toFixed(1);
        listHTML += `
            <li class="file-list-item">
                <span class="file-index">${index + 1}.</span>
                <span class="file-name">📄 ${escapeHtml(file.name)}</span>
                <span class="file-size">(${size} KB)</span>
                <button class="file-remove-btn" data-index="${index}" title="Remove this file">✕</button>
            </li>
        `;
    });
    listHTML += `</ul>`;
    listHTML += `
        <div class="file-list-actions">
            <button id="clearAllSourceFiles" class="btn-small btn-danger">🗑️ Clear all</button>
        </div>
    `;
    listContainer.innerHTML = listHTML;
    
    // Add remove event listeners
    document.querySelectorAll('.file-remove-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const index = parseInt(this.dataset.index);
            removeSourceFile(index);
        });
    });
    
    const clearBtn = document.getElementById('clearAllSourceFiles');
    if (clearBtn) {
        clearBtn.addEventListener('click', clearAllSourceFiles);
    }
}

/**
 * Remove a single source file
 */
function removeSourceFile(index) {
    if (index >= 0 && index < enrichSourceFiles.length) {
        enrichSourceFiles.splice(index, 1);
        // Update the file input
        const dt = new DataTransfer();
        enrichSourceFiles.forEach(f => dt.items.add(f));
        document.getElementById('sourceFileInput').files = dt.files;
        updateSourceFilesUI();
        updateEnrichButtonState();
    }
}

/**
 * Clear all source files
 */
function clearAllSourceFiles() {
    enrichSourceFiles = [];
    document.getElementById('sourceFileInput').value = '';
    updateSourceFilesUI();
    updateEnrichButtonState();
}

/**
 * Handle single file selection for target
 */
function handleEnrichFileSelect(file, type) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) {
        showError('Please select an Excel file (.xlsx or .xls)');
        return;
    }
    
    if (file.size > 50 * 1024 * 1024) {
        showError('File size exceeds 50MB limit');
        return;
    }
    
    if (type === 'target') {
        enrichTargetFile = file;
        updateEnrichFileStatus('targetFileStatus', file, true);
        const statusText = document.getElementById('targetFileStatusText');
        if (statusText) statusText.textContent = `✅ ${file.name}`;
        const area = document.getElementById('targetUploadArea');
        if (area) area.classList.add('has-file');
    }
    
    updateEnrichButtonState();
}

/**
 * Update file status display
 */
function updateEnrichFileStatus(elementId, file, success) {
    const container = document.getElementById(elementId);
    if (!container) return;
    
    if (success) {
        const size = (file.size / 1024).toFixed(1);
        container.innerHTML = `
            <span class="status-icon success">✅</span>
            <span class="file-name">${file.name}</span>
            <span class="file-size">(${size} KB)</span>
        `;
    } else {
        container.innerHTML = `
            <span class="status-icon error">❌</span>
            <span>Upload failed</span>
        `;
    }
}

/**
 * Update enrich button state
 */
function updateEnrichButtonState() {
    const btn = document.getElementById('enrichBtn');
    const statusEl = document.getElementById('enrichReadyStatus');
    
    if (!btn || !statusEl) return;
    
    if (enrichTargetFile && enrichSourceFiles.length > 0) {
        btn.disabled = false;
        statusEl.textContent = `✅ Ready to enrich (${enrichSourceFiles.length} source file(s))`;
        statusEl.style.color = '#2ecc71';
    } else if (enrichTargetFile) {
        btn.disabled = true;
        statusEl.textContent = '⏳ Waiting for source file(s)';
        statusEl.style.color = '#f39c12';
    } else if (enrichSourceFiles.length > 0) {
        btn.disabled = true;
        statusEl.textContent = '⏳ Waiting for target file';
        statusEl.style.color = '#f39c12';
    } else {
        btn.disabled = true;
        statusEl.textContent = '⏳ Waiting for files';
        statusEl.style.color = '#7f8c8d';
    }
}

/**
 * Handle enrichment processing with multiple source files
 */
async function handleEnrich() {
    if (!enrichTargetFile) {
        showError('Please select a target file');
        return;
    }
    
    if (enrichSourceFiles.length === 0) {
        showError('Please select at least one source file');
        return;
    }
    
    const enrichBtn = document.getElementById('enrichBtn');
    const loadingDiv = document.getElementById('loading');
    const errorDiv = document.getElementById('error');
    const successDiv = document.getElementById('success');
    
    // Disable button and show loading
    if (enrichBtn) enrichBtn.disabled = true;
    if (loadingDiv) loadingDiv.style.display = 'block';
    if (errorDiv) errorDiv.style.display = 'none';
    if (successDiv) successDiv.style.display = 'none';
    
    const formData = new FormData();
    formData.append('targetFile', enrichTargetFile);
    
    // Append all source files
    enrichSourceFiles.forEach(file => {
        formData.append('sourceFiles', file);
    });
    
    try {
        const response = await fetch('/api/enrich', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Enrichment failed');
        }
        
        // Get the blob from response
        const blob = await response.blob();
        
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        // Extract filename from Content-Disposition header
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = 'enriched_file.xlsx';
        if (contentDisposition) {
            const match = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
            if (match && match[1]) {
                filename = match[1].replace(/['"]/g, '');
                try {
                    filename = decodeURIComponent(filename);
                } catch (e) {}
            }
        }
        
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        // Show success
        if (successDiv) {
            successDiv.textContent = `✅ Enrichment complete! ${enrichSourceFiles.length} source file(s) processed. File downloaded successfully.`;
            successDiv.style.display = 'block';
        }
        
        // Reset after success
        setTimeout(() => {
            resetEnrichmentState();
        }, 5000);
        
    } catch (error) {
        if (errorDiv) {
            errorDiv.textContent = `❌ Error: ${error.message}`;
            errorDiv.style.display = 'block';
        }
        console.error('Enrichment error:', error);
    } finally {
        if (loadingDiv) loadingDiv.style.display = 'none';
        if (enrichBtn) enrichBtn.disabled = false;
    }
}

/**
 * Reset enrichment state
 */
function resetEnrichmentState() {
    enrichTargetFile = null;
    enrichSourceFiles = [];
    
    // Reset file inputs
    const targetInput = document.getElementById('targetFileInput');
    const sourceInput = document.getElementById('sourceFileInput');
    if (targetInput) targetInput.value = '';
    if (sourceInput) sourceInput.value = '';
    
    // Reset status displays
    const targetStatus = document.getElementById('targetFileStatus');
    const sourceStatus = document.getElementById('sourceFilesStatus');
    const sourceList = document.getElementById('sourceFilesList');
    const targetStatusText = document.getElementById('targetFileStatusText');
    const sourceCount = document.getElementById('sourceFilesCount');
    const targetArea = document.getElementById('targetUploadArea');
    const sourceArea = document.getElementById('sourceUploadArea');
    
    if (targetStatus) targetStatus.innerHTML = '';
    if (sourceStatus) sourceStatus.innerHTML = '';
    if (sourceList) sourceList.innerHTML = '';
    if (targetStatusText) targetStatusText.textContent = 'Not selected';
    if (sourceCount) sourceCount.textContent = '0 selected';
    if (targetArea) targetArea.classList.remove('has-file');
    if (sourceArea) sourceArea.classList.remove('has-file');
    
    updateEnrichButtonState();
}

// Helper functions
function showError(message) {
    const errorDiv = document.getElementById('error');
    if (errorDiv) {
        errorDiv.textContent = message;
        errorDiv.style.display = 'block';
        setTimeout(() => {
            errorDiv.style.display = 'none';
        }, 5000);
    }
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}