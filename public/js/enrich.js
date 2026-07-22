// Enrich Mode JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // Mode toggle - will be handled by the main index.js, but we keep this for the enrichment-specific logic
    
    // We'll initialize enrichment when the enrich mode button is clicked
    // The main index.js will call initEnrichmentMode() when switching to enrich mode
});

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

                <!-- Source File -->
                <div class="upload-group">
                    <label class="upload-label">📋 Source File (data source)</label>
                    <p class="upload-hint">Participant registry file containing Giới tính, Dân tộc, Nơi sinh</p>
                    <div class="upload-area enrich-upload-area" id="sourceUploadArea">
                        <div class="upload-icon">📋</div>
                        <p>Drop source Excel file here or click to select</p>
                        <p class="small-text">.xlsx, .xls supported</p>
                        <input type="file" id="sourceFileInput" accept=".xlsx,.xls">
                    </div>
                    <div id="sourceFileStatus" class="file-status"></div>
                </div>
            </div>

            <div id="enrichStatusBar" class="status-bar">
                <div class="status-item">
                    <span class="status-label">Target File:</span>
                    <span id="targetFileStatusText" class="status-value">Not selected</span>
                </div>
                <div class="status-item">
                    <span class="status-label">Source File:</span>
                    <span id="sourceFileStatusText" class="status-value">Not selected</span>
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

let enrichTargetFile = null;
let enrichSourceFile = null;

/**
 * Setup enrichment event listeners
 */
function setupEnrichEvents() {
    // Target file upload
    const targetArea = document.getElementById('targetUploadArea');
    const targetInput = document.getElementById('targetFileInput');
    
    // Source file upload
    const sourceArea = document.getElementById('sourceUploadArea');
    const sourceInput = document.getElementById('sourceFileInput');
    
    // Setup both upload areas
    if (targetArea && targetInput) {
        setupEnrichUpload(targetArea, targetInput, 'target');
    }
    
    if (sourceArea && sourceInput) {
        setupEnrichUpload(sourceArea, sourceInput, 'source');
    }
    
    // Enrich button
    const enrichBtn = document.getElementById('enrichBtn');
    if (enrichBtn) {
        enrichBtn.addEventListener('click', handleEnrich);
    }
}

/**
 * Setup individual enrichment upload area
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
            // Update the input's files
            const dt = new DataTransfer();
            dt.items.add(file);
            input.files = dt.files;
            handleEnrichFileSelect(file, type);
        }
    });
}

/**
 * Handle file selection for enrichment
 */
function handleEnrichFileSelect(file, type) {
    // Validate file type
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) {
        showError('Please select an Excel file (.xlsx or .xls)');
        return;
    }
    
    // Validate file size (max 50MB)
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
    } else {
        enrichSourceFile = file;
        updateEnrichFileStatus('sourceFileStatus', file, true);
        const statusText = document.getElementById('sourceFileStatusText');
        if (statusText) statusText.textContent = `✅ ${file.name}`;
        const area = document.getElementById('sourceUploadArea');
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
    
    if (enrichTargetFile && enrichSourceFile) {
        btn.disabled = false;
        statusEl.textContent = '✅ Ready to enrich';
        statusEl.style.color = '#2ecc71';
    } else if (enrichTargetFile) {
        btn.disabled = true;
        statusEl.textContent = '⏳ Waiting for source file';
        statusEl.style.color = '#f39c12';
    } else if (enrichSourceFile) {
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
 * Handle enrichment processing
 */
async function handleEnrich() {
    if (!enrichTargetFile || !enrichSourceFile) {
        showError('Please select both target and source files');
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
    formData.append('sourceFile', enrichSourceFile);
    
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
            successDiv.textContent = '✅ Enrichment complete! File downloaded successfully.';
            successDiv.style.display = 'block';
        }
        
        // Reset after success
        setTimeout(() => {
            resetEnrichmentState();
        }, 3000);
        
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
    enrichSourceFile = null;
    
    // Reset file inputs
    const targetInput = document.getElementById('targetFileInput');
    const sourceInput = document.getElementById('sourceFileInput');
    if (targetInput) targetInput.value = '';
    if (sourceInput) sourceInput.value = '';
    
    // Reset status displays
    const targetStatus = document.getElementById('targetFileStatus');
    const sourceStatus = document.getElementById('sourceFileStatus');
    const targetStatusText = document.getElementById('targetFileStatusText');
    const sourceStatusText = document.getElementById('sourceFileStatusText');
    const targetArea = document.getElementById('targetUploadArea');
    const sourceArea = document.getElementById('sourceUploadArea');
    
    if (targetStatus) targetStatus.innerHTML = '';
    if (sourceStatus) sourceStatus.innerHTML = '';
    if (targetStatusText) targetStatusText.textContent = 'Not selected';
    if (sourceStatusText) sourceStatusText.textContent = 'Not selected';
    if (targetArea) targetArea.classList.remove('has-file');
    if (sourceArea) sourceArea.classList.remove('has-file');
    
    updateEnrichButtonState();
}

// Helper functions (reuse from legacy.js or define here)
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