// Legacy Mode JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // Mode toggle
    const modernBtn = document.getElementById('modernModeBtn');
    const legacyBtn = document.getElementById('legacyModeBtn');
    const modernContent = document.getElementById('modernContent');
    const legacyContent = document.getElementById('legacyContent');
    
    if (modernBtn && legacyBtn) {
        modernBtn.addEventListener('click', function() {
            modernBtn.classList.add('active');
            legacyBtn.classList.remove('active');
            modernContent.style.display = 'block';
            legacyContent.style.display = 'none';
        });
        
        legacyBtn.addEventListener('click', function() {
            legacyBtn.classList.add('active');
            modernBtn.classList.remove('active');
            modernContent.style.display = 'none';
            legacyContent.style.display = 'block';
            loadLegacyContent();
        });
    }
    
    // Load legacy content dynamically
    function loadLegacyContent() {
        if (document.getElementById('legacyContentPlaceholder').innerHTML === '') {
            document.getElementById('legacyContentPlaceholder').innerHTML = `
                <div class="step">
                    <h3>📜 Legacy Mode (2015 Data)</h3>
                    <p class="step-description">Upload multiple 2015 Excel files - they will be merged into one</p>
                    
                    <div class="legacy-override-form">
                        <h4>Override Values</h4>
                        <p class="small-text">These values will be applied to all records from all files</p>
                        <div class="form-grid">
                            <div class="form-group">
                                <label for="legacyHoiDongThi">Hội đồng thi</label>
                                <input type="text" id="legacyHoiDongThi" value="Đại học Cần Thơ">
                            </div>
                            <div class="form-group">
                                <label for="legacyTenCoQuan">Tên cơ quan cấp bằng</label>
                                <input type="text" id="legacyTenCoQuan" value="Đại học Cần Thơ">
                            </div>
                            <div class="form-group">
                                <label for="legacyChucDanh">Chức danh người ký</label>
                                <input type="text" id="legacyChucDanh" value="Hiệu trưởng">
                            </div>
                            <div class="form-group">
                                <label for="legacyNguoiKy">Họ tên người ký</label>
                                <input type="text" id="legacyNguoiKy" placeholder="Nhập tên người ký">
                            </div>
                            <div class="form-group">
                                <label for="legacyDiaDanh">Địa danh</label>
                                <input type="text" id="legacyDiaDanh" value="thành phố Cần Thơ">
                            </div>
                            <div class="form-group">
                                <label for="legacyTrangThai">Trạng thái số hóa</label>
                                <input type="text" id="legacyTrangThai" value="Đã số hóa đầy đủ, chính xác">
                            </div>
                        </div>
                    </div>
                    
                    <div class="upload-area" id="legacyUploadArea">
                        <div class="upload-icon">📊</div>
                        <p>Drop Excel file(s) here or click to select</p>
                        <p class="small-text">.xlsx, .xls supported | Multiple files can be selected</p>
                        <input type="file" id="legacyFileInput" accept=".xlsx,.xls" multiple>
                    </div>
                </div>

                <div id="legacyStatusBar" class="status-bar">
                    <div class="status-item">
                        <span class="status-label">Files:</span>
                        <span id="legacyFileCount" class="status-value">0</span>
                    </div>
                    <div class="status-item">
                        <span class="status-label">Total Records:</span>
                        <span id="legacyRecordCount" class="status-value">0</span>
                    </div>
                </div>

                <div id="legacyProcessSection">
                    <button id="legacyProcessBtn" class="btn-primary" disabled>🔄 Process & Export</button>
                </div>
            `;
            
            // Initialize legacy event listeners
            initLegacyEvents();
        }
    }
    
    // Legacy event handlers
    function initLegacyEvents() {
        const fileInput = document.getElementById('legacyFileInput');
        const uploadArea = document.getElementById('legacyUploadArea');
        const processBtn = document.getElementById('legacyProcessBtn');
        const fileCountSpan = document.getElementById('legacyFileCount');
        const recordCountSpan = document.getElementById('legacyRecordCount');
        let selectedFiles = [];
        
        if (uploadArea) {
            uploadArea.addEventListener('click', () => {
                if (fileInput) fileInput.click();
            });
            
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
                if (files.length > 0 && fileInput) {
                    const dt = new DataTransfer();
                    files.forEach(f => dt.items.add(f));
                    fileInput.files = dt.files;
                    fileInput.dispatchEvent(new Event('change'));
                }
            });
        }
        
        if (fileInput) {
            fileInput.addEventListener('change', function() {
                selectedFiles = Array.from(this.files);
                if (selectedFiles.length > 0) {
                    fileCountSpan.textContent = selectedFiles.length;
                    // Preview the first file for record count (they should all have similar structure)
                    previewLegacyFiles(selectedFiles);
                } else {
                    fileCountSpan.textContent = '0';
                    recordCountSpan.textContent = '0';
                    processBtn.disabled = true;
                }
            });
        }
        
        // Preview files to show total record count
        async function previewLegacyFiles(files) {
            const recordCountSpan = document.getElementById('legacyRecordCount');
            const processBtn = document.getElementById('legacyProcessBtn');
            let totalRecords = 0;
            
            // Show loading indicator
            recordCountSpan.textContent = 'Loading...';
            processBtn.disabled = true;
            
            try {
                // Preview each file and sum up records
                for (const file of files) {
                    const formData = new FormData();
                    formData.append('userFile', file);
                    
                    const response = await fetch('/api/legacy/preview', {
                        method: 'POST',
                        body: formData
                    });
                    
                    const data = await response.json();
                    if (response.ok && data.recordCount) {
                        totalRecords += data.recordCount;
                    }
                }
                
                recordCountSpan.textContent = totalRecords;
                processBtn.disabled = false;
                
            } catch (error) {
                console.error('Preview error:', error);
                recordCountSpan.textContent = '0';
                processBtn.disabled = true;
            }
        }
        
        // Process button
        if (processBtn) {
            processBtn.addEventListener('click', function() {
                if (selectedFiles.length === 0) {
                    showError('Please select at least one file');
                    return;
                }
                processLegacyFiles(selectedFiles);
            });
        }
    }
    
    // Process multiple legacy files
    // Process multiple legacy files
    async function processLegacyFiles(files) {
        const loadingDiv = document.getElementById('loading');
        const errorDiv = document.getElementById('error');
        const successDiv = document.getElementById('success');
        const processBtn = document.getElementById('legacyProcessBtn');
        
        const formData = new FormData();
        files.forEach(file => {
            formData.append('userFiles', file);
        });
        
        // Add override values 
        const overrides = {
            hoiDongThi: document.getElementById('legacyHoiDongThi')?.value || '',
            tenCoQuanCapBang: document.getElementById('legacyTenCoQuan')?.value || '',
            chucDanhNguoiKy: document.getElementById('legacyChucDanh')?.value || '',
            nguoiKyBang: document.getElementById('legacyNguoiKy')?.value || '',
            diaDanh: document.getElementById('legacyDiaDanh')?.value || '',
            trangThaiSoHoa: document.getElementById('legacyTrangThai')?.value || ''
        };
        
        formData.append('overrides', JSON.stringify(overrides));
        
        loadingDiv.style.display = 'block';
        errorDiv.style.display = 'none';
        successDiv.style.display = 'none';
        processBtn.disabled = true;
        
        try {
            const response = await fetch('/api/legacy/process', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || 'Processing failed');
            }
            
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            
            // ============ GET FILENAME FROM SERVER ============
            // Try to get filename from Content-Disposition header
            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = 'legacy_certificates_merged.xlsx'; // fallback
            
            if (contentDisposition) {
                // Try to match filename="..." or filename=...
                const match = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                if (match && match[1]) {
                    // Remove quotes if present
                    filename = match[1].replace(/['"]/g, '');
                    // Decode URI component if encoded
                    try {
                        filename = decodeURIComponent(filename);
                    } catch (e) {
                        // If decoding fails, use as-is
                    }
                }
            }
            
            console.log('📥 Downloading file:', filename);
            
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
            showSuccess(`${files.length} file(s) processed successfully!`);
            
        } catch (error) {
            showError(error.message);
        } finally {
            loadingDiv.style.display = 'none';
            processBtn.disabled = false;
        }
    }
    
    // Helper functions
    function showError(message) {
        const errorDiv = document.getElementById('error');
        errorDiv.textContent = message;
        errorDiv.style.display = 'block';
        setTimeout(() => {
            errorDiv.style.display = 'none';
        }, 5000);
    }
    
    function showSuccess(message) {
        const successDiv = document.getElementById('success');
        successDiv.textContent = message;
        successDiv.style.display = 'block';
        setTimeout(() => {
            successDiv.style.display = 'none';
        }, 3000);
    }
});