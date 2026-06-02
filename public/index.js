// DOM Elements
const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const previewSection = document.getElementById('previewSection');
const previewContent = document.getElementById('previewContent');
const processBtn = document.getElementById('processBtn');
const fileNameSpan = document.getElementById('fileName');
const recordCountSpan = document.getElementById('recordCount');
const loadingDiv = document.getElementById('loading');
const errorDiv = document.getElementById('error');
const successDiv = document.getElementById('success');

let currentFile = null;

// Upload area click handler
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

// File input change handler
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        currentFile = file;
        fileNameSpan.textContent = file.name;
        previewFile(file);
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
    
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
        currentFile = file;
        fileNameSpan.textContent = file.name;
        previewFile(file);
    } else if (file) {
        showError('Please upload an Excel file (.xlsx or .xls)');
    }
});

// Preview file function
async function previewFile(file) {
    loadingDiv.style.display = 'block';
    errorDiv.style.display = 'none';
    previewSection.style.display = 'none';
    
    const formData = new FormData();
    formData.append('userFile', file);
    
    try {
        const response = await fetch('/api/preview', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (!response.ok) {
            throw new Error(data.error || 'Preview failed');
        }
        
        recordCountSpan.textContent = data.recordCount;
        displayPreview(data.preview);
        previewSection.style.display = 'block';
        showSuccess(`Loaded ${data.recordCount} records successfully`);
        
    } catch (error) {
        showError(error.message);
    } finally {
        loadingDiv.style.display = 'none';
    }
}

// Display preview table
function displayPreview(previewData) {
    if (!previewData || previewData.length === 0) {
        previewContent.innerHTML = '<p>No data to preview</p>';
        return;
    }
    
    const headers = Object.keys(previewData[0]);
    
    const table = document.createElement('table');
    table.className = 'preview-table';
    
    // Create header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headers.slice(0, 8).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body
    const tbody = document.createElement('tbody');
    previewData.slice(0, 5).forEach(row => {
        const tr = document.createElement('tr');
        headers.slice(0, 8).forEach(header => {
            const td = document.createElement('td');
            let value = row[header];
            if (value && typeof value === 'string' && value.length > 30) {
                value = value.substring(0, 30) + '...';
            }
            td.textContent = value || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    
    previewContent.innerHTML = '';
    previewContent.appendChild(table);
}

// Process button handler
processBtn.addEventListener('click', async () => {
    if (!currentFile) {
        showError('Please select a file first');
        return;
    }
    
    loadingDiv.style.display = 'block';
    errorDiv.style.display = 'none';
    successDiv.style.display = 'none';
    processBtn.disabled = true;
    
    const formData = new FormData();
    formData.append('userFile', currentFile);
    
    try {
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Processing failed');
        }
        
        // Download the file
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Danh_sach_chung_chi_so.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showSuccess('File processed and downloaded successfully!');
        
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