const fs = require('fs');
const path = require('path');
const { extractLegacyUserData } = require('../services/userDataParserLegacy');
const { mapToTemplate } = require('../services/templateMapper');

/**
 * Preview legacy file - return record count
 */
async function previewLegacyFile(req, res) {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        const filePath = req.file.path;
        const result = extractLegacyUserData(filePath);
        
        // Clean up uploaded file
        fs.unlink(filePath, () => {});
        
        res.json({
            success: true,
            recordCount: result.records.length,
            certificateName: result.certificateName,
            issueDate: result.issueDate
        });
        
    } catch (error) {
        console.error('Legacy preview error:', error);
        res.status(500).json({ error: error.message });
    }
}

/**
 * Process multiple legacy files with override values
 */
async function processLegacyFiles(req, res) {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }
        
        // Parse overrides from request
        let overrides = {};
        try {
            overrides = req.body.overrides ? JSON.parse(req.body.overrides) : {};
        } catch (e) {
            overrides = {};
        }
        
        // Collect all records from all files
        let allRecords = [];
        let fileCount = 0;
        let totalRecords = 0;
        let firstFileName = '';
        let certificateName = '';
        let issueDate = '';
        
        console.log(`Processing ${req.files.length} legacy file(s)...`);
        
        for (let i = 0; i < req.files.length; i++) {
            const file = req.files[i];
            const filePath = file.path;
            const fileName = file.originalname;
            
            // Store first file name for export naming
            if (!firstFileName) {
                firstFileName = path.parse(fileName).name;
                console.log(`📄 First file name: ${firstFileName}`);
            }
            
            console.log(`Processing file ${i + 1}/${req.files.length}: ${fileName}`);
            
            // Extract data from legacy file
            const result = extractLegacyUserData(filePath);
            
            if (result.records.length === 0) {
                console.log(`No records found in ${fileName}, skipping`);
                continue;
            }
            
            // Use certificate name from first file with data (only as fallback)
            if (!certificateName && result.certificateName) {
                certificateName = result.certificateName;
            }
            
            // Use issue date from first file with data
            if (!issueDate && result.issueDate) {
                issueDate = result.issueDate;
            }
            
            // Add records to collection
            allRecords.push(...result.records);
            fileCount++;
            totalRecords += result.records.length;
        }
        
        // Clean up uploaded files
        req.files.forEach(file => {
            fs.unlink(file.path, () => {});
        });
        
        if (allRecords.length === 0) {
            return res.status(400).json({ error: 'No valid records found in uploaded files' });
        }
        
        console.log(`Total records collected: ${allRecords.length} from ${fileCount} files`);
        
        // Apply overrides to all records
        const processedRecords = allRecords.map((record, index) => ({
            ...record,
            tenChungChi: record.tenChungChi || certificateName || '',
            hoiDongThi: overrides.hoiDongThi || '',
            diaDanh: overrides.diaDanh || '',
            tenCoQuanCapBang: overrides.tenCoQuanCapBang || '',
            chucDanhNguoiKy: overrides.chucDanhNguoiKy || '',
            nguoiKyBang: overrides.nguoiKyBang || '',
            trangThaiSoHoa: overrides.trangThaiSoHoa || 'Đã số hóa đầy đủ, chính xác',
            STT: index + 1
        }));
        
        // Template path
        const templatePath = path.join(__dirname, '../../template/template.xlsx');
        const outputPath = path.join(__dirname, '../../uploads', `legacy_merged_${Date.now()}.xlsx`);
        
        // Map to template
        await mapToTemplate(processedRecords, templatePath, outputPath);
        
        // ============ GENERATE DYNAMIC FILENAME ============
        // Format: Import_{original_filename}_{record_count}CC
        const baseName = firstFileName || 'legacy_certificates';
        // Clean the base name (remove special characters for safe filename)
        const cleanBaseName = baseName.replace(/[^a-zA-Z0-9]/g, '_');
        const exportFileName = `Import_${cleanBaseName}_${allRecords.length}CC.xlsx`;
        
        console.log(`📤 Exporting file: ${exportFileName}`);
        console.log(`📁 Output path: ${outputPath}`);
        
        // Send file with explicit headers
        res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(exportFileName)}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.sendFile(outputPath, (err) => {
            if (err) {
                console.error('Error sending file:', err);
                // If sendFile fails, try download as fallback
                res.download(outputPath, exportFileName, (downloadErr) => {
                    if (downloadErr) console.error('Download fallback error:', downloadErr);
                    setTimeout(() => {
                        if (fs.existsSync(outputPath)) fs.unlink(outputPath, () => {});
                    }, 5000);
                });
            } else {
                setTimeout(() => {
                    if (fs.existsSync(outputPath)) fs.unlink(outputPath, () => {});
                }, 5000);
            }
        });
        
    } catch (error) {
        console.error('Legacy process error:', error);
        res.status(500).json({ error: error.message });
    }
}

module.exports = {
    previewLegacyFile,
    processLegacyFiles
};