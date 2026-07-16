const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

/**
 * Process multiple user files and merge into one Excel
 */
const { mapToTemplate } = require('../services/templateMapper');

/**
 * Process multiple user files and merge into one Excel
 * Uses the modern template mapping
 */
async function mergeMultipleFiles(req, res) {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }
        
        const { extractUserData } = require('../services/userDataParser');
        const fs = require('fs');
        const path = require('path');
        
        // Collect all records from all files
        let allRecords = [];
        let currentSTT = 1;
        
        console.log(`Processing ${req.files.length} file(s)...`);
        
        // Process each uploaded file
        for (let fileIndex = 0; fileIndex < req.files.length; fileIndex++) {
            const uploadedFile = req.files[fileIndex];
            const filePath = uploadedFile.path;
            const fileName = uploadedFile.originalname;
            
            console.log(`Processing file ${fileIndex + 1}/${req.files.length}: ${fileName}`);
            
            // Extract data from this file using the modern parser
            const userData = extractUserData(filePath);
            
            if (userData.length === 0) {
                console.log(`No records found in ${fileName}, skipping`);
                continue;
            }
            
            // Add STT and source info
            userData.forEach(record => {
                // Preserve all fields from the parser
                const newRecord = {
                    ...record,
                    STT: currentSTT++,
                    'Nguồn file': fileName
                };
                allRecords.push(newRecord);
            });
        }
        
        if (allRecords.length === 0) {
            return res.status(400).json({ error: 'No valid records found in uploaded files' });
        }
        
        console.log(`Total records collected: ${allRecords.length}`);
        
        // Load the template
        const templatePath = path.join(__dirname, '../../template/template.xlsx');
        const outputPath = path.join(__dirname, '../../uploads', `merged_${Date.now()}.xlsx`);
        
        // Use the modern mapper to map ALL records to template
        await mapToTemplate(allRecords, templatePath, outputPath);
        
        // Clean up uploaded files
        req.files.forEach(file => {
            fs.unlink(file.path, () => {});
        });
        
        // Send file
        res.download(outputPath, 'merged_certificates.xlsx', (err) => {
            if (err) console.error('Error sending file:', err);
            setTimeout(() => {
                if (fs.existsSync(outputPath)) fs.unlink(outputPath, () => {});
            }, 5000);
        });
        
    } catch (error) {
        console.error('Error merging multiple files:', error);
        res.status(500).json({ error: 'Error merging files: ' + error.message });
    }
}

/**
 * Extract data from a single user file
 */
async function extractFileData(filePath) {
    const { extractUserData } = require('../services/userDataParser');
    const userData = extractUserData(filePath);``
    
    // Extract certificate name and issue date from the file
    // These are stored in the userData records if we add them
    let certificateName = '';
    let issueDate = '';
    
    if (userData.length > 0) {
        certificateName = userData[0].tenChungChi || '';
        issueDate = userData[0].ngayCapBang || '';
    }
    
    return {
        records: userData,
        certificateName: certificateName,
        issueDate: issueDate
    };
}

module.exports = {
    mergeMultipleFiles
};