const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

/**
 * Process multiple user files and merge into one Excel
 */
async function mergeMultipleFiles(req, res) {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }
        
        // Load the template
        const templatePath = path.join(__dirname, '../../template/template.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        
        // Get the Data sheet
        const worksheet = workbook.getWorksheet('Data');
        if (!worksheet) {
            throw new Error('Could not find "Data" worksheet in template');
        }
        
        // Find the data start row in template
        let dataStartRow = -1;
        for (let rowNum = 1; rowNum <= 50; rowNum++) {
            const row = worksheet.getRow(rowNum);
            const firstCell = row.getCell(1).value;
            if (firstCell === 'STT' || (firstCell && String(firstCell).toUpperCase() === 'STT')) {
                dataStartRow = rowNum + 1;
                break;
            }
        }
        
        if (dataStartRow === -1) {
            dataStartRow = 4;
        }
        
        // Clear existing data
        let currentRow = dataStartRow;
        while (currentRow <= worksheet.rowCount) {
            const row = worksheet.getRow(currentRow);
            const firstCell = row.getCell(1).value;
            if (!firstCell || String(firstCell).trim() === '') {
                break;
            }
            currentRow++;
        }
        
        if (currentRow > dataStartRow) {
            worksheet.spliceRows(dataStartRow, currentRow - dataStartRow);
        }
        
        // Track all records and their sources
        let allRecords = [];
        let fileSources = [];
        let currentSTT = 1;
        
        // Process each uploaded file
        for (let fileIndex = 0; fileIndex < req.files.length; fileIndex++) {
            const uploadedFile = req.files[fileIndex];
            const filePath = uploadedFile.path;
            const fileName = uploadedFile.originalname;
            
            console.log(`Processing file ${fileIndex + 1}/${req.files.length}: ${fileName}`);
            
            // Extract data from this file
            const fileData = await extractFileData(filePath);
            
            if (fileData.records.length === 0) {
                console.log(`No records found in ${fileName}, skipping`);
                continue;
            }
            
            // Add source info to each record
            fileData.records.forEach(record => {
                allRecords.push({
                    ...record,
                    STT: currentSTT++,
                    'Nguồn file': fileName,
                    'Tên chứng chỉ (file)': fileData.certificateName,
                    'Ngày cấp (file)': fileData.issueDate
                });
            });
            
            fileSources.push({
                fileName: fileName,
                certificateName: fileData.certificateName,
                issueDate: fileData.issueDate,
                recordCount: fileData.records.length,
                startRow: currentSTT - fileData.records.length,
                endRow: currentSTT - 1
            });
        }
        
        // Insert all records into worksheet
        for (let i = 0; i < allRecords.length; i++) {
            const record = allRecords[i];
            const rowNum = dataStartRow + i;
            const row = worksheet.getRow(rowNum);
            
            // Map fields to template columns
            row.getCell(1).value = record.STT;
            row.getCell(2).value = record.soHieuChungChi || '';
            row.getCell(3).value = record.soVaoSo || '';
            row.getCell(4).value = record['Tên chứng chỉ (file)'] || '';
            row.getCell(5).value = record.hoTen || '';
            row.getCell(9).value = record.ngaySinh || '';
            row.getCell(10).value = record.gioiTinh || '';
            row.getCell(11).value = record.danToc || '';
            row.getCell(13).value = record.noiSinh || '';
            row.getCell(16).value = record.ketQuaThi || '';
            row.getCell(17).value = record.hoiDongThi || '';
            row.getCell(18).value = record.diaDanh || '';
            row.getCell(19).value = record['Ngày cấp (file)'] || '';
            row.getCell(20).value = record.tenCoQuanCapBang || '';
            row.getCell(21).value = record.chucDanhNguoiKy || '';
            row.getCell(22).value = record.nguoiKyBang || '';
            row.getCell(25).value = 'Đã số hóa đầy đủ, chính xác';
            
            row.height = 30;
        }
        
        // Add a summary sheet
        const summarySheet = workbook.addWorksheet('File_Summary');
        summarySheet.getCell('A1').value = 'File Name';
        summarySheet.getCell('B1').value = 'Certificate Name';
        summarySheet.getCell('C1').value = 'Issue Date';
        summarySheet.getCell('D1').value = 'Record Count';
        summarySheet.getCell('E1').value = 'Start Row';
        summarySheet.getCell('F1').value = 'End Row';
        
        // Style header row
        summarySheet.getRow(1).font = { bold: true };
        summarySheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        
        // Add data rows
        for (let i = 0; i < fileSources.length; i++) {
            const source = fileSources[i];
            const rowNum = i + 2;
            summarySheet.getCell(`A${rowNum}`).value = source.fileName;
            summarySheet.getCell(`B${rowNum}`).value = source.certificateName;
            summarySheet.getCell(`C${rowNum}`).value = source.issueDate;
            summarySheet.getCell(`D${rowNum}`).value = source.recordCount;
            summarySheet.getCell(`E${rowNum}`).value = source.startRow;
            summarySheet.getCell(`F${rowNum}`).value = source.endRow;
        }
        
        // Auto-size columns in summary sheet
        summarySheet.columns.forEach(column => {
            column.width = 25;
        });
        
        // Save the workbook
        const tempFilePath = path.join(__dirname, '../../uploads', `merged_multiple_${Date.now()}.xlsx`);
        await workbook.xlsx.writeFile(tempFilePath);
        
        // Clean up uploaded files
        req.files.forEach(file => {
            fs.unlink(file.path, () => {});
        });
        
        // Send file
        res.download(tempFilePath, 'merged_certificates.xlsx', (err) => {
            if (err) console.error('Error sending file:', err);
            setTimeout(() => {
                if (fs.existsSync(tempFilePath)) fs.unlink(tempFilePath, () => {});
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
    const userData = extractUserData(filePath);
    
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