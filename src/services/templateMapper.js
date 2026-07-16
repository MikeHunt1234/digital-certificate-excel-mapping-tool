const ExcelJS = require('exceljs');
const path = require('path');

async function mapToTemplate(userData, templatePath, outputPath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const worksheet = workbook.getWorksheet('Data');
    if (!worksheet) {
        throw new Error('Could not find "Data" worksheet in template');
    }
    
    let dataStartRow = -1;
    
    for (let rowNum = 1; rowNum <= 50; rowNum++) {
        const row = worksheet.getRow(rowNum);
        for (let col = 1; col <= 10; col++) {
            const cell = row.getCell(col);
            const value = cell.value ? String(cell.value).toUpperCase().trim() : '';
            if (value === 'STT') {
                const nextRow = worksheet.getRow(rowNum + 1);
                const nextCell = nextRow.getCell(1).value;
                if (nextCell && String(nextCell).trim() === 'Mẫu') {
                    dataStartRow = rowNum + 2;
                } else {
                    dataStartRow = rowNum + 1;
                }
                break;
            }
        }
        if (dataStartRow !== -1) break;
    }
    
    if (dataStartRow === -1) {
        dataStartRow = 5;
    }
    
    const fullColumnMappings = {
        1: 'STT',
        2: 'soHieuChungChi',
        3: 'soVaoSo',
        4: 'tenChungChi',
        5: 'hoTen',
        9: 'ngaySinh',
        10: 'gioiTinh',
        11: 'danToc',
        13: 'noiSinh',
        19: 'ketQuaThi',
        20: 'hoiDongThi',
        21: 'diaDanh',
        22: 'ngayCapBang',
        23: 'tenCoQuanCapBang',
        24: 'chucDanhNguoiKy',
        25: 'nguoiKyBang',
        28: 'trangThaiSoHoa'
    };
    
    const validGenders = ['Nam', 'Nữ', 'Khác'];
    
    // Clear existing data rows
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
    
    // Insert new data rows
    for (let i = 0; i < userData.length; i++) {
        const record = userData[i];
        const rowNum = dataStartRow + i;
        const row = worksheet.getRow(rowNum);
        
        for (const [colIndex, fieldName] of Object.entries(fullColumnMappings)) {
            let value = record[fieldName] || '';
            const cell = row.getCell(parseInt(colIndex));
            const colIdx = parseInt(colIndex);
            
            if (colIdx === 10) { // Column J - Giới tính
                if (value && validGenders.includes(value)) {
                    cell.value = value;
                } else {
                    cell.value = '';
                }
                
                cell.dataValidation = {
                    type: 'list',
                    formulae: ['"Nam,Nữ,Khác"'],
                    allowBlank: true,
                    showErrorMessage: true,
                    errorTitle: 'Invalid Gender',
                    error: 'Please select: Nam, Nữ, or Khác'
                };
            } else if (colIdx === 19) { // Column S - Kết quả thi các môn
                // Format as number with 1 decimal place
                if (value !== '' && !isNaN(parseFloat(value))) {
                    cell.value = parseFloat(value);
                    cell.numFmt = '0.0'; // Always show 1 decimal place
                } else {
                    cell.value = '';
                }
            } else {
                cell.value = value;
            }
        }
        
        row.height = 25;
    }
    
    // Set column widths
    worksheet.getColumn(1).width = 8;
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 25;
    worksheet.getColumn(4).width = 30;
    worksheet.getColumn(5).width = 25;
    worksheet.getColumn(6).width = 15;
    worksheet.getColumn(7).width = 20;
    worksheet.getColumn(8).width = 15;
    worksheet.getColumn(9).width = 15;
    worksheet.getColumn(10).width = 10;
    worksheet.getColumn(11).width = 12;
    worksheet.getColumn(12).width = 15;
    worksheet.getColumn(13).width = 25;
    worksheet.getColumn(14).width = 25;
    worksheet.getColumn(15).width = 15;
    worksheet.getColumn(16).width = 25;
    worksheet.getColumn(17).width = 15;
    worksheet.getColumn(18).width = 15;
    worksheet.getColumn(19).width = 15;
    worksheet.getColumn(20).width = 25;
    worksheet.getColumn(21).width = 25;
    worksheet.getColumn(22).width = 15;
    worksheet.getColumn(23).width = 25;
    worksheet.getColumn(24).width = 20;
    worksheet.getColumn(25).width = 25;
    worksheet.getColumn(26).width = 15;
    worksheet.getColumn(27).width = 20;
    worksheet.getColumn(28).width = 25;
    worksheet.getColumn(29).width = 30;
    
    await workbook.xlsx.writeFile(outputPath);
    
    return userData.length;
}

module.exports = { mapToTemplate };