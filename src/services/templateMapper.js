const ExcelJS = require('exceljs');
const path = require('path');

/**
 * Map user data to template and generate new Excel file
 */
async function mapToTemplate(userData, templatePath, outputPath) {
    // Load template workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    
    // Get the Data sheet
    const worksheet = workbook.getWorksheet('Data');
    if (!worksheet) {
        throw new Error('Could not find "Data" worksheet in template');
    }
    
    console.log('Template loaded, finding data start row...');
    
    // Find the data start row - look for row containing "STT" in any cell
    let dataStartRow = -1;
    
    for (let rowNum = 1; rowNum <= 50; rowNum++) {
        const row = worksheet.getRow(rowNum);
        for (let col = 1; col <= 10; col++) {
            const cell = row.getCell(col);
            const value = cell.value ? String(cell.value).toUpperCase().trim() : '';
            if (value === 'STT') {
                dataStartRow = rowNum + 1;
                console.log(`Found STT at row ${rowNum}, data starts at row ${dataStartRow}`);
                break;
            }
        }
        if (dataStartRow !== -1) break;
    }
    
    if (dataStartRow === -1) {
        // Default to row 4 if STT not found
        dataStartRow = 4;
        console.log('STT not found, defaulting data start row to 4');
    }
    
    // Define column mappings based on your template structure (1-indexed Excel columns)
    // A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, N=14, O=15, P=16, Q=17, R=18, S=19, T=20, U=21, V=22, W=23, X=24, Y=25, Z=26
    const fullColumnMappings = {
        1: 'STT',                    // A
        2: 'soHieuChungChi',         // B
        3: 'soVaoSo',                // C
        4: 'tenChungChi',            // D
        5: 'hoTen',                  // E
        9: 'ngaySinh',               // I
        10: 'gioiTinh',              // J
        11: 'danToc',                // K
        13: 'noiSinh',               // M
        16: 'ketQuaThi',             // P
        17: 'hoiDongThi',            // Q
        18: 'diaDanh',               // R
        19: 'ngayCapBang',           // S
        20: 'tenCoQuanCapBang',      // T
        21: 'chucDanhNguoiKy',       // U
        22: 'nguoiKyBang',           // V
        25: 'trangThaiSoHoa'         // Y
    };
    
    // Clear existing data rows (start from dataStartRow)
    let currentRow = dataStartRow;
    while (currentRow <= worksheet.rowCount) {
        const row = worksheet.getRow(currentRow);
        const firstCell = row.getCell(1).value;
        // Stop if we hit an empty row
        if (!firstCell || String(firstCell).trim() === '') {
            break;
        }
        currentRow++;
    }
    
    // Delete existing data rows
    if (currentRow > dataStartRow) {
        worksheet.spliceRows(dataStartRow, currentRow - dataStartRow);
        console.log(`Cleared ${currentRow - dataStartRow} existing rows`);
    }
    
    // Insert new data rows
    console.log(`Inserting ${userData.length} new rows...`);
    
    for (let i = 0; i < userData.length; i++) {
        const record = userData[i];
        const rowNum = dataStartRow + i;
        const row = worksheet.getRow(rowNum);
        
        // Set each cell based on mapping
        for (const [colIndex, fieldName] of Object.entries(fullColumnMappings)) {
            let value = record[fieldName] || '';
            const cell = row.getCell(parseInt(colIndex));
            cell.value = value;
        }
        
        // Set row height for better visibility
        row.height = 25;
    }
    
    console.log(`Successfully mapped ${userData.length} records`);
    
    // Save the workbook
    await workbook.xlsx.writeFile(outputPath);
    
    return userData.length;
}

module.exports = { mapToTemplate };