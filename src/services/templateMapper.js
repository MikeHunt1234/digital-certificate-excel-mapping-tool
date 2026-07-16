const ExcelJS = require('exceljs');
const path = require('path');

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
    
    // ============ UPDATED COLUMN MAPPING FOR 29-COLUMN TEMPLATE ============
    // A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, 
    // N=14, O=15, P=16, Q=17, R=18, S=19, T=20, U=21, V=22, W=23, X=24, 
    // Y=25, Z=26, AA=27, AB=28, AC=29
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
        19: 'ketQuaThi',             // S (moved from P)
        20: 'hoiDongThi',            // T (moved from Q)
        21: 'diaDanh',               // U (moved from R)
        22: 'ngayCapBang',           // V (moved from S)
        23: 'tenCoQuanCapBang',      // W (moved from T)
        24: 'chucDanhNguoiKy',       // X (moved from U)
        25: 'nguoiKyBang',           // Y (moved from V)
        28: 'trangThaiSoHoa'         // AB (moved from Y)
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
    
    // Set column widths for the new 29-column template
    worksheet.getColumn(1).width = 8;   // A - STT
    worksheet.getColumn(2).width = 15;  // B - Số hiệu chứng chỉ
    worksheet.getColumn(3).width = 25;  // C - Số vào sổ gốc
    worksheet.getColumn(4).width = 30;  // D - Tên chứng chỉ
    worksheet.getColumn(5).width = 25;  // E - Họ, chữ đệm và tên
    worksheet.getColumn(6).width = 15;  // F - Mã người học
    worksheet.getColumn(7).width = 20;  // G - Số định danh cá nhân
    worksheet.getColumn(8).width = 15;  // H - Hộ chiếu
    worksheet.getColumn(9).width = 15;  // I - Ngày, tháng, năm sinh
    worksheet.getColumn(10).width = 10; // J - Giới tính
    worksheet.getColumn(11).width = 12; // K - Dân tộc
    worksheet.getColumn(12).width = 15; // L - Quốc tịch
    worksheet.getColumn(13).width = 25; // M - Nơi sinh
    worksheet.getColumn(14).width = 25; // N - Tên trường đã học
    worksheet.getColumn(15).width = 15; // O - Niên khóa
    worksheet.getColumn(16).width = 25; // P - Tên chương trình đào tạo
    worksheet.getColumn(17).width = 15; // Q - Thời gian đào tạo từ
    worksheet.getColumn(18).width = 15; // R - Thời gian đào tạo đến
    worksheet.getColumn(19).width = 15; // S - Kết quả thi các môn
    worksheet.getColumn(20).width = 25; // T - Hội đồng thi
    worksheet.getColumn(21).width = 25; // U - Địa danh nơi cơ quan cấp bằng
    worksheet.getColumn(22).width = 15; // V - Ngày tháng năm cấp bằng
    worksheet.getColumn(23).width = 25; // W - Tên cơ quan cấp bằng
    worksheet.getColumn(24).width = 20; // X - Chức danh người ký bằng
    worksheet.getColumn(25).width = 25; // Y - Họ, chữ đệm, tên người ký bằng
    worksheet.getColumn(26).width = 15; // Z - Tình trạng chứng chỉ
    worksheet.getColumn(27).width = 20; // AA - Ghi chú
    worksheet.getColumn(28).width = 25; // AB - Trạng thái số hóa
    worksheet.getColumn(29).width = 30; // AC - Thông tin chưa bảo đảm...
    
    // Save the workbook
    await workbook.xlsx.writeFile(outputPath);
    
    return userData.length;
}

module.exports = { mapToTemplate };