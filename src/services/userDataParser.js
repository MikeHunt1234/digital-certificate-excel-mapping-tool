const XLSX = require('xlsx');

/**
 * Extract data from user's uploaded Excel file
 */
function extractUserData(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log('=== RAW DATA DEBUG ===');
    console.log(`Total rows: ${rawData.length}`);
    
    // Print first 10 rows for debugging
    for (let i = 0; i < Math.min(10, rawData.length); i++) {
        const row = rawData[i];
        if (row) {
            console.log(`Row ${i}:`, JSON.stringify(row.slice(0, 5)));
        }
    }
    
    // Find header row (contains "TT", "Họ và tên người học")
    let headerRowIndex = -1;
    let dataStartRow = -1;
    
    for (let i = 0; i < Math.min(rawData.length, 30); i++) {
        const row = rawData[i];
        if (!row) continue;
        
        if (row[0] === 'TT' && row[1] === 'Họ và tên người học') {
            headerRowIndex = i;
            dataStartRow = i + 1;
            console.log(`Found header at row ${i}, data starts at row ${dataStartRow}`);
            break;
        }
    }
    
    if (dataStartRow === -1) {
        throw new Error('Could not find data start row in user file');
    }
    
    // Extract certificate name from row 3 (index 3)
    // The merged cell B4:L4 means the value is at column A (index 0)
    const titleRow = rawData[3] || [];
    let b4Value = '';
    
    // Try to get from column A first (index 0)
    if (titleRow[0]) {
        b4Value = titleRow[0];
    } else if (titleRow[1]) {
        b4Value = titleRow[1];
    }
    
    let certificateName = extractCertificateName(b4Value);
    console.log(`Certificate name from row 3: "${b4Value}" → "${certificateName}"`);
    
    // Extract signature info from specific cells
    const signatureRow = rawData[4] || [];
    const signatoryNameRow = rawData[5] || [];
    const dateRow = rawData[6] || [];
    
    const chucDanhNguoiKy = signatureRow[10] || '';
    const nguoiKyBang = signatoryNameRow[10] || '';
    
    // Extract date - could be Excel serial number or text
    let ngayCapBang = '';
    const dateCell = dateRow[10];
    
    if (dateCell) {
        if (typeof dateCell === 'number') {
            // Excel serial number to date conversion
            const excelDate = new Date(Math.round((dateCell - 25569) * 86400 * 1000));
            ngayCapBang = `${excelDate.getDate().toString().padStart(2, '0')}/${(excelDate.getMonth() + 1).toString().padStart(2, '0')}/${excelDate.getFullYear()}`;
            console.log(`Converted Excel serial date: ${dateCell} → ${ngayCapBang}`);
        } else {
            const dateStr = String(dateCell);
            const dateMatch = dateStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
            if (dateMatch) {
                ngayCapBang = dateMatch[0];
            } else {
                const vnDateMatch = dateStr.match(/ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/i);
                if (vnDateMatch) {
                    ngayCapBang = `${vnDateMatch[1].padStart(2, '0')}/${vnDateMatch[2].padStart(2, '0')}/${vnDateMatch[3]}`;
                }
            }
        }
    }
    
    console.log(`Signature - Title: "${chucDanhNguoiKy}", Name: "${nguoiKyBang}", Date: "${ngayCapBang}"`);
    
    // Extract data rows
    const userData = [];
    
    for (let i = dataStartRow; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        const ttValue = row[0] ? String(row[0]).trim() : '';
        
        // Stop if we hit empty row or signature section
        if (ttValue === '') {
            if (row[1] && String(row[1]).includes('HIỆU TRƯỞNG')) break;
            continue;
        }
        
        // Convert date from DD.M.YYYY to DD/MM/YYYY
        let ngaySinhValue = '';
        const rawDate = row[2];
        if (rawDate) {
            let dateStr = String(rawDate).trim();
            ngaySinhValue = dateStr.replace(/\./g, '/');
            const parts = ngaySinhValue.split('/');
            if (parts.length === 3) {
                parts[0] = parts[0].padStart(2, '0');
                parts[1] = parts[1].padStart(2, '0');
                ngaySinhValue = parts.join('/');
            }
        }
        
        const userRecord = {
            STT: ttValue,
            soHieuChungChi: row[7] ? String(row[7]).trim() : '',
            soVaoSo: row[9] ? String(row[9]).trim() : '',
            ngaySinh: ngaySinhValue,
            gioiTinh: row[4] ? String(row[4]).trim() : '',
            danToc: row[5] ? String(row[5]).trim() : '',
            noiSinh: row[3] ? String(row[3]).trim() : '',
            ketQuaThi: row[6] ? String(row[6]).trim() : '',
            hoTen: row[1] ? String(row[1]).trim() : '',
            tenChungChi: certificateName,
            hoiDongThi: 'Đại học Cần Thơ',
            diaDanh: 'thành phố Cần Thơ',
            ngayCapBang: ngayCapBang,
            tenCoQuanCapBang: 'Đại học Cần Thơ',
            chucDanhNguoiKy: chucDanhNguoiKy,
            nguoiKyBang: nguoiKyBang,
            trangThaiSoHoa: 'Đã số hóa đầy đủ, chính xác'
        };
        
        userData.push(userRecord);
    }
    
    console.log(`Extracted ${userData.length} records`);
    if (userData.length > 0) {
        console.log('First record:', JSON.stringify(userData[0], null, 2));
    }
    
    return userData;
}

/**
 * Extract certificate name from B4 cell
 */
function extractCertificateName(b4Value) {
    if (!b4Value) return '';
    let name = String(b4Value);
    // Remove "SỔ GỐC CẤP" prefix
    name = name.replace(/SỔ\s*GỐC\s*CẤP\s*/i, '');
    // Remove any extra spaces
    name = name.trim();
    // Convert to proper case: first letter uppercase, rest lowercase
    name = name.toLowerCase();
    name = name.charAt(0).toUpperCase() + name.slice(1);
    return name;
}

module.exports = { extractUserData };