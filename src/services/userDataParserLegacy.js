const XLSX = require('xlsx');

/**
 * Parse legacy 2015 Excel format
 * Expected columns: IDVANBANG, SOHIEU, SOVAOSO, HOTEN, GIOITINH, 
 * NGAYSINH, NOISINH, NGAYKIEMTRA, NAMKIEMTRA, HOIDONG, DIEM, 
 * XEPLOAI, NAMTOTNGHIEP, NGAYKY, TENVANBANG, TENDONVI
 */
function extractLegacyUserData(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Find header row
    let headerRowIndex = -1;
    let dataStartRow = -1;
    
    for (let i = 0; i < Math.min(rawData.length, 30); i++) {
        const row = rawData[i];
        if (!row) continue;
        const rowStr = row.join(' ');
        if (rowStr.includes('SOHIEU') || rowStr.includes('HOTEN') || rowStr.includes('NGAYSINH')) {
            headerRowIndex = i;
            dataStartRow = i + 1;
            break;
        }
    }
    
    if (headerRowIndex === -1) {
        throw new Error('Could not find header row in legacy file');
    }
    
    const headers = rawData[headerRowIndex];
    
    const colMap = {
        idVanBang: headers.findIndex(h => h === 'IDVANBANG'),
        soHieu: headers.findIndex(h => h === 'SOHIEU'),
        soVaoSo: headers.findIndex(h => h === 'SOVAOSO'),
        hoTen: headers.findIndex(h => h === 'HOTEN'),
        gioiTinh: headers.findIndex(h => h === 'GIOITINH'),
        ngaySinh: headers.findIndex(h => h === 'NGAYSINH'),
        noiSinh: headers.findIndex(h => h === 'NOISINH'),
        ngayKiemTra: headers.findIndex(h => h === 'NGAYKIEMTRA'),
        namKiemTra: headers.findIndex(h => h === 'NAMKIEMTRA'),
        hoiDong: headers.findIndex(h => h === 'HOIDONG'),
        diem: headers.findIndex(h => h === 'DIEM'),
        xepLoai: headers.findIndex(h => h === 'XEPLOAI'),
        namTotNghiep: headers.findIndex(h => h === 'NAMTOTNGHIEP'),
        ngayKy: headers.findIndex(h => h === 'NGAYKY'),
        tenVanBang: headers.findIndex(h => h === 'TENVANBANG'),
        tenDonVi: headers.findIndex(h => h === 'TENDONVI')
    };
    
    let fallbackCertificateName = '';
    let issueDate = '';
    const records = [];
    
    for (let i = dataStartRow; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        const soHieu = row[colMap.soHieu] ? String(row[colMap.soHieu]).trim() : '';
        const hoTen = row[colMap.hoTen] ? String(row[colMap.hoTen]).trim() : '';
        
        if (!soHieu && !hoTen) continue;
        
        // ============ GET TENVANBANG ============
        let tenChungChi = '';
        if (colMap.tenVanBang !== -1 && row[colMap.tenVanBang]) {
            tenChungChi = String(row[colMap.tenVanBang]).trim();
        }
        
        if (!fallbackCertificateName && tenChungChi) {
            fallbackCertificateName = tenChungChi;
        }
        
        // ============ CONVERT NGAYKY ============
        let ngayCapBang = '';
        if (colMap.ngayKy !== -1 && row[colMap.ngayKy]) {
            const rawDate = row[colMap.ngayKy];
            if (typeof rawDate === 'number') {
                const excelDate = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
                ngayCapBang = `${excelDate.getDate().toString().padStart(2, '0')}/${(excelDate.getMonth() + 1).toString().padStart(2, '0')}/${excelDate.getFullYear()}`;
            } else {
                ngayCapBang = convertOracleTimestamp(String(rawDate));
            }
            if (!issueDate && ngayCapBang) {
                issueDate = ngayCapBang;
            }
        }
        
        // ============ FORMAT NGAYSINH ============
        let ngaySinh = '';
        if (colMap.ngaySinh !== -1 && row[colMap.ngaySinh]) {
            const rawDate = row[colMap.ngaySinh];
            
            if (typeof rawDate === 'number') {
                // It's a number - could be a year or Excel serial
                if (rawDate >= 1900 && rawDate <= 2100) {
                    // Treat as year: 1991 → 01/01/1991
                    ngaySinh = `01/01/${rawDate}`;
                } else {
                    // Try as Excel serial date
                    try {
                        const excelDate = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
                        const year = excelDate.getFullYear();
                        if (year >= 1900 && year <= 2100) {
                            ngaySinh = `${excelDate.getDate().toString().padStart(2, '0')}/${(excelDate.getMonth() + 1).toString().padStart(2, '0')}/${year}`;
                        } else {
                            ngaySinh = `01/01/${rawDate}`;
                        }
                    } catch (e) {
                        ngaySinh = `01/01/${rawDate}`;
                    }
                }
            } else {
                // It's a string
                const dateStr = String(rawDate).trim();
                
                // Check if it's a year only (4 digits, no slashes or dashes)
                if (/^\d{4}$/.test(dateStr)) {
                    // Year only: 1991 → 01/01/1991
                    ngaySinh = `01/01/${dateStr}`;
                } else {
                    // Try to format as DD/MM/YYYY
                    const dateMatch = dateStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                    if (dateMatch) {
                        ngaySinh = `${dateMatch[1].padStart(2, '0')}/${dateMatch[2].padStart(2, '0')}/${dateMatch[3]}`;
                    } else {
                        const dotMatch = dateStr.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
                        if (dotMatch) {
                            ngaySinh = `${dotMatch[1].padStart(2, '0')}/${dotMatch[2].padStart(2, '0')}/${dotMatch[3]}`;
                        } else {
                            // Try to extract year from string
                            const yearMatch = dateStr.match(/\b(\d{4})\b/);
                            if (yearMatch) {
                                ngaySinh = `01/01/${yearMatch[1]}`;
                            } else {
                                ngaySinh = dateStr;
                            }
                        }
                    }
                }
            }
        }
        
        records.push({
            soHieuChungChi: soHieu,
            soVaoSo: row[colMap.soVaoSo] ? String(row[colMap.soVaoSo]).trim() : '',
            hoTen: hoTen,
            ngaySinh: ngaySinh,
            gioiTinh: colMap.gioiTinh !== -1 && row[colMap.gioiTinh] ? String(row[colMap.gioiTinh]).trim() : '',
            noiSinh: colMap.noiSinh !== -1 && row[colMap.noiSinh] ? String(row[colMap.noiSinh]).trim() : '',
            tenChungChi: tenChungChi || fallbackCertificateName || '',
            hoiDongThi: '',
            diaDanh: '',
            ngayCapBang: ngayCapBang,
            tenCoQuanCapBang: '',
            chucDanhNguoiKy: '',
            nguoiKyBang: '',
            trangThaiSoHoa: '',
            ketQuaThi: colMap.diem !== -1 && row[colMap.diem] ? String(row[colMap.diem]).trim() : '',
            danToc: '',
            _original: row
        });
    }
    
    return {
        records: records,
        certificateName: fallbackCertificateName,
        issueDate: issueDate
    };
}

/**
 * Convert Oracle timestamp string to DD/MM/YYYY
 * Handles formats like:
 * - 07-JUN-15 12.00.00.000000000 AM
 * - 28-JUL-15 12.00.00.000000000 AM
 * - 15/01/2025
 * - 2025-01-15
 */
function convertOracleTimestamp(dateStr) {
    if (!dateStr) return '';
    
    const str = String(dateStr).trim();
    
    // Check if already in DD/MM/YYYY format
    let match = str.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) {
        return `${match[1].padStart(2, '0')}/${match[2].padStart(2, '0')}/${match[3]}`;
    }
    
    // Check if in YYYY-MM-DD format
    match = str.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (match) {
        return `${match[3].padStart(2, '0')}/${match[2].padStart(2, '0')}/${match[1]}`;
    }
    
    // Handle Oracle timestamp: 07-JUN-15 12.00.00.000000000 AM
    match = str.match(/(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
    if (match) {
        const months = {
            'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04',
            'MAY': '05', 'JUN': '06', 'JUL': '07', 'AUG': '08',
            'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'
        };
        const month = months[match[2].toUpperCase()] || '01';
        let year = match[3];
        // Handle 2-digit year (15 → 2015)
        if (year.length === 2) {
            year = '20' + year;
        }
        return `${match[1].padStart(2, '0')}/${month}/${year}`;
    }
    
    // Handle DD.MM.YYYY
    match = str.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (match) {
        return `${match[1].padStart(2, '0')}/${match[2].padStart(2, '0')}/${match[3]}`;
    }
    
    // If nothing matches, return original
    return str;
}

module.exports = { extractLegacyUserData };