const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

/**
 * Normalize a string for matching (trim, case-insensitive)
 */
function normalizeString(str) {
    if (!str) return '';
    return String(str).trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Normalize date for matching - handles multiple formats
 */
function normalizeDate(dateStr) {
    if (!dateStr) return '';
    
    const str = String(dateStr).trim();
    
    // If it's already in dd/MM/yyyy format
    const parts = str.split('/');
    if (parts.length === 3) {
        // Validate it's a proper date
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]);
        const year = parseInt(parts[2]);
        if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1900) {
            return `${day.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/${year}`;
        }
    }
    
    // Try parsing with Date object
    try {
        const date = new Date(str);
        if (!isNaN(date.getTime())) {
            const day = date.getDate().toString().padStart(2, '0');
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        }
    } catch (e) {
        // Ignore
    }
    
    // If it's an Excel serial number
    if (!isNaN(str) && str.length > 5) {
        try {
            const date = new Date((parseFloat(str) - 25569) * 86400 * 1000);
            const day = date.getDate().toString().padStart(2, '0');
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        } catch (e) {
            // Ignore
        }
    }
    
    return str;
}

/**
 * Build a lookup map from the source file (File B - registry)
 */
function buildLookupMap(sourceData, colMap) {
    const lookupMap = new Map();
    let duplicateCount = 0;
    let totalRecords = 0;
    
    sourceData.forEach(record => {
        const fullName = record[colMap.hoTen] ? String(record[colMap.hoTen]).trim() : '';
        const birthdate = record[colMap.ngaySinh] ? String(record[colMap.ngaySinh]).trim() : '';
        
        if (!fullName || !birthdate) {
            return; // Skip records without key data
        }
        
        totalRecords++;
        
        const normalizedName = normalizeString(fullName);
        const normalizedBirth = normalizeDate(birthdate);
        const key = normalizedName + '|' + normalizedBirth;
        
        const gender = record[colMap.gioiTinh] ? String(record[colMap.gioiTinh]).trim() : '';
        const ethnicity = record[colMap.danToc] ? String(record[colMap.danToc]).trim() : '';
        const birthplace = record[colMap.noiSinh] ? String(record[colMap.noiSinh]).trim() : '';
        
        const data = { gender, ethnicity, birthplace };
        
        if (lookupMap.has(key)) {
            const existing = lookupMap.get(key);
            if (Array.isArray(existing)) {
                existing.push(data);
            } else {
                lookupMap.set(key, [existing, data]);
            }
            duplicateCount++;
        } else {
            lookupMap.set(key, data);
        }
    });
    
    console.log(`📊 Source lookup built: ${lookupMap.size} unique keys from ${totalRecords} records (${duplicateCount} duplicates)`);
    return lookupMap;
}

/**
 * Parse source Excel file (registry data)
 */
function parseSourceExcel(filePath) {
    console.log('📂 Parsing source file:', filePath);
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    console.log(`📄 Source file has ${rawData.length} rows`);
    
    // Log first few rows for debugging
    console.log('📋 First 5 rows of source file:');
    for (let i = 0; i < Math.min(rawData.length, 5); i++) {
        console.log(`Row ${i}:`, rawData[i] ? rawData[i].slice(0, 8).join(' | ') : 'empty');
    }
    
    // Find header row containing "Họ tên" and "Ngày sinh"
    let headerRowIndex = -1;
    let dataStartRow = -1;
    
    for (let i = 0; i < Math.min(rawData.length, 20); i++) {
        const row = rawData[i];
        if (!row) continue;
        const rowStr = row.join(' ');
        if (rowStr.includes('Họ tên') && rowStr.includes('Ngày sinh')) {
            headerRowIndex = i;
            dataStartRow = i + 1;
            console.log(`✅ Found source header row at index ${i}`);
            break;
        }
    }
    
    // Try alternate header detection
    if (headerRowIndex === -1) {
        console.log('⚠️ Looking for source header with alternate criteria...');
        for (let i = 0; i < Math.min(rawData.length, 20); i++) {
            const row = rawData[i];
            if (!row) continue;
            const rowStr = row.join(' ');
            if (rowStr.includes('Họ tên') || rowStr.includes('Họ và tên')) {
                headerRowIndex = i;
                dataStartRow = i + 1;
                console.log(`✅ Found source header row at index ${i} (alternate)`);
                break;
            }
        }
    }
    
    if (headerRowIndex === -1) {
        throw new Error('Could not find header row in source file');
    }
    
    const headers = rawData[headerRowIndex];
    console.log('📋 Source headers:', headers);
    
    // Map column indices
    const colMap = {
        hoTen: headers.findIndex(h => h && String(h).includes('Họ tên')),
        ngaySinh: headers.findIndex(h => h && String(h).includes('Ngày sinh')),
        gioiTinh: headers.findIndex(h => h && String(h).includes('Giới tính')),
        danToc: headers.findIndex(h => h && String(h).includes('Dân tộc')),
        noiSinh: headers.findIndex(h => h && String(h).includes('Nơi sinh')),
    };
    
    console.log('📋 Source column map:', colMap);
    
    // Extract data rows
    const records = [];
    for (let i = dataStartRow; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        const hoTen = row[colMap.hoTen] ? String(row[colMap.hoTen]).trim() : '';
        if (!hoTen) continue;
        
        records.push({
            hoTen: hoTen,
            ngaySinh: row[colMap.ngaySinh] ? String(row[colMap.ngaySinh]).trim() : '',
            gioiTinh: row[colMap.gioiTinh] ? String(row[colMap.gioiTinh]).trim() : '',
            danToc: row[colMap.danToc] ? String(row[colMap.danToc]).trim() : '',
            noiSinh: row[colMap.noiSinh] ? String(row[colMap.noiSinh]).trim() : '',
        });
    }
    
    console.log(`✅ Parsed ${records.length} records from source file`);
    return {
        records,
        colMap,
        totalRecords: records.length
    };
}

/**
 * Parse target Excel file (template-mapped)
 */
function parseTargetExcel(filePath) {
    console.log('📂 Parsing target file:', filePath);
    const workbook = XLSX.readFile(filePath);
    
    // Try to find the "Data" sheet first, otherwise use first sheet
    let sheetName = 'Data';
    if (!workbook.SheetNames.includes('Data')) {
        // Try other possible names
        const possibleNames = ['Data', 'Dữ liệu', 'Danh sách', 'Sheet1'];
        for (const name of possibleNames) {
            if (workbook.SheetNames.includes(name)) {
                sheetName = name;
                break;
            }
        }
        // If still not found, use first sheet
        if (sheetName === 'Data') {
            sheetName = workbook.SheetNames[0];
            console.log(`⚠️ "Data" sheet not found, using first sheet: "${sheetName}"`);
        }
    } else {
        console.log(`✅ Found "Data" sheet`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    console.log(`📄 Target file has ${rawData.length} rows in sheet "${sheetName}"`);
    
    // Log first few rows for debugging
    console.log('📋 First 10 rows of target file:');
    for (let i = 0; i < Math.min(rawData.length, 10); i++) {
        const row = rawData[i];
        if (row) {
            const displayRow = row.slice(0, 8).map(cell => String(cell).substring(0, 30)).join(' | ');
            console.log(`Row ${i}: ${displayRow}`);
        } else {
            console.log(`Row ${i}: empty`);
        }
    }
    
    // Find header row containing "STT" and "Họ" or the certificate data headers
    let headerRowIndex = -1;
    let dataStartRow = -1;
    
    for (let i = 0; i < Math.min(rawData.length, 30); i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        const rowStr = row.join(' ');
        
        // Look for the certificate data headers (STT, Số hiệu chứng chỉ, Họ, chữ đệm và tên, etc.)
        if (rowStr.includes('STT') && 
            (rowStr.includes('Số hiệu chứng chỉ') || 
             rowStr.includes('Số vào sổ gốc') || 
             rowStr.includes('Họ, chữ đệm và tên') || 
             rowStr.includes('Họ tên'))) {
            headerRowIndex = i;
            dataStartRow = i + 1;
            console.log(`✅ Found target header row at index ${i}:`, row.slice(0, 10));
            break;
        }
    }
    
    // If not found, try to find row with "STT" and any certificate-related headers
    if (headerRowIndex === -1) {
        console.log('⚠️ Looking for target header with alternate criteria...');
        for (let i = 0; i < Math.min(rawData.length, 30); i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            const rowStr = row.join(' ');
            
            // Look for STT and any of these keywords
            if (rowStr.includes('STT') && 
                (rowStr.includes('chứng chỉ') || 
                 rowStr.includes('giới tính') || 
                 rowStr.includes('dân tộc') || 
                 rowStr.includes('nơi sinh') ||
                 rowStr.includes('ngày sinh'))) {
                headerRowIndex = i;
                dataStartRow = i + 1;
                console.log(`✅ Found target header row at index ${i} (alternate):`, row.slice(0, 10));
                break;
            }
        }
    }
    
    // Last resort: find a row that looks like data (has numbers and text)
    if (headerRowIndex === -1) {
        console.log('⚠️ Using last resort header detection...');
        for (let i = 0; i < Math.min(rawData.length, 30); i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            
            // Check if this row has at least 5 non-empty cells
            const nonEmpty = row.filter(cell => String(cell).trim() !== '');
            if (nonEmpty.length >= 5) {
                // Check if first column is a number (STT)
                const firstCell = row[0] ? String(row[0]).trim() : '';
                if (!isNaN(firstCell) && firstCell !== '') {
                    // This is a data row, use row above as header
                    headerRowIndex = i - 1;
                    dataStartRow = i;
                    console.log(`✅ Found data row at index ${i}, using row ${headerRowIndex} as header`);
                    break;
                }
                // Or if no number, just use this as header
                if (headerRowIndex === -1) {
                    headerRowIndex = i;
                    dataStartRow = i + 1;
                    console.log(`✅ Using row ${i} as header (last resort)`);
                    break;
                }
            }
        }
    }
    
    if (headerRowIndex === -1 || headerRowIndex < 0) {
        throw new Error('Could not find header row in target file. Please ensure the file follows the template format with a "Data" sheet.');
    }
    
    const headers = rawData[headerRowIndex];
    console.log('📋 Target headers found:', headers.slice(0, 15));
    
    // Map column indices with flexible matching
    const colMap = {
        stt: -1,
        sbd: -1,
        hoTen: -1,
        ngaySinh: -1,
        gioiTinh: -1,
        danToc: -1,
        noiSinh: -1,
    };
    
    headers.forEach((header, index) => {
        if (!header) return;
        const headerStr = String(header).trim();
        
        // STT mapping
        if (headerStr === 'STT' || headerStr.includes('STT')) {
            colMap.stt = index;
        }
        // Số hiệu or SBD
        else if (headerStr.includes('SBD') || headerStr.includes('Số hiệu') || headerStr.includes('Số vào sổ')) {
            colMap.sbd = index;
        }
        // Họ tên
        else if (headerStr.includes('Họ, chữ đệm và tên') || headerStr.includes('Họ tên') || headerStr.includes('Họ và tên')) {
            colMap.hoTen = index;
        }
        // Ngày sinh
        else if (headerStr.includes('Ngày, tháng, năm sinh') || headerStr.includes('Ngày sinh')) {
            colMap.ngaySinh = index;
        }
        // Giới tính
        else if (headerStr.includes('Giới tính')) {
            colMap.gioiTinh = index;
        }
        // Dân tộc
        else if (headerStr.includes('Dân tộc')) {
            colMap.danToc = index;
        }
        // Nơi sinh
        else if (headerStr.includes('Nơi sinh')) {
            colMap.noiSinh = index;
        }
    });
    
    console.log('📋 Target column map:', colMap);
    
    // Verify required columns are found
    if (colMap.hoTen === -1) {
        console.warn('⚠️ "Họ, chữ đệm và tên" column not found');
        // Try to find by position if header row is known
        if (headers.length > 4) {
            colMap.hoTen = 4; // Column 5 (0-indexed)
            console.log(`Using column ${colMap.hoTen} as "Họ, chữ đệm và tên"`);
        }
    }
    
    if (colMap.ngaySinh === -1) {
        console.warn('⚠️ "Ngày, tháng, năm sinh" column not found');
        if (headers.length > 8) {
            colMap.ngaySinh = 8; // Column 9 (0-indexed)
            console.log(`Using column ${colMap.ngaySinh} as "Ngày, tháng, năm sinh"`);
        }
    }
    
    // Extract data rows
    const records = [];
    for (let i = dataStartRow; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        // Check if row has data (at least STT or name)
        const stt = row[colMap.stt] !== undefined && row[colMap.stt] !== null ? String(row[colMap.stt]).trim() : '';
        const name = colMap.hoTen !== -1 && row[colMap.hoTen] !== undefined && row[colMap.hoTen] !== null ? String(row[colMap.hoTen]).trim() : '';
        
        if (!stt && !name) continue;
        
        // Skip if STT is not a number (might be header or empty)
        if (stt && isNaN(stt) && stt !== '') continue;
        
        const record = {
            stt: stt,
            sbd: colMap.sbd !== -1 && row[colMap.sbd] !== undefined && row[colMap.sbd] !== null ? String(row[colMap.sbd]).trim() : '',
            hoTen: name,
            ngaySinh: colMap.ngaySinh !== -1 && row[colMap.ngaySinh] !== undefined && row[colMap.ngaySinh] !== null ? String(row[colMap.ngaySinh]).trim() : '',
            gioiTinh: colMap.gioiTinh !== -1 && row[colMap.gioiTinh] !== undefined && row[colMap.gioiTinh] !== null ? String(row[colMap.gioiTinh]).trim() : '',
            danToc: colMap.danToc !== -1 && row[colMap.danToc] !== undefined && row[colMap.danToc] !== null ? String(row[colMap.danToc]).trim() : '',
            noiSinh: colMap.noiSinh !== -1 && row[colMap.noiSinh] !== undefined && row[colMap.noiSinh] !== null ? String(row[colMap.noiSinh]).trim() : '',
            fullRow: row,
            rowIndex: i
        };
        
        records.push(record);
    }
    
    console.log(`✅ Parsed ${records.length} records from target file`);
    
    // Log first few records for debugging
    if (records.length > 0) {
        console.log('📋 First 3 records:');
        for (let i = 0; i < Math.min(records.length, 3); i++) {
            const r = records[i];
            console.log(`  Record ${i}: STT=${r.stt}, Name="${r.hoTen}", Birth="${r.ngaySinh}"`);
        }
    }
    
    return {
        records,
        colMap,
        headerRow: rawData[headerRowIndex],
        headerRowIndex,
        dataStartRow,
        totalRecords: records.length,
        filePath: filePath,
        rawData: rawData,
        sheetName: sheetName
    };
}

/**
 * Enrich target data with source data
 * @param {Array} targetRecords - Target records to enrich
 * @param {Array} sourceRecords - Source records with data
 * @param {Object} targetColMap - Column mapping for target
 * @param {Object} sourceColMap - Column mapping for source
 * @param {Boolean} overwriteExisting - If true, overwrite existing data; if false, only fill empty fields
 */
function enrichData(targetRecords, sourceRecords, targetColMap, sourceColMap, overwriteExisting = false) {
    console.log('🔍 Building lookup map from source data...');
    const lookupMap = buildLookupMap(sourceRecords, sourceColMap);
    
    const usageTracker = new Map();
    let enrichedCount = 0;
    let matchedCount = 0;
    let overwrittenCount = 0;
    
    // Log sample keys for debugging
    let sampleKeys = [];
    let i = 0;
    for (const [key] of lookupMap) {
        if (i < 5) {
            sampleKeys.push(key);
            i++;
        } else break;
    }
    console.log('📋 Sample lookup keys:', sampleKeys);
    console.log(`📋 Overwrite mode: ${overwriteExisting ? 'ON (will overwrite existing data)' : 'OFF (only fill empty fields)'}`);
    
    const enrichedRecords = targetRecords.map((record, index) => {
        const fullName = record.hoTen || '';
        const birthdate = record.ngaySinh || '';
        
        if (!fullName || !birthdate) {
            console.log(`⚠️ Record ${index} missing name or birthdate: "${fullName}" | "${birthdate}"`);
            return record;
        }
        
        const normalizedName = normalizeString(fullName);
        const normalizedBirth = normalizeDate(birthdate);
        const key = normalizedName + '|' + normalizedBirth;
        
        // Log first few records for debugging
        if (index < 3) {
            console.log(`🔑 Record ${index}: "${fullName}" + "${birthdate}" -> key: "${key}"`);
        }
        
        if (!lookupMap.has(key)) {
            if (index < 3) {
                console.log(`❌ No match for key: "${key}"`);
            }
            return record;
        }
        
        matchedCount++;
        let matchData = lookupMap.get(key);
        
        // Handle duplicates
        if (Array.isArray(matchData)) {
            if (!usageTracker.has(key)) {
                usageTracker.set(key, 0);
            }
            const currentIndex = usageTracker.get(key);
            if (currentIndex < matchData.length) {
                matchData = matchData[currentIndex];
                usageTracker.set(key, currentIndex + 1);
            } else {
                console.log(`⚠️ No more duplicates for key: "${key}"`);
                return record;
            }
        }
        
        // Create new record with source data (overwrite or fill)
        const newRecord = { ...record };
        let fieldsChanged = false;
        
        // Check if we should overwrite or just fill empty fields
        const shouldUpdateGender = overwriteExisting || !record.gioiTinh || record.gioiTinh.trim() === '';
        const shouldUpdateEthnicity = overwriteExisting || !record.danToc || record.danToc.trim() === '';
        const shouldUpdateBirthplace = overwriteExisting || !record.noiSinh || record.noiSinh.trim() === '';
        
        // Update gender
        if (shouldUpdateGender && matchData.gender) {
            const oldValue = newRecord.gioiTinh || '(empty)';
            newRecord.gioiTinh = matchData.gender;
            fieldsChanged = true;
            if (index < 3) {
                console.log(`✅ Updated gender: "${oldValue}" -> "${matchData.gender}" for "${fullName}"`);
            }
        }
        
        // Update ethnicity
        if (shouldUpdateEthnicity && matchData.ethnicity) {
            const oldValue = newRecord.danToc || '(empty)';
            newRecord.danToc = matchData.ethnicity;
            fieldsChanged = true;
            if (index < 3) {
                console.log(`✅ Updated ethnicity: "${oldValue}" -> "${matchData.ethnicity}" for "${fullName}"`);
            }
        }
        
        // Update birthplace
        if (shouldUpdateBirthplace && matchData.birthplace) {
            const oldValue = newRecord.noiSinh || '(empty)';
            newRecord.noiSinh = matchData.birthplace;
            fieldsChanged = true;
            if (index < 3) {
                console.log(`✅ Updated birthplace: "${oldValue}" -> "${matchData.birthplace}" for "${fullName}"`);
            }
        }
        
        if (fieldsChanged) {
            enrichedCount++;
            if (overwriteExisting) {
                overwrittenCount++;
            }
        } else {
            if (index < 3) {
                if (matchData.gender || matchData.ethnicity || matchData.birthplace) {
                    console.log(`ℹ️ Record "${fullName}" - no fields needed updating (data already present or no source data)`);
                }
            }
        }
        
        return newRecord;
    });
    
    console.log(`✅ Enrichment complete: ${enrichedCount} records updated out of ${targetRecords.length} total (${matchedCount} matched)`);
    if (overwriteExisting) {
        console.log(`✅ ${overwrittenCount} records had existing data overwritten`);
    }
    
    return {
        enrichedRecords,
        enrichedCount,
        totalRecords: targetRecords.length,
        matchedCount: matchedCount,
        overwrittenCount: overwrittenCount,
        lookupMapSize: lookupMap.size,
        overwriteMode: overwriteExisting
    };
}

/**
 * Rebuild Excel file with enriched data
 */
/**
 * Rebuild Excel file with enriched data - NOW OVERWRITES EXISTING DATA
 */
async function rebuildExcelFile(target, enrichedRecords, outputPath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(target.filePath);
    
    const worksheet = workbook.getWorksheet(target.sheetName || 'Data');
    if (!worksheet) {
        throw new Error(`Could not find "${target.sheetName || 'Data'}" worksheet`);
    }
    
    // Find the data start row
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
        // Try to find first row with a number in column A
        for (let rowNum = 1; rowNum <= 50; rowNum++) {
            const row = worksheet.getRow(rowNum);
            const firstCell = row.getCell(1).value;
            if (firstCell && !isNaN(firstCell) && String(firstCell).trim() !== '') {
                dataStartRow = rowNum;
                break;
            }
        }
    }
    
    if (dataStartRow === -1) {
        console.warn('⚠️ Could not find data start row, using row 5');
        dataStartRow = 5;
    }
    
    console.log(`📝 Data starts at row ${dataStartRow}`);
    
    // Find column indices
    const headerRow = worksheet.getRow(dataStartRow - 1);
    let gioiTinhCol = -1;
    let danTocCol = -1;
    let noiSinhCol = -1;
    
    for (let col = 1; col <= 29; col++) {
        const cell = headerRow.getCell(col);
        const value = cell.value ? String(cell.value).trim() : '';
        if (value.includes('Giới tính')) gioiTinhCol = col;
        if (value.includes('Dân tộc')) danTocCol = col;
        if (value.includes('Nơi sinh')) noiSinhCol = col;
    }
    
    // Fallback positions if not found
    if (gioiTinhCol === -1) gioiTinhCol = 10;
    if (danTocCol === -1) danTocCol = 11;
    if (noiSinhCol === -1) noiSinhCol = 13;
    
    console.log(`📝 Column positions: Giới tính=${gioiTinhCol}, Dân tộc=${danTocCol}, Nơi sinh=${noiSinhCol}`);
    
    let updatedRows = 0;
    for (let i = 0; i < enrichedRecords.length; i++) {
        const record = enrichedRecords[i];
        const rowNum = dataStartRow + i;
        
        // Check if row exists
        const row = worksheet.getRow(rowNum);
        if (!row) continue;
        
        let rowUpdated = false;
        
        // OVERWRITE gender - always update with enriched data
        if (gioiTinhCol !== -1) {
            const cell = row.getCell(gioiTinhCol);
            const oldValue = cell.value || '(empty)';
            const newValue = record.gioiTinh || '';
            // Always update, even if empty (to clear old data)
            cell.value = newValue;
            if (oldValue !== newValue) {
                rowUpdated = true;
                if (i < 3) {
                    console.log(`📝 Row ${rowNum}: Gender "${oldValue}" -> "${newValue}"`);
                }
            }
        }
        
        // OVERWRITE ethnicity - always update with enriched data
        if (danTocCol !== -1) {
            const cell = row.getCell(danTocCol);
            const oldValue = cell.value || '(empty)';
            const newValue = record.danToc || '';
            cell.value = newValue;
            if (oldValue !== newValue) {
                rowUpdated = true;
                if (i < 3) {
                    console.log(`📝 Row ${rowNum}: Ethnicity "${oldValue}" -> "${newValue}"`);
                }
            }
        }
        
        // OVERWRITE birthplace - always update with enriched data
        if (noiSinhCol !== -1) {
            const cell = row.getCell(noiSinhCol);
            const oldValue = cell.value || '(empty)';
            const newValue = record.noiSinh || '';
            cell.value = newValue;
            if (oldValue !== newValue) {
                rowUpdated = true;
                if (i < 3) {
                    console.log(`📝 Row ${rowNum}: Birthplace "${oldValue}" -> "${newValue}"`);
                }
            }
        }
        
        if (rowUpdated) updatedRows++;
    }
    
    console.log(`✅ Updated ${updatedRows} rows with enriched data (overwriting existing values)`);
    await workbook.xlsx.writeFile(outputPath);
}

/**
 * Main enrich function
 * @param {string} targetPath - Path to target file
 * @param {string} sourcePath - Path to source file
 * @param {string} outputPath - Path for output file
 * @param {boolean} overwriteExisting - If true, overwrite existing data; if false, only fill empty fields
 */
async function enrichDataFromFiles(targetPath, sourcePath, outputPath, overwriteExisting = false) {
    console.log('📂 Reading target file...');
    const target = parseTargetExcel(targetPath);
    console.log(`📄 Target: ${target.totalRecords} records`);
    
    console.log('📂 Reading source file...');
    const source = parseSourceExcel(sourcePath);
    console.log(`📄 Source: ${source.totalRecords} records`);
    
    console.log(`🔍 Enriching data (overwrite mode: ${overwriteExisting})...`);
    
    const targetColMap = {
        hoTen: 'hoTen',
        ngaySinh: 'ngaySinh',
        gioiTinh: 'gioiTinh',
        danToc: 'danToc',
        noiSinh: 'noiSinh'
    };
    
    const sourceColMap = {
        hoTen: 'hoTen',
        ngaySinh: 'ngaySinh',
        gioiTinh: 'gioiTinh',
        danToc: 'danToc',
        noiSinh: 'noiSinh'
    };
    
    const result = enrichData(
        target.records,
        source.records,
        targetColMap,
        sourceColMap,
        overwriteExisting
    );
    
    console.log(`✅ Enriched ${result.enrichedCount} of ${result.totalRecords} records`);
    console.log(`✅ Matched ${result.matchedCount} records`);
    if (overwriteExisting) {
        console.log(`✅ Overwritten ${result.overwrittenCount} records`);
    }
    
    // Rebuild Excel file with enriched data
    console.log('📝 Writing enriched data to Excel...');
    await rebuildExcelFile(target, result.enrichedRecords, outputPath);
    console.log(`✅ File saved to: ${outputPath}`);
    
    return {
        totalRecords: result.totalRecords,
        enrichedCount: result.enrichedCount,
        matchedCount: result.matchedCount,
        overwrittenCount: result.overwrittenCount || 0,
        sourceRecords: source.totalRecords,
        outputPath: outputPath,
        overwriteMode: overwriteExisting
    };
}

module.exports = {
    enrichDataFromFiles,
    parseTargetExcel,
    parseSourceExcel,
    enrichData
};