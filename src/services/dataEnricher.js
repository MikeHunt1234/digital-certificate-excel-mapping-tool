const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ============ CONFIGURATION ============
// You can extend these with more variations
// Add to COLUMN_PATTERNS in dataEnricher.js
const COLUMN_PATTERNS = {
    hoTen: [
        'họ tên', 'họ và tên', 'họ, chữ đệm và tên', 'full name', 'name',
        'người được cấp', 'tên người được cấp'
    ],
    ho: ['họ', 'ho', 'last name', 'surname'],
    ten: ['tên', 'ten', 'first name', 'given name'],
    ngaySinh: [
        'ngày sinh', 'ngày tháng năm sinh', 'ngày, tháng, năm sinh',
        'date of birth', 'dob', 'birthdate', 'birth day', 'birthday'
    ],
    gioiTinh: [
        'giới tính', 'giới', 'gender', 'sex'
    ],
    danToc: [
        'dân tộc', 'ethnicity', 'ethnic group'
    ],
    noiSinh: [
        'nơi sinh', 'place of birth', 'birthplace'
    ],
    diaChi: [
        'địa chỉ', 'address', 'địa chỉ liên hệ', 'địa chỉ thường trú'
    ]
};

// ============ CORE FUNCTIONS ============

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
    
    // Fast path: already in dd/MM/yyyy format
    if (str.length === 10 && str[2] === '/' && str[5] === '/') {
        const day = parseInt(str.substring(0, 2));
        const month = parseInt(str.substring(3, 5));
        const year = parseInt(str.substring(6, 10));
        if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1900 && year < 2100) {
            return str;
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
    } catch (e) { /* Ignore */ }
    
    // If it's an Excel serial number
    if (typeof dateStr === 'number' && dateStr > 30000 && dateStr < 50000) {
        try {
            const date = new Date((dateStr - 25569) * 86400 * 1000);
            const day = date.getDate().toString().padStart(2, '0');
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        } catch (e) { /* Ignore */ }
    }
    
    return str;
}

/**
 * Detect column mapping from headers using pattern matching
 * Returns: { hoTen: index, ngaySinh: index, gioiTinh: index, danToc: index, noiSinh: index }
 */
function detectColumnMapping(headers) {
    const mapping = {
        hoTen: -1,
        ngaySinh: -1,
        gioiTinh: -1,
        danToc: -1,
        noiSinh: -1,
        diaChi: -1
    };
    
    // Score each column for each field
    const scores = {};
    
    headers.forEach((header, index) => {
        if (!header) return;
        const headerLower = String(header).trim().toLowerCase();
        
        // Calculate score for each field
        for (const [field, patterns] of Object.entries(COLUMN_PATTERNS)) {
            let score = 0;
            for (const pattern of patterns) {
                if (headerLower === pattern) {
                    score = 100; // Perfect match
                    break;
                } else if (headerLower.includes(pattern) || pattern.includes(headerLower)) {
                    score += 10;
                }
            }
            if (score > 0) {
                scores[`${field}_${index}`] = score;
            }
        }
    });
    
    // Select best match for each field
    const usedIndices = new Set();
    for (const [field, _] of Object.entries(COLUMN_PATTERNS)) {
        let bestIndex = -1;
        let bestScore = 0;
        
        for (const [key, score] of Object.entries(scores)) {
            const [f, idx] = key.split('_');
            if (f === field && score > bestScore && !usedIndices.has(parseInt(idx))) {
                bestScore = score;
                bestIndex = parseInt(idx);
            }
        }
        
        if (bestIndex !== -1) {
            mapping[field] = bestIndex;
            usedIndices.add(bestIndex);
        }
    }
    
    return mapping;
}

/**
 * Extract column mappings from user-provided configuration
 */
function getUserColumnMapping(headers, userMapping) {
    const mapping = {
        hoTen: -1,
        ngaySinh: -1,
        gioiTinh: -1,
        danToc: -1,
        noiSinh: -1,
        diaChi: -1
    };
    
    if (userMapping && Object.keys(userMapping).length > 0) {
        // Use user-provided mapping
        for (const [field, columnName] of Object.entries(userMapping)) {
            const index = headers.findIndex(h => 
                h && String(h).trim().toLowerCase() === columnName.toLowerCase()
            );
            if (index !== -1) {
                mapping[field] = index;
            }
        }
    }
    
    return mapping;
}

/**
 * Parse source file with flexible schema detection
 */
/**
 * Parse source Excel file - AUTO-DETECTS FORMAT
 * Supports multiple formats:
 * - Format A: "Họ tên", "Ngày sinh", "Giới tính", "Dân tộc", "Nơi sinh"
 * - Format B: "NAME", "GIOITINH", "NGAYSINH", "DANTOC", "NOISINH"
 * - Format C: "Họ" + "TÊN" split columns, "X" for gender
 * - Multiple sheets support
 */
function parseSourceExcel(filePath, userMapping = null) {
    console.log('📂 Parsing source file with flexible schema detection:', filePath);
    const startTime = Date.now();
    
    const workbook = XLSX.readFile(filePath);
    
    // Get all sheets, but skip metadata/instruction sheets
    const dataSheets = [];
    const skipPatterns = ['ds vang', 'bo tri', 'metadata', 'instruction', 'hướng dẫn'];
    
    for (const sheetName of workbook.SheetNames) {
        const lowerName = sheetName.toLowerCase();
        // Skip sheets that look like metadata or instructions
        if (skipPatterns.some(pattern => lowerName.includes(pattern))) {
            console.log(`⏭️ Skipping metadata sheet: "${sheetName}"`);
            continue;
        }
        dataSheets.push(sheetName);
    }
    
    // If no data sheets found, use all sheets
    if (dataSheets.length === 0) {
        console.log('⚠️ No data sheets found, using all sheets');
        dataSheets.push(...workbook.SheetNames);
    }
    
    console.log(`📄 Found ${dataSheets.length} data sheet(s): ${dataSheets.join(', ')}`);
    
    let allRecords = [];
    let allStats = {
        totalRows: 0,
        emptyNameCount: 0,
        emptyBirthCount: 0,
        defaultEthnicityCount: 0,
        addressFallbackCount: 0,
        sheetsProcessed: 0,
        nameCombinedCount: 0
    };
    
    // Process each data sheet
    for (const sheetName of dataSheets) {
        console.log(`\n📋 Processing sheet: "${sheetName}"`);
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        console.log(`   📄 Sheet has ${rawData.length.toLocaleString()} rows`);
        
        // Log first few rows for debugging
        console.log('   📋 First 5 rows:');
        for (let i = 0; i < Math.min(rawData.length, 5); i++) {
            const row = rawData[i];
            if (row) {
                const displayRow = row.slice(0, 8).map(cell => String(cell).substring(0, 20)).join(' | ');
                console.log(`   Row ${i}: ${displayRow}`);
            }
        }
        
        // Find header row (first row with substantial content)
        let headerRowIndex = -1;
        let dataStartRow = -1;
        
        for (let i = 0; i < Math.min(rawData.length, 30); i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            
            const nonEmpty = row.filter(cell => String(cell).trim() !== '');
            if (nonEmpty.length >= 3) {
                headerRowIndex = i;
                dataStartRow = i + 1;
                console.log(`   ✅ Found header row at index ${i} with ${nonEmpty.length} columns`);
                break;
            }
        }
        
        if (headerRowIndex === -1) {
            console.log(`   ⚠️ Could not find header row in sheet "${sheetName}", skipping`);
            continue;
        }
        
        const headers = rawData[headerRowIndex];
        console.log('   📋 Headers found:', headers.slice(0, 15));
        
        // Detect column mapping
        let colMap;
        if (userMapping && Object.keys(userMapping).length > 0) {
            console.log('   📋 Using user-provided column mapping');
            colMap = getUserColumnMapping(headers, userMapping);
        } else {
            console.log('   📋 Auto-detecting column mapping...');
            colMap = detectColumnMapping(headers);
        }
        
        console.log('   📋 Column mapping:', colMap);
        
        // Log which columns were found
        const foundFields = Object.entries(colMap)
            .filter(([_, idx]) => idx !== -1)
            .map(([field, idx]) => `${field} -> column ${idx + 1} ("${headers[idx] || ''}")`);
        console.log(`   ✅ Found: ${foundFields.join(', ')}`);
        
        const missingFields = Object.entries(colMap)
            .filter(([_, idx]) => idx === -1)
            .map(([field]) => field);
        if (missingFields.length > 0) {
            console.log(`   ⚠️ Missing fields: ${missingFields.join(', ')}`);
        }
        
        // Check if we have split name columns (Họ + TÊN)
        const hasSplitName = colMap.ho !== -1 && colMap.ten !== -1;
        const hasCombinedName = colMap.hoTen !== -1;
        
        if (hasSplitName && !hasCombinedName) {
            console.log('   📋 Using split name columns (Họ + TÊN)');
        }
        
        // Parse data rows from this sheet
        const sheetRecords = [];
        let sheetProcessed = 0;
        let sheetEmptyName = 0;
        let sheetEmptyBirth = 0;
        let sheetDefaultEthnicity = 0;
        let sheetAddressFallback = 0;
        let sheetNameCombined = 0;
        
        for (let i = dataStartRow; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            
            // Get name - try combined first, then split
            let hoTen = '';
            let ho = '';
            let ten = '';
            
            if (colMap.hoTen !== -1 && row[colMap.hoTen] !== undefined && row[colMap.hoTen] !== null) {
                hoTen = String(row[colMap.hoTen]).trim();
            }
            
            // If combined name is empty or not found, try split columns
            if (!hoTen && hasSplitName) {
                if (colMap.ho !== -1 && row[colMap.ho] !== undefined && row[colMap.ho] !== null) {
                    ho = String(row[colMap.ho]).trim();
                }
                if (colMap.ten !== -1 && row[colMap.ten] !== undefined && row[colMap.ten] !== null) {
                    ten = String(row[colMap.ten]).trim();
                }
                if (ho || ten) {
                    hoTen = `${ho} ${ten}`.trim();
                    sheetNameCombined++;
                }
            }
            
            // Also check if we have a "TÊN" column that's actually the full name
            if (!hoTen && colMap.ten !== -1 && colMap.ho === -1) {
                const tenVal = row[colMap.ten];
                if (tenVal !== undefined && tenVal !== null) {
                    const tenStr = String(tenVal).trim();
                    // If it looks like a full name (has space), use it
                    if (tenStr.includes(' ') && tenStr.length > 10) {
                        hoTen = tenStr;
                    }
                }
            }
            
            if (!hoTen) {
                sheetEmptyName++;
                continue;
            }
            
            sheetProcessed++;
            
            // Progress logging every 10k records
            if (sheetProcessed % 10000 === 0) {
                console.log(`   📊 Processed ${sheetProcessed.toLocaleString()} records in sheet...`);
            }
            
            // Extract birthdate
            let birthdate = '';
            if (colMap.ngaySinh !== -1) {
                const birthCell = row[colMap.ngaySinh];
                if (birthCell !== undefined && birthCell !== null && birthCell !== '') {
                    birthdate = parseBirthDate(birthCell);
                }
            }
            
            if (!birthdate) {
                sheetEmptyBirth++;
            }
            
            // Extract gender - handle "X" as female
            let gender = '';
            if (colMap.gioiTinh !== -1) {
                const genderCell = row[colMap.gioiTinh];
                if (genderCell !== undefined && genderCell !== null) {
                    const genderStr = String(genderCell).trim().toUpperCase();
                    if (genderStr === 'X' || genderStr === 'NỮ' || genderStr === 'NU') {
                        gender = 'Nữ';
                    } else if (genderStr === 'NAM' || genderStr === '1' || genderStr === 'M' || genderStr === 'MALE') {
                        gender = 'Nam';
                    } else if (genderStr) {
                        // Keep as-is if it's already "Nam" or "Nữ"
                        const normalized = normalizeGender(genderStr);
                        if (normalized) gender = normalized;
                    }
                }
            }
            
            // Extract ethnicity with default
            let danToc = 'Kinh';
            if (colMap.danToc !== -1) {
                const ethnicityCell = row[colMap.danToc];
                if (ethnicityCell !== undefined && ethnicityCell !== null) {
                    const ethnicityValue = String(ethnicityCell).trim();
                    if (ethnicityValue && ethnicityValue !== '') {
                        danToc = ethnicityValue;
                    } else {
                        sheetDefaultEthnicity++;
                    }
                } else {
                    sheetDefaultEthnicity++;
                }
            } else {
                sheetDefaultEthnicity++;
            }
            
            // Extract birthplace with address fallback
            let noiSinh = '';
            if (colMap.noiSinh !== -1) {
                const birthplaceCell = row[colMap.noiSinh];
                if (birthplaceCell !== undefined && birthplaceCell !== null) {
                    const birthplaceValue = String(birthplaceCell).trim();
                    if (birthplaceValue && birthplaceValue !== '') {
                        noiSinh = birthplaceValue;
                    }
                }
            }
            
            // Fallback to address if available and noisinh is empty
            if (!noiSinh && colMap.diaChi !== -1) {
                const addressCell = row[colMap.diaChi];
                if (addressCell !== undefined && addressCell !== null) {
                    const addressValue = String(addressCell).trim();
                    if (addressValue && addressValue !== '') {
                        noiSinh = addressValue;
                        sheetAddressFallback++;
                    }
                }
            }
            
            sheetRecords.push({
                hoTen: hoTen,
                ngaySinh: birthdate,
                gioiTinh: gender,
                danToc: danToc,
                noiSinh: noiSinh,
                // Keep original data for debugging
                _ho: ho,
                _ten: ten,
                _sheet: sheetName
            });
        }
        
        // Add sheet stats to total
        allStats.totalRows += sheetProcessed;
        allStats.emptyNameCount += sheetEmptyName;
        allStats.emptyBirthCount += sheetEmptyBirth;
        allStats.defaultEthnicityCount += sheetDefaultEthnicity;
        allStats.addressFallbackCount += sheetAddressFallback;
        allStats.nameCombinedCount += sheetNameCombined;
        allStats.sheetsProcessed++;
        
        console.log(`   ✅ Parsed ${sheetRecords.length.toLocaleString()} records from sheet "${sheetName}"`);
        console.log(`   📊 Empty name: ${sheetEmptyName.toLocaleString()}`);
        console.log(`   📊 Empty birthdate: ${sheetEmptyBirth.toLocaleString()}`);
        console.log(`   📊 Default ethnicity: ${sheetDefaultEthnicity.toLocaleString()}`);
        console.log(`   📊 Address fallback: ${sheetAddressFallback.toLocaleString()}`);
        if (sheetNameCombined > 0) {
            console.log(`   📊 Name combined from Họ + TÊN: ${sheetNameCombined.toLocaleString()}`);
        }
        
        allRecords = allRecords.concat(sheetRecords);
    }
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`\n✅ Parsed ${allRecords.length.toLocaleString()} total records from ${allStats.sheetsProcessed} sheet(s) in ${elapsed}s`);
    console.log(`   📊 Empty name: ${allStats.emptyNameCount.toLocaleString()}`);
    console.log(`   📊 Empty birthdate: ${allStats.emptyBirthCount.toLocaleString()}`);
    console.log(`   📊 Default ethnicity "Kinh": ${allStats.defaultEthnicityCount.toLocaleString()}`);
    console.log(`   📊 Address fallback: ${allStats.addressFallbackCount.toLocaleString()}`);
    if (allStats.nameCombinedCount > 0) {
        console.log(`   📊 Name combined from split columns: ${allStats.nameCombinedCount.toLocaleString()}`);
    }
    
    // Log first few records for debugging
    if (allRecords.length > 0) {
        console.log('\n📋 Sample records:');
        for (let i = 0; i < Math.min(allRecords.length, 5); i++) {
            const r = allRecords[i];
            console.log(`   Record ${i + 1}: Name="${r.hoTen}", Birth="${r.ngaySinh}", Gender="${r.gioiTinh}", Ethnicity="${r.danToc}", Birthplace="${r.noiSinh}"`);
        }
    }
    
    return {
        records: allRecords,
        totalRecords: allRecords.length,
        stats: allStats,
        sheets: dataSheets,
        elapsedSeconds: parseFloat(elapsed)
    };
}

/**
 * Parse birth date from various formats
 */
function parseBirthDate(birthCell) {
    if (birthCell === undefined || birthCell === null || birthCell === '') return '';
    
    // If it's a string
    if (typeof birthCell === 'string') {
        const trimmed = birthCell.trim();
        
        // Check if it's dd/mm/yyyy format
        if (trimmed.length === 10 && trimmed[2] === '/' && trimmed[5] === '/') {
            const day = parseInt(trimmed.substring(0, 2));
            const month = parseInt(trimmed.substring(3, 5));
            const year = parseInt(trimmed.substring(6, 10));
            if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1900 && year < 2100) {
                return trimmed;
            }
        }
        
        // Check if it's yyyy-mm-dd format
        if (trimmed.length === 10 && trimmed[4] === '-' && trimmed[7] === '-') {
            const year = parseInt(trimmed.substring(0, 4));
            const month = parseInt(trimmed.substring(5, 7));
            const day = parseInt(trimmed.substring(8, 10));
            if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1900 && year < 2100) {
                return `${day.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/${year}`;
            }
        }
        
        // Try parsing with Date
        try {
            const date = new Date(trimmed);
            if (!isNaN(date.getTime()) && date.getFullYear() > 1900) {
                const day = date.getDate().toString().padStart(2, '0');
                const month = (date.getMonth() + 1).toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${day}/${month}/${year}`;
            }
        } catch (e) { /* Ignore */ }
        
        // If it's just a year (4 digits)
        if (/^\d{4}$/.test(trimmed)) {
            const year = parseInt(trimmed);
            if (year > 1900 && year < 2100) {
                return `01/01/${year}`;
            }
        }
        
        return trimmed;
    }
    
    // If it's a number (Excel serial date or just a year)
    if (typeof birthCell === 'number') {
        // If it's between 1900 and 2100, it's probably just a year
        if (birthCell >= 1900 && birthCell <= 2100) {
            return `01/01/${Math.floor(birthCell)}`;
        }
        // Otherwise it's an Excel serial date
        try {
            const date = new Date((birthCell - 25569) * 86400 * 1000);
            if (!isNaN(date.getTime())) {
                const day = date.getDate().toString().padStart(2, '0');
                const month = (date.getMonth() + 1).toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${day}/${month}/${year}`;
            }
        } catch (e) { /* Ignore */ }
    }
    
    // If it's a Date object
    if (birthCell instanceof Date) {
        try {
            const day = birthCell.getDate().toString().padStart(2, '0');
            const month = (birthCell.getMonth() + 1).toString().padStart(2, '0');
            const year = birthCell.getFullYear();
            return `${day}/${month}/${year}`;
        } catch (e) { /* Ignore */ }
    }
    
    return String(birthCell).trim();
}

/**
 * Normalize gender string
 */
function normalizeGender(genderStr) {
    if (!genderStr) return '';
    const g = genderStr.trim().toUpperCase();
    if (g === 'X' || g === 'NỮ' || g === 'NU' || g === 'F' || g === 'FEMALE' || g === '0') {
        return 'Nữ';
    }
    if (g === 'NAM' || g === 'M' || g === 'MALE' || g === '1') {
        return 'Nam';
    }
    return genderStr.trim();
}

/**
 * Build a lookup map from source records - optimized for large datasets
 */
function buildLookupMap(sourceRecords) {
    console.log(`🔍 Building lookup map from ${sourceRecords.length.toLocaleString()} records...`);
    const startTime = Date.now();
    
    const lookupMap = new Map();
    let duplicateCount = 0;
    let skippedCount = 0;
    
    for (let i = 0; i < sourceRecords.length; i++) {
        const record = sourceRecords[i];
        const fullName = record.hoTen || '';
        const birthdate = record.ngaySinh || '';
        
        if (!fullName || !birthdate) {
            skippedCount++;
            continue;
        }
        
        const key = normalizeString(fullName) + '|' + normalizeDate(birthdate);
        
        const data = {
            gender: record.gioiTinh || '',
            ethnicity: record.danToc || 'Kinh',
            birthplace: record.noiSinh || ''
        };
        
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
        
        if ((i + 1) % 10000 === 0) {
            console.log(`   📊 Processed ${(i + 1).toLocaleString()} records... (${lookupMap.size.toLocaleString()} unique keys)`);
        }
    }
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`✅ Lookup map built in ${elapsed}s: ${lookupMap.size.toLocaleString()} unique keys`);
    console.log(`   📊 ${duplicateCount.toLocaleString()} duplicates, ${skippedCount.toLocaleString()} skipped`);
    
    return lookupMap;
}

/**
 * Parse target Excel file (template-mapped)
 */
function parseTargetExcel(filePath) {
    console.log('📂 Parsing target file:', filePath);
    const workbook = XLSX.readFile(filePath);
    
    // Try to find the "Data" sheet first
    let sheetName = 'Data';
    if (!workbook.SheetNames.includes('Data')) {
        const possibleNames = ['Data', 'Dữ liệu', 'Danh sách', 'Sheet1'];
        for (const name of possibleNames) {
            if (workbook.SheetNames.includes(name)) {
                sheetName = name;
                break;
            }
        }
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
    
    // Find header row
    let headerRowIndex = -1;
    let dataStartRow = -1;
    
    for (let i = 0; i < Math.min(rawData.length, 30); i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        const rowStr = row.join(' ');
        
        if (rowStr.includes('STT') && 
            (rowStr.includes('Số hiệu chứng chỉ') || 
             rowStr.includes('Họ, chữ đệm và tên') || 
             rowStr.includes('Họ tên'))) {
            headerRowIndex = i;
            dataStartRow = i + 1;
            console.log(`✅ Found target header row at index ${i}`);
            break;
        }
    }
    
    if (headerRowIndex === -1) {
        // Try alternate detection
        for (let i = 0; i < Math.min(rawData.length, 30); i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            const rowStr = row.join(' ');
            if (rowStr.includes('STT') && 
                (rowStr.includes('chứng chỉ') || 
                 rowStr.includes('giới tính') || 
                 rowStr.includes('dân tộc') || 
                 rowStr.includes('nơi sinh'))) {
                headerRowIndex = i;
                dataStartRow = i + 1;
                console.log(`✅ Found target header row at index ${i} (alternate)`);
                break;
            }
        }
    }
    
    // Last resort: find first row with STT
    if (headerRowIndex === -1) {
        for (let i = 0; i < Math.min(rawData.length, 30); i++) {
            const row = rawData[i];
            if (!row) continue;
            const nonEmpty = row.filter(cell => String(cell).trim() !== '');
            if (nonEmpty.length >= 5) {
                const firstCell = row[0] ? String(row[0]).trim() : '';
                if (!isNaN(firstCell) && firstCell !== '') {
                    headerRowIndex = i - 1;
                    dataStartRow = i;
                    console.log(`✅ Found data row at index ${i}, using row ${headerRowIndex} as header`);
                    break;
                }
            }
        }
    }
    
    if (headerRowIndex === -1 || headerRowIndex < 0) {
        throw new Error('Could not find header row in target file');
    }
    
    const headers = rawData[headerRowIndex];
    console.log('📋 Target headers found:', headers.slice(0, 15));
    
    // Map column indices
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
        
        if (headerStr === 'STT' || headerStr.includes('STT')) {
            colMap.stt = index;
        } else if (headerStr.includes('SBD') || headerStr.includes('Số hiệu') || headerStr.includes('Số vào sổ')) {
            colMap.sbd = index;
        } else if (headerStr.includes('Họ, chữ đệm và tên') || headerStr.includes('Họ tên') || headerStr.includes('Họ và tên')) {
            colMap.hoTen = index;
        } else if (headerStr.includes('Ngày, tháng, năm sinh') || headerStr.includes('Ngày sinh')) {
            colMap.ngaySinh = index;
        } else if (headerStr.includes('Giới tính')) {
            colMap.gioiTinh = index;
        } else if (headerStr.includes('Dân tộc')) {
            colMap.danToc = index;
        } else if (headerStr.includes('Nơi sinh')) {
            colMap.noiSinh = index;
        }
    });
    
    console.log('📋 Target column map:', colMap);
    
    // Extract data rows
    const records = [];
    for (let i = dataStartRow; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        const stt = row[colMap.stt] !== undefined && row[colMap.stt] !== null ? String(row[colMap.stt]).trim() : '';
        const name = colMap.hoTen !== -1 && row[colMap.hoTen] !== undefined && row[colMap.hoTen] !== null ? String(row[colMap.hoTen]).trim() : '';
        
        if (!stt && !name) continue;
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
    
    console.log(`✅ Parsed ${records.length.toLocaleString()} records from target file`);
    
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
 */
function enrichData(targetRecords, sourceRecords, overwriteExisting = false) {
    console.log('🔍 Enriching data...');
    const startTime = Date.now();
    
    const lookupMap = buildLookupMap(sourceRecords);
    
    const usageTracker = new Map();
    let enrichedCount = 0;
    let matchedCount = 0;
    let overwrittenCount = 0;
    
    const enrichedRecords = targetRecords.map((record, index) => {
        const fullName = record.hoTen || '';
        const birthdate = record.ngaySinh || '';
        
        if (!fullName || !birthdate) {
            return record;
        }
        
        const key = normalizeString(fullName) + '|' + normalizeDate(birthdate);
        
        if (!lookupMap.has(key)) {
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
                return record;
            }
        }
        
        const newRecord = { ...record };
        let fieldsChanged = false;
        
        const shouldUpdateGender = overwriteExisting || !record.gioiTinh || record.gioiTinh.trim() === '';
        const shouldUpdateEthnicity = overwriteExisting || !record.danToc || record.danToc.trim() === '';
        const shouldUpdateBirthplace = overwriteExisting || !record.noiSinh || record.noiSinh.trim() === '';
        
        if (shouldUpdateGender && matchData.gender) {
            newRecord.gioiTinh = matchData.gender;
            fieldsChanged = true;
        }
        
        if (shouldUpdateEthnicity && matchData.ethnicity) {
            newRecord.danToc = matchData.ethnicity;
            fieldsChanged = true;
        }
        
        if (shouldUpdateBirthplace && matchData.birthplace) {
            newRecord.noiSinh = matchData.birthplace;
            fieldsChanged = true;
        }
        
        if (fieldsChanged) {
            enrichedCount++;
            if (overwriteExisting) overwrittenCount++;
        }
        
        return newRecord;
    });
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`✅ Enrichment complete in ${elapsed}s: ${enrichedCount.toLocaleString()} of ${targetRecords.length.toLocaleString()} records updated`);
    console.log(`   📊 ${matchedCount.toLocaleString()} records matched, ${overwrittenCount.toLocaleString()} overwritten`);
    
    return {
        enrichedRecords,
        enrichedCount,
        totalRecords: targetRecords.length,
        matchedCount: matchedCount,
        overwrittenCount: overwrittenCount,
        lookupMapSize: lookupMap.size
    };
}

/**
 * Rebuild Excel file with enriched data
 */
async function rebuildExcelFile(target, enrichedRecords, outputPath) {
    console.log('📝 Writing enriched data to Excel...');
    const startTime = Date.now();
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(target.filePath);
    
    const worksheet = workbook.getWorksheet(target.sheetName || 'Data');
    if (!worksheet) {
        throw new Error(`Could not find "${target.sheetName || 'Data'}" worksheet`);
    }
    
    // Find data start row
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
    
    // Find column indices
    const headerRow = worksheet.getRow(dataStartRow - 1);
    let gioiTinhCol = -1, danTocCol = -1, noiSinhCol = -1;
    
    for (let col = 1; col <= 29; col++) {
        const cell = headerRow.getCell(col);
        const value = cell.value ? String(cell.value).trim() : '';
        if (value.includes('Giới tính')) gioiTinhCol = col;
        if (value.includes('Dân tộc')) danTocCol = col;
        if (value.includes('Nơi sinh')) noiSinhCol = col;
    }
    
    if (gioiTinhCol === -1) gioiTinhCol = 10;
    if (danTocCol === -1) danTocCol = 11;
    if (noiSinhCol === -1) noiSinhCol = 13;
    
    let updatedRows = 0;
    for (let i = 0; i < enrichedRecords.length; i++) {
        const record = enrichedRecords[i];
        const rowNum = dataStartRow + i;
        const row = worksheet.getRow(rowNum);
        if (!row) continue;
        
        let rowUpdated = false;
        
        if (gioiTinhCol !== -1) {
            const cell = row.getCell(gioiTinhCol);
            cell.value = record.gioiTinh || '';
            rowUpdated = true;
        }
        
        if (danTocCol !== -1) {
            const cell = row.getCell(danTocCol);
            cell.value = record.danToc || '';
            rowUpdated = true;
        }
        
        if (noiSinhCol !== -1) {
            const cell = row.getCell(noiSinhCol);
            cell.value = record.noiSinh || '';
            rowUpdated = true;
        }
        
        if (rowUpdated) updatedRows++;
    }
    
    await workbook.xlsx.writeFile(outputPath);
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`✅ Excel file written in ${elapsed}s: ${updatedRows.toLocaleString()} rows updated`);
}

/**
 * Main enrich function - SUPPORTS MULTIPLE SOURCE FILES
 * @param {string} targetPath - Path to target file
 * @param {string|string[]} sourcePaths - Path to source file(s) - can be single or array
 * @param {string} outputPath - Path for output file
 * @param {Object} options - Configuration options
 * @param {Object} options.userMapping - User-defined column mapping
 * @param {boolean} options.overwriteExisting - If true, overwrite existing data
 */
async function enrichDataFromFiles(targetPath, sourcePaths, outputPath, options = {}) {
    const {
        userMapping = null,
        overwriteExisting = false
    } = options;
    
    console.log('🚀 Starting flexible data enrichment...');
    console.log(`   📋 Overwrite mode: ${overwriteExisting ? 'ON' : 'OFF'}`);
    console.log(`   📋 User mapping: ${userMapping ? 'PROVIDED' : 'AUTO-DETECT'}`);
    
    const startTime = Date.now();
    
    // Parse target file
    console.log('\n📂 Reading target file...');
    const target = parseTargetExcel(targetPath);
    console.log(`📄 Target: ${target.totalRecords.toLocaleString()} records`);
    
    // Parse ALL source files and merge records
    console.log('\n📂 Reading source files...');
    
    // Normalize sourcePaths to always be an array
    const sourcePathsArray = Array.isArray(sourcePaths) ? sourcePaths : [sourcePaths];
    console.log(`   📋 Source files: ${sourcePathsArray.length}`);
    
    let allSourceRecords = [];
    let sourceStats = {
        totalFiles: sourcePathsArray.length,
        totalRecords: 0,
        filesProcessed: 0,
        totalEmptyName: 0,
        totalEmptyBirth: 0,
        totalDefaultEthnicity: 0,
        totalAddressFallback: 0
    };
    
    for (let i = 0; i < sourcePathsArray.length; i++) {
        const sourcePath = sourcePathsArray[i];
        const fileName = path.basename(sourcePath);
        console.log(`\n📄 Processing source file ${i + 1}/${sourcePathsArray.length}: ${fileName}`);
        
        try {
            // Make sure we have a valid path string
            if (typeof sourcePath !== 'string' || !sourcePath) {
                console.error(`   ❌ Invalid source path: ${sourcePath}`);
                continue;
            }
            
            // Check if file exists
            if (!fs.existsSync(sourcePath)) {
                console.error(`   ❌ File not found: ${sourcePath}`);
                continue;
            }
            
            const source = parseSourceExcel(sourcePath, userMapping);
            console.log(`   ✅ Parsed ${source.records.length.toLocaleString()} records`);
            
            allSourceRecords = allSourceRecords.concat(source.records);
            sourceStats.totalRecords += source.records.length;
            sourceStats.filesProcessed++;
            
            if (source.stats) {
                sourceStats.totalEmptyName += source.stats.emptyNameCount || 0;
                sourceStats.totalEmptyBirth += source.stats.emptyBirthCount || 0;
                sourceStats.totalDefaultEthnicity += source.stats.defaultEthnicityCount || 0;
                sourceStats.totalAddressFallback += source.stats.addressFallbackCount || 0;
            }
        } catch (error) {
            console.error(`   ❌ Error parsing file ${fileName}:`, error.message);
            // Continue with other files
        }
    }
    
    console.log(`\n📊 Total source records: ${allSourceRecords.length.toLocaleString()} from ${sourceStats.filesProcessed} file(s)`);
    console.log(`   📊 Empty name: ${sourceStats.totalEmptyName.toLocaleString()}`);
    console.log(`   📊 Empty birthdate: ${sourceStats.totalEmptyBirth.toLocaleString()}`);
    console.log(`   📊 Default ethnicity: ${sourceStats.totalDefaultEthnicity.toLocaleString()}`);
    console.log(`   📊 Address fallback: ${sourceStats.totalAddressFallback.toLocaleString()}`);
    
    if (allSourceRecords.length === 0) {
        throw new Error('No valid records found in any source file');
    }
    
    // Enrich data
    console.log('\n🔍 Enriching data...');
    const result = enrichData(
        target.records,
        allSourceRecords,
        overwriteExisting
    );
    
    // Rebuild Excel file
    console.log('\n📝 Writing enriched data to Excel...');
    await rebuildExcelFile(target, result.enrichedRecords, outputPath);
    
    const totalElapsed = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`\n✅ Enrichment completed in ${totalElapsed}s`);
    console.log(`   📊 ${result.enrichedCount.toLocaleString()} records enriched out of ${result.totalRecords.toLocaleString()}`);
    console.log(`   📊 ${result.matchedCount.toLocaleString()} records matched`);
    console.log(`   📊 ${sourceStats.filesProcessed} source file(s) processed`);
    console.log(`   📊 ${allSourceRecords.length.toLocaleString()} total source records`);
    console.log(`   📊 File saved to: ${outputPath}`);
    
    return {
        totalRecords: result.totalRecords,
        enrichedCount: result.enrichedCount,
        matchedCount: result.matchedCount,
        overwrittenCount: result.overwrittenCount || 0,
        sourceFilesProcessed: sourceStats.filesProcessed,
        sourceRecordsTotal: allSourceRecords.length,
        outputPath: outputPath,
        elapsedSeconds: parseFloat(totalElapsed),
        sourceStats: sourceStats
    };
}

module.exports = {
    enrichDataFromFiles,
    parseTargetExcel,
    parseSourceExcel,
    enrichData,
    detectColumnMapping,
    COLUMN_PATTERNS
};