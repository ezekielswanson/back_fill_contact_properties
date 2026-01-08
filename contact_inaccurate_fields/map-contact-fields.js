const xlsx = require('xlsx');
const path = require('path');

/**
 * Excel Email Field Mapping Script
 * 
 * Reads an Excel file with two sheets, matches emails between them,
 * handles multiple Stripe IDs per email by selecting the one with the latest create date,
 * and outputs a new Excel file with updated field mappings.
 */

// Configuration
const INPUT_FILE = path.join(__dirname, 'input_files/subscriptions_contact_map_fields.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'input_files/contacts_with_updated_fields.xlsx');

// Sheet names
const SHEET1_NAME = 'Export For Stripe Subs Field Up';
const SHEET2_NAME = 'Export For Contact Field Update';

// Columns to map from Sheet 1 to Sheet 2
const COLUMNS_TO_MAP = [
    'Billing Start Date',
    'Billing End Date',
    'Status',
    'Products',
    'Discount',
    'Coupon'
];

/**
 * Converts Excel serial date number to JavaScript Date object.
 * Handles potential non-numeric input gracefully.
 */
function excelSerialDateToJSDate(serial) {
    const serialNumber = Number(serial);
    if (isNaN(serialNumber) || serialNumber <= 0) {
        // Handle non-numeric input or dates before 1900
        // Also handle cases where the date might already be in a standard string format
        if (typeof serial === 'string' && !isNaN(Date.parse(serial))) {
            return new Date(serial);
        }
        return null;
    }

    // Excel's epoch starts January 1, 1900 (Windows Excel format)
    // Base date: Dec 30, 1899. Days are 1-based from Jan 1, 1900.
    const excelEpochOffsetDays = 25569; // Days between 1970-01-01 and 1899-12-30
    const daysInMilliseconds = (serialNumber - excelEpochOffsetDays) * 86400000; // 24 * 60 * 60 * 1000

    const jsDate = new Date(daysInMilliseconds);

    // Validate the date
    if (jsDate.getFullYear() < 1900 || jsDate.getFullYear() > 2100) {
        // Attempt parsing as a string just in case
        if (typeof serial === 'string' && !isNaN(Date.parse(serial))) {
            return new Date(serial);
        }
        return null;
    }

    return jsDate;
}

/**
 * Parses a date value from Excel (handles serial numbers and string formats)
 * Returns a Date object for comparison, or null if invalid
 */
function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) return dateValue;
    
    // Handle Excel date serial number
    if (typeof dateValue === 'number') {
        return excelSerialDateToJSDate(dateValue);
    }
    
    // Handle string dates
    if (typeof dateValue === 'string') {
        const parsed = new Date(dateValue);
        return isNaN(parsed.getTime()) ? null : parsed;
    }
    
    return null;
}

/**
 * Formats a date to a standardized string format for output
 */
function formatDate(date) {
    if (!date) return '';
    if (!(date instanceof Date)) {
        const parsed = parseDate(date);
        if (!parsed) return date; // Return original if can't parse
        date = parsed;
    }
    
    // Format as MM/DD/YYYY HH:MM:SS
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    
    return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
}

/**
 * Main function to process the Excel file and map fields
 */
async function mapContactFields() {
    console.log('🔄 Starting Excel Email Field Mapping Process');
    console.log('='.repeat(60));
    
    try {
        // Step 1: Load the workbook
        console.log('\n📖 STEP 1: Loading workbook...');
        const workbook = xlsx.readFile(INPUT_FILE);
        console.log(`   ✅ Loaded: ${INPUT_FILE}`);
        console.log(`   📋 Sheets found: ${workbook.SheetNames.join(', ')}`);
        
        // Verify sheets exist
        if (!workbook.SheetNames.includes(SHEET1_NAME)) {
            throw new Error(`Sheet "${SHEET1_NAME}" not found in workbook`);
        }
        if (!workbook.SheetNames.includes(SHEET2_NAME)) {
            throw new Error(`Sheet "${SHEET2_NAME}" not found in workbook`);
        }
        
        // Step 2: Load and process Sheet 1 data
        console.log('\n📊 STEP 2: Processing Sheet 1 (Stripe Subs)...');
        const sheet1Data = processSheet1(workbook);
        
        // Step 3: Load and process Sheet 2 data
        console.log('\n📊 STEP 3: Processing Sheet 2 (Contact Field Update)...');
        const sheet2Data = processSheet2(workbook);
        
        // Step 4: Match emails and map fields
        console.log('\n🔄 STEP 4: Matching emails and mapping fields...');
        const matchedData = matchAndMapFields(sheet1Data, sheet2Data);
        
        // Step 5: Create output Excel file
        console.log('\n📁 STEP 5: Creating output Excel file...');
        await createOutputFile(matchedData);
        
        // Step 6: Display results
        displayResults(matchedData);
        
        console.log('\n✅ PROCESS COMPLETE!');
        console.log(`📁 Output file: ${OUTPUT_FILE}`);
        
    } catch (error) {
        console.error('\n❌ Error during processing:', error.message);
        console.error(error.stack);
        process.exit(1);
    }
}

/**
 * Process Sheet 1: Group by email and select latest row per email
 */
function processSheet1(workbook) {
    const worksheet = workbook.Sheets[SHEET1_NAME];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    
    console.log(`   📋 Total rows in Sheet 1: ${jsonData.length}`);
    
    // Group rows by email
    const emailGroups = new Map();
    
    jsonData.forEach((row, index) => {
        const email = row['Stripe Customer Email'];
        if (!email || typeof email !== 'string') {
            return; // Skip rows without valid email
        }
        
        const normalizedEmail = email.toLowerCase().trim();
        const createDate = parseDate(row['Most Recent Create Date']);
        
        if (!emailGroups.has(normalizedEmail)) {
            emailGroups.set(normalizedEmail, []);
        }
        
        emailGroups.get(normalizedEmail).push({
            rowIndex: index,
            row: row,
            createDate: createDate,
            email: normalizedEmail
        });
    });
    
    console.log(`   📧 Unique emails found: ${emailGroups.size}`);
    
    // For each email, select the row with the latest create date
    const emailToRowMap = new Map();
    let multipleIdCount = 0;
    
    emailGroups.forEach((rows, email) => {
        if (rows.length > 1) {
            multipleIdCount++;
        }
        
        // Sort by create date (latest first), handling null dates
        rows.sort((a, b) => {
            if (!a.createDate && !b.createDate) return 0;
            if (!a.createDate) return 1; // null dates go to end
            if (!b.createDate) return -1;
            return b.createDate.getTime() - a.createDate.getTime(); // Latest first
        });
        
        // Select the row with the latest date
        const selectedRow = rows[0];
        emailToRowMap.set(email, selectedRow.row);
    });
    
    console.log(`   🔄 Emails with multiple IDs: ${multipleIdCount}`);
    console.log(`   ✅ Processed ${emailToRowMap.size} unique email mappings`);
    
    return emailToRowMap;
}

/**
 * Process Sheet 2: Extract all rows
 */
function processSheet2(workbook) {
    const worksheet = workbook.Sheets[SHEET2_NAME];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    
    console.log(`   📋 Total rows in Sheet 2: ${jsonData.length}`);
    
    return jsonData;
}

/**
 * Match emails between sheets and map the specified columns
 */
function matchAndMapFields(sheet1Data, sheet2Data) {
    const matchedRows = [];
    let matchCount = 0;
    let noMatchCount = 0;
    
    sheet2Data.forEach((row, index) => {
        const email = row['Email'];
        if (!email || typeof email !== 'string') {
            noMatchCount++;
            return; // Skip rows without valid email
        }
        
        const normalizedEmail = email.toLowerCase().trim();
        const sheet1Row = sheet1Data.get(normalizedEmail);
        
        if (!sheet1Row) {
            noMatchCount++;
            return; // No match found, skip this row (per requirements: exclude unmatched)
        }
        
        // Match found - create new row with mapped columns
        const matchedRow = {
            'Email': email, // Keep original email from Sheet 2
        };
        
        // Copy the 6 specified columns from Sheet 1
        COLUMNS_TO_MAP.forEach(columnName => {
            const value = sheet1Row[columnName];
            
            // Format dates if they are date columns
            if ((columnName === 'Billing Start Date' || columnName === 'Billing End Date') && value) {
                matchedRow[columnName] = formatDate(value);
            } else {
                // Copy value as-is (including empty values)
                matchedRow[columnName] = value !== undefined && value !== null ? value : '';
            }
        });
        
        matchedRows.push(matchedRow);
        matchCount++;
    });
    
    console.log(`   ✅ Matched rows: ${matchCount}`);
    console.log(`   ❌ Unmatched rows (excluded): ${noMatchCount}`);
    console.log(`   📊 Match rate: ${((matchCount / sheet2Data.length) * 100).toFixed(2)}%`);
    
    return {
        rows: matchedRows,
        stats: {
            total: sheet2Data.length,
            matched: matchCount,
            unmatched: noMatchCount,
            matchRate: ((matchCount / sheet2Data.length) * 100).toFixed(2)
        }
    };
}

/**
 * Create output Excel file with matched rows
 */
async function createOutputFile(matchedData) {
    // Create new workbook
    const newWorkbook = xlsx.utils.book_new();
    
    // Convert matched data to worksheet
    const worksheet = xlsx.utils.json_to_sheet(matchedData.rows);
    
    // Add worksheet to workbook
    xlsx.utils.book_append_sheet(newWorkbook, worksheet, SHEET2_NAME);
    
    // Write Excel file
    xlsx.writeFile(newWorkbook, OUTPUT_FILE);
    
    console.log(`   ✅ Excel file created: ${OUTPUT_FILE}`);
    console.log(`   📊 Records written: ${matchedData.rows.length}`);
}

/**
 * Display processing results
 */
function displayResults(matchedData) {
    console.log('\n📈 PROCESSING RESULTS');
    console.log('='.repeat(60));
    
    console.log('\n📊 SUMMARY STATISTICS:');
    console.log(`   🔷 Total rows in Sheet 2: ${matchedData.stats.total}`);
    console.log(`   ✅ Successfully matched: ${matchedData.stats.matched}`);
    console.log(`   ❌ Unmatched (excluded): ${matchedData.stats.unmatched}`);
    console.log(`   📈 Match rate: ${matchedData.stats.matchRate}%`);
    
    if (matchedData.rows.length > 0) {
        console.log('\n🔍 SAMPLE MATCHED RECORDS:');
        const sampleRecords = matchedData.rows.slice(0, 3);
        sampleRecords.forEach((record, index) => {
            console.log(`   ${index + 1}. ${record['Email']}`);
            console.log(`      Status: ${record['Status'] || '(empty)'}`);
            console.log(`      Products: ${record['Products'] || '(empty)'}`);
        });
    }
}

// Execute the process
console.log('🚀 STARTING EXCEL EMAIL FIELD MAPPING');
console.log('📋 Process: Email-based field mapping between Stripe Subs and Contact Fields');
console.log('🎯 Goal: Map billing and subscription fields from Sheet 1 to Sheet 2\n');

mapContactFields().catch(error => {
    console.error('❌ Process failed:', error);
    process.exit(1);
});

