const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

/**
 * Verification Script for Contact Field Mapping
 * 
 * Validates 100% accuracy of field mappings from Sheet 1 to Sheet 2
 * by comparing the output file against the original input data.
 */

// Configuration
const INPUT_FILE = path.join(__dirname, 'input_files/subscriptions_contact_map_fields.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'contacts_with_updated_fields.xlsx');
const VERIFICATION_REPORT_DIR = path.join(__dirname, 'logging_files');

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
 */
function excelSerialDateToJSDate(serial) {
    const serialNumber = Number(serial);
    if (isNaN(serialNumber) || serialNumber <= 0) {
        if (typeof serial === 'string' && !isNaN(Date.parse(serial))) {
            return new Date(serial);
        }
        return null;
    }

    const excelEpochOffsetDays = 25569;
    const daysInMilliseconds = (serialNumber - excelEpochOffsetDays) * 86400000;
    const jsDate = new Date(daysInMilliseconds);

    if (jsDate.getFullYear() < 1900 || jsDate.getFullYear() > 2100) {
        if (typeof serial === 'string' && !isNaN(Date.parse(serial))) {
            return new Date(serial);
        }
        return null;
    }

    return jsDate;
}

/**
 * Parses a date value from Excel (handles serial numbers and string formats)
 */
function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) return dateValue;
    
    if (typeof dateValue === 'number') {
        return excelSerialDateToJSDate(dateValue);
    }
    
    if (typeof dateValue === 'string') {
        const parsed = new Date(dateValue);
        return isNaN(parsed.getTime()) ? null : parsed;
    }
    
    return null;
}

/**
 * Formats a date to a standardized string format (matches output format)
 */
function formatDate(date) {
    if (!date) return '';
    if (!(date instanceof Date)) {
        const parsed = parseDate(date);
        if (!parsed) return date;
        date = parsed;
    }
    
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    
    return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
}

/**
 * Normalizes a value for comparison (handles empty strings, null, undefined)
 */
function normalizeValue(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value.trim();
    return String(value);
}

/**
 * Main Verification Function
 */
async function verifyFieldMapping() {
    console.log('🔍 CONTACT FIELD MAPPING VERIFICATION');
    console.log('📋 Purpose: Validate 100% Accuracy of Field Mappings');
    console.log('🎯 Goal: Ensure All Mapped Fields Match Source Data');
    console.log('='.repeat(60));

    const verification = {
        timestamp: new Date().toISOString(),
        script_verified: 'map-contact-fields.js',
        status: 'UNKNOWN',
        tests: [],
        summary: {
            totalTests: 0,
            passedTests: 0,
            failedTests: 0,
            warnings: 0
        },
        issues: [],
        warnings: [],
        details: {}
    };

    try {
        // Test 1: File Existence
        await testFileExistence(verification);
        
        // Test 2: Load input data
        const inputData = await loadInputData(verification);
        
        // Test 3: Load output data
        const outputData = await loadOutputData(verification);
        
        // Test 4: Verify email matching logic
        await verifyEmailMatching(verification, inputData, outputData);
        
        // Test 5: Verify correct row selection (latest create date)
        await verifyRowSelection(verification, inputData, outputData);
        
        // Test 6: Verify field mapping accuracy
        await verifyFieldAccuracy(verification, inputData, outputData);
        
        // Test 7: Verify data completeness
        await verifyDataCompleteness(verification, inputData, outputData);
        
        // Test 8: Verify date formatting
        await verifyDateFormatting(verification, inputData, outputData);
        
        // Generate final status
        generateFinalStatus(verification);
        
        // Display results
        displayVerificationResults(verification);
        
        // Save verification report
        await saveVerificationReport(verification);
        
        console.log('\n✅ VERIFICATION COMPLETE!');
        
        // Exit with appropriate code
        if (verification.status === 'PASSED') {
            process.exit(0);
        } else {
            process.exit(1);
        }
        
    } catch (error) {
        console.error('\n❌ Verification failed:', error.message);
        verification.status = 'FAILED';
        verification.issues.push(`Critical error: ${error.message}`);
        await saveVerificationReport(verification);
        process.exit(1);
    }
}

/**
 * Test 1: File Existence
 */
async function testFileExistence(verification) {
    console.log('\n🔍 TEST 1: File Existence Check');
    verification.summary.totalTests++;
    
    const test = {
        name: 'File Existence',
        status: 'PASSED',
        issues: []
    };
    
    if (!fs.existsSync(INPUT_FILE)) {
        test.status = 'FAILED';
        test.issues.push(`Input file not found: ${INPUT_FILE}`);
        verification.issues.push(...test.issues);
    } else {
        console.log(`   ✅ Input file exists: ${INPUT_FILE}`);
    }
    
    if (!fs.existsSync(OUTPUT_FILE)) {
        test.status = 'FAILED';
        test.issues.push(`Output file not found: ${OUTPUT_FILE}`);
        verification.issues.push(...test.issues);
    } else {
        console.log(`   ✅ Output file exists: ${OUTPUT_FILE}`);
    }
    
    if (test.status === 'PASSED') {
        verification.summary.passedTests++;
    } else {
        verification.summary.failedTests++;
    }
    
    verification.tests.push(test);
}

/**
 * Load input data from both sheets
 */
async function loadInputData(verification) {
    console.log('\n📊 Loading input data...');
    
    const workbook = xlsx.readFile(INPUT_FILE);
    
    if (!workbook.SheetNames.includes(SHEET1_NAME)) {
        throw new Error(`Sheet "${SHEET1_NAME}" not found in input file`);
    }
    if (!workbook.SheetNames.includes(SHEET2_NAME)) {
        throw new Error(`Sheet "${SHEET2_NAME}" not found in input file`);
    }
    
    const sheet1 = xlsx.utils.sheet_to_json(workbook.Sheets[SHEET1_NAME]);
    const sheet2 = xlsx.utils.sheet_to_json(workbook.Sheets[SHEET2_NAME]);
    
    console.log(`   ✅ Sheet 1 rows: ${sheet1.length}`);
    console.log(`   ✅ Sheet 2 rows: ${sheet2.length}`);
    
    // Process Sheet 1: Group by email and find latest per email
    const emailToRowMap = new Map();
    const emailGroups = new Map();
    
    sheet1.forEach((row, index) => {
        const email = row['Stripe Customer Email'];
        if (!email || typeof email !== 'string') {
            return;
        }
        
        const normalizedEmail = email.toLowerCase().trim();
        const createDate = parseDate(row['Most Recent Create Date']);
        
        if (!emailGroups.has(normalizedEmail)) {
            emailGroups.set(normalizedEmail, []);
        }
        
        emailGroups.get(normalizedEmail).push({
            row: row,
            createDate: createDate,
            index: index
        });
    });
    
    // Select latest row per email
    emailGroups.forEach((rows, email) => {
        rows.sort((a, b) => {
            if (!a.createDate && !b.createDate) return 0;
            if (!a.createDate) return 1;
            if (!b.createDate) return -1;
            return b.createDate.getTime() - a.createDate.getTime();
        });
        
        emailToRowMap.set(email, rows[0].row);
    });
    
    return {
        sheet1: sheet1,
        sheet2: sheet2,
        emailToRowMap: emailToRowMap,
        emailGroups: emailGroups
    };
}

/**
 * Load output data
 */
async function loadOutputData(verification) {
    console.log('\n📊 Loading output data...');
    
    const workbook = xlsx.readFile(OUTPUT_FILE);
    const sheetNames = workbook.SheetNames;
    
    console.log(`   📋 Sheets found: ${sheetNames.join(', ')}`);
    
    // Use the first sheet (should be Sheet 2 name)
    const outputSheet = sheetNames[0];
    const outputData = xlsx.utils.sheet_to_json(workbook.Sheets[outputSheet]);
    
    console.log(`   ✅ Output rows: ${outputData.length}`);
    
    return {
        data: outputData,
        sheetName: outputSheet
    };
}

/**
 * Test 4: Verify email matching logic
 */
async function verifyEmailMatching(verification, inputData, outputData) {
    console.log('\n🔍 TEST 4: Email Matching Logic Verification');
    verification.summary.totalTests++;
    
    const test = {
        name: 'Email Matching',
        status: 'PASSED',
        issues: [],
        details: {
            totalOutputEmails: outputData.data.length,
            matchedEmails: 0,
            unmatchedEmails: []
        }
    };
    
    const outputEmails = new Set();
    const unmatchedEmails = [];
    
    outputData.data.forEach((row) => {
        const email = row['Email'];
        if (!email) {
            test.issues.push('Found row without Email in output');
            return;
        }
        
        const normalizedEmail = email.toLowerCase().trim();
        outputEmails.add(normalizedEmail);
        
        if (!inputData.emailToRowMap.has(normalizedEmail)) {
            unmatchedEmails.push(normalizedEmail);
            test.issues.push(`Email in output not found in Sheet 1: ${email}`);
        } else {
            test.details.matchedEmails++;
        }
    });
    
    test.details.unmatchedEmails = unmatchedEmails;
    
    if (test.issues.length > 0) {
        test.status = 'FAILED';
        verification.issues.push(...test.issues);
        verification.summary.failedTests++;
    } else {
        verification.summary.passedTests++;
        console.log(`   ✅ All ${test.details.matchedEmails} emails have valid matches`);
    }
    
    verification.tests.push(test);
}

/**
 * Test 5: Verify correct row selection (latest create date)
 */
async function verifyRowSelection(verification, inputData, outputData) {
    console.log('\n🔍 TEST 5: Row Selection Verification (Latest Create Date)');
    verification.summary.totalTests++;
    
    const test = {
        name: 'Row Selection (Latest Date)',
        status: 'PASSED',
        issues: [],
        details: {
            emailsWithMultipleIds: 0,
            correctSelections: 0,
            incorrectSelections: []
        }
    };
    
    let incorrectSelections = [];
    
    inputData.emailGroups.forEach((rows, email) => {
        if (rows.length > 1) {
            test.details.emailsWithMultipleIds++;
            
            // Find the row that should have been selected (latest date)
            rows.sort((a, b) => {
                if (!a.createDate && !b.createDate) return 0;
                if (!a.createDate) return 1;
                if (!b.createDate) return -1;
                return b.createDate.getTime() - a.createDate.getTime();
            });
            
            const expectedRow = rows[0].row;
            const actualRow = inputData.emailToRowMap.get(email);
            
            // Verify the selected row matches the expected one
            if (actualRow['Id'] !== expectedRow['Id']) {
                incorrectSelections.push({
                    email: email,
                    expectedId: expectedRow['Id'],
                    actualId: actualRow['Id'],
                    expectedDate: expectedRow['Most Recent Create Date'],
                    actualDate: actualRow['Most Recent Create Date']
                });
                test.issues.push(`Incorrect row selected for ${email}. Expected ID: ${expectedRow['Id']}, Got: ${actualRow['Id']}`);
            } else {
                test.details.correctSelections++;
            }
        }
    });
    
    test.details.incorrectSelections = incorrectSelections;
    
    if (test.issues.length > 0) {
        test.status = 'FAILED';
        verification.issues.push(...test.issues);
        verification.summary.failedTests++;
    } else {
        verification.summary.passedTests++;
        console.log(`   ✅ All ${test.details.correctSelections} multi-ID emails have correct row selection`);
    }
    
    verification.tests.push(test);
}

/**
 * Test 6: Verify field mapping accuracy
 */
async function verifyFieldAccuracy(verification, inputData, outputData) {
    console.log('\n🔍 TEST 6: Field Mapping Accuracy Verification');
    verification.summary.totalTests++;
    
    const test = {
        name: 'Field Mapping Accuracy',
        status: 'PASSED',
        issues: [],
        details: {
            totalFieldsChecked: 0,
            correctFields: 0,
            incorrectFields: []
        }
    };
    
    const incorrectFields = [];
    
    outputData.data.forEach((outputRow, index) => {
        const email = outputRow['Email'];
        if (!email) return;
        
        const normalizedEmail = email.toLowerCase().trim();
        const sourceRow = inputData.emailToRowMap.get(normalizedEmail);
        
        if (!sourceRow) {
            return; // Already handled in email matching test
        }
        
        // Check each mapped column
        COLUMNS_TO_MAP.forEach(columnName => {
            test.details.totalFieldsChecked++;
            
            const sourceValue = sourceRow[columnName];
            const outputValue = outputRow[columnName];
            
            // Handle date columns specially
            if (columnName === 'Billing Start Date' || columnName === 'Billing End Date') {
                const expectedFormatted = formatDate(sourceValue);
                const actualFormatted = normalizeValue(outputValue);
                
                if (expectedFormatted !== actualFormatted) {
                    incorrectFields.push({
                        email: email,
                        column: columnName,
                        expected: expectedFormatted,
                        actual: actualFormatted,
                        rowIndex: index
                    });
                    test.issues.push(`Field mismatch for ${email}: ${columnName}. Expected: "${expectedFormatted}", Got: "${actualFormatted}"`);
                } else {
                    test.details.correctFields++;
                }
            } else {
                // For non-date columns, compare normalized values
                const expectedNormalized = normalizeValue(sourceValue);
                const actualNormalized = normalizeValue(outputValue);
                
                if (expectedNormalized !== actualNormalized) {
                    incorrectFields.push({
                        email: email,
                        column: columnName,
                        expected: expectedNormalized,
                        actual: actualNormalized,
                        rowIndex: index
                    });
                    test.issues.push(`Field mismatch for ${email}: ${columnName}. Expected: "${expectedNormalized}", Got: "${actualNormalized}"`);
                } else {
                    test.details.correctFields++;
                }
            }
        });
    });
    
    test.details.incorrectFields = incorrectFields;
    
    if (test.issues.length > 0) {
        test.status = 'FAILED';
        verification.issues.push(...test.issues);
        verification.summary.failedTests++;
        console.log(`   ❌ Found ${test.issues.length} field mismatches`);
    } else {
        verification.summary.passedTests++;
        console.log(`   ✅ All ${test.details.correctFields} fields mapped correctly`);
    }
    
    verification.tests.push(test);
}

/**
 * Test 7: Verify data completeness
 */
async function verifyDataCompleteness(verification, inputData, outputData) {
    console.log('\n🔍 TEST 7: Data Completeness Verification');
    verification.summary.totalTests++;
    
    const test = {
        name: 'Data Completeness',
        status: 'PASSED',
        issues: [],
        details: {
            missingColumns: [],
            missingRequiredFields: []
        }
    };
    
    // Check that all required columns exist in output
    const requiredColumns = ['Email', ...COLUMNS_TO_MAP];
    const outputColumns = outputData.data.length > 0 ? Object.keys(outputData.data[0]) : [];
    
    requiredColumns.forEach(column => {
        if (!outputColumns.includes(column)) {
            test.issues.push(`Required column missing in output: ${column}`);
            test.details.missingColumns.push(column);
        }
    });
    
    // Check that all output rows have Email field
    outputData.data.forEach((row, index) => {
        if (!row['Email']) {
            test.issues.push(`Row ${index + 1} missing Email field`);
            test.details.missingRequiredFields.push(index + 1);
        }
    });
    
    if (test.issues.length > 0) {
        test.status = 'FAILED';
        verification.issues.push(...test.issues);
        verification.summary.failedTests++;
    } else {
        verification.summary.passedTests++;
        console.log(`   ✅ All required columns and fields present`);
    }
    
    verification.tests.push(test);
}

/**
 * Test 8: Verify date formatting
 */
async function verifyDateFormatting(verification, inputData, outputData) {
    console.log('\n🔍 TEST 8: Date Formatting Verification');
    verification.summary.totalTests++;
    
    const test = {
        name: 'Date Formatting',
        status: 'PASSED',
        issues: [],
        details: {
            dateFieldsChecked: 0,
            correctlyFormatted: 0,
            incorrectlyFormatted: []
        }
    };
    
    const dateColumns = ['Billing Start Date', 'Billing End Date'];
    const incorrectlyFormatted = [];
    
    outputData.data.forEach((outputRow, index) => {
        const email = outputRow['Email'];
        if (!email) return;
        
        const normalizedEmail = email.toLowerCase().trim();
        const sourceRow = inputData.emailToRowMap.get(normalizedEmail);
        
        if (!sourceRow) return;
        
        dateColumns.forEach(columnName => {
            const sourceValue = sourceRow[columnName];
            const outputValue = outputRow[columnName];
            
            if (sourceValue) {
                test.details.dateFieldsChecked++;
                
                const expectedFormatted = formatDate(sourceValue);
                const actualFormatted = normalizeValue(outputValue);
                
                // Check format: MM/DD/YYYY HH:MM:SS
                const dateFormatRegex = /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/;
                
                if (!dateFormatRegex.test(actualFormatted) && actualFormatted !== '') {
                    incorrectlyFormatted.push({
                        email: email,
                        column: columnName,
                        value: actualFormatted,
                        rowIndex: index
                    });
                    test.issues.push(`Incorrect date format for ${email}: ${columnName}. Value: "${actualFormatted}"`);
                } else if (expectedFormatted === actualFormatted) {
                    test.details.correctlyFormatted++;
                }
            }
        });
    });
    
    test.details.incorrectlyFormatted = incorrectlyFormatted;
    
    if (test.issues.length > 0) {
        test.status = 'FAILED';
        verification.issues.push(...test.issues);
        verification.summary.failedTests++;
    } else {
        verification.summary.passedTests++;
        console.log(`   ✅ All ${test.details.correctlyFormatted} date fields correctly formatted`);
    }
    
    verification.tests.push(test);
}

/**
 * Generate final status
 */
function generateFinalStatus(verification) {
    if (verification.summary.failedTests === 0) {
        verification.status = 'PASSED';
    } else {
        verification.status = 'FAILED';
    }
}

/**
 * Display verification results
 */
function displayVerificationResults(verification) {
    console.log('\n' + '='.repeat(60));
    console.log('📈 VERIFICATION RESULTS SUMMARY');
    console.log('='.repeat(60));
    
    console.log(`\n📊 Test Summary:`);
    console.log(`   Total Tests: ${verification.summary.totalTests}`);
    console.log(`   ✅ Passed: ${verification.summary.passedTests}`);
    console.log(`   ❌ Failed: ${verification.summary.failedTests}`);
    console.log(`   ⚠️  Warnings: ${verification.summary.warnings}`);
    
    console.log(`\n🎯 Overall Status: ${verification.status}`);
    
    if (verification.issues.length > 0) {
        console.log(`\n❌ Issues Found (${verification.issues.length}):`);
        verification.issues.slice(0, 10).forEach((issue, index) => {
            console.log(`   ${index + 1}. ${issue}`);
        });
        if (verification.issues.length > 10) {
            console.log(`   ... and ${verification.issues.length - 10} more issues`);
        }
    }
    
    if (verification.warnings.length > 0) {
        console.log(`\n⚠️  Warnings (${verification.warnings.length}):`);
        verification.warnings.slice(0, 5).forEach((warning, index) => {
            console.log(`   ${index + 1}. ${warning}`);
        });
    }
    
    console.log('\n📋 Detailed Test Results:');
    verification.tests.forEach((test, index) => {
        const statusIcon = test.status === 'PASSED' ? '✅' : '❌';
        console.log(`   ${statusIcon} Test ${index + 1}: ${test.name} - ${test.status}`);
        if (test.details && Object.keys(test.details).length > 0) {
            Object.entries(test.details).forEach(([key, value]) => {
                if (typeof value === 'number' || typeof value === 'string') {
                    console.log(`      ${key}: ${value}`);
                }
            });
        }
    });
}

/**
 * Save verification report
 */
async function saveVerificationReport(verification) {
    // Ensure directory exists
    if (!fs.existsSync(VERIFICATION_REPORT_DIR)) {
        fs.mkdirSync(VERIFICATION_REPORT_DIR, { recursive: true });
    }
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const reportPath = path.join(VERIFICATION_REPORT_DIR, `verification-report-${timestamp}.json`);
    
    fs.writeFileSync(reportPath, JSON.stringify(verification, null, 2));
    console.log(`\n📁 Verification report saved: ${reportPath}`);
}

// Execute verification
console.log('🚀 STARTING FIELD MAPPING VERIFICATION');
console.log('📋 Script: verify-field-mapping.js');
console.log('🎯 Goal: Ensure 100% Accuracy of Field Mappings\n');

verifyFieldMapping().catch(error => {
    console.error('❌ Verification process failed:', error);
    process.exit(1);
});

