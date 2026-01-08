const xlsx = require('xlsx');
const path = require('path');

/**
 * List emails with multiple subscription IDs
 * Shows which emails have multiple IDs and which one was selected (latest create date)
 */

// Configuration
const INPUT_FILE = path.join(__dirname, 'input_files/subscriptions_contact_map_fields.xlsx');
const SHEET1_NAME = 'Export For Stripe Subs Field Up';

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
 * Parses a date value from Excel
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
 * Formats a date for display
 */
function formatDateForDisplay(date) {
    if (!date) return 'N/A';
    if (!(date instanceof Date)) {
        const parsed = parseDate(date);
        if (!parsed) return String(date);
        date = parsed;
    }
    
    return date.toLocaleString('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
}

function listMultipleIds() {
    console.log('🔍 Analyzing emails with multiple subscription IDs');
    console.log('='.repeat(80));
    
    try {
        // Load workbook
        const workbook = xlsx.readFile(INPUT_FILE);
        const worksheet = workbook.Sheets[SHEET1_NAME];
        const jsonData = xlsx.utils.sheet_to_json(worksheet);
        
        console.log(`\n📊 Total rows in Sheet 1: ${jsonData.length}`);
        
        // Group rows by email
        const emailGroups = new Map();
        
        jsonData.forEach((row, index) => {
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
                rowIndex: index + 2, // +2 for Excel row number (1-indexed + header)
                id: row['Id'],
                createDate: createDate,
                createDateRaw: row['Most Recent Create Date'],
                status: row['Status'],
                products: row['Products']
            });
        });
        
        // Find emails with multiple IDs
        const multipleIdEmails = [];
        
        emailGroups.forEach((rows, email) => {
            if (rows.length > 1) {
                // Sort by create date (latest first)
                rows.sort((a, b) => {
                    if (!a.createDate && !b.createDate) return 0;
                    if (!a.createDate) return 1;
                    if (!b.createDate) return -1;
                    return b.createDate.getTime() - a.createDate.getTime();
                });
                
                multipleIdEmails.push({
                    email: email,
                    count: rows.length,
                    rows: rows,
                    selected: rows[0] // The one with latest date
                });
            }
        });
        
        console.log(`\n📧 Emails with multiple IDs: ${multipleIdEmails.length}`);
        console.log(`\n📋 Sample Examples (showing first 20):\n`);
        
        // Display examples
        multipleIdEmails.slice(0, 20).forEach((emailData, index) => {
            console.log(`${index + 1}. ${emailData.email}`);
            console.log(`   Total subscription IDs: ${emailData.count}`);
            console.log(`   ✅ SELECTED (Latest Create Date):`);
            console.log(`      - Row: ${emailData.selected.rowIndex}`);
            console.log(`      - ID: ${emailData.selected.id}`);
            console.log(`      - Create Date: ${formatDateForDisplay(emailData.selected.createDate)} (${emailData.selected.createDateRaw})`);
            console.log(`      - Status: ${emailData.selected.status || 'N/A'}`);
            console.log(`      - Products: ${emailData.selected.products || 'N/A'}`);
            
            if (emailData.rows.length > 1) {
                console.log(`   ⚠️  Other IDs (not selected):`);
                emailData.rows.slice(1).forEach((row, idx) => {
                    console.log(`      ${idx + 1}. Row ${row.rowIndex} - ID: ${row.id} - Create Date: ${formatDateForDisplay(row.createDate)} (${row.createDateRaw})`);
                });
            }
            console.log('');
        });
        
        if (multipleIdEmails.length > 20) {
            console.log(`\n... and ${multipleIdEmails.length - 20} more emails with multiple IDs\n`);
        }
        
        // Summary statistics
        console.log('\n📊 Summary Statistics:');
        const idCounts = {};
        multipleIdEmails.forEach(emailData => {
            const count = emailData.count;
            idCounts[count] = (idCounts[count] || 0) + 1;
        });
        
        Object.keys(idCounts).sort((a, b) => Number(a) - Number(b)).forEach(count => {
            console.log(`   ${count} subscription IDs: ${idCounts[count]} emails`);
        });
        
        // Show the example from the README if it exists
        const readmeExample = multipleIdEmails.find(e => e.email === 'sratner@umich.edu');
        if (readmeExample) {
            console.log('\n📝 Example from README (sratner@umich.edu):');
            console.log(`   Total IDs: ${readmeExample.count}`);
            console.log(`   ✅ Selected ID: ${readmeExample.selected.id}`);
            console.log(`   Create Date: ${formatDateForDisplay(readmeExample.selected.createDate)}`);
        }
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        process.exit(1);
    }
}

// Run the script
listMultipleIds();

