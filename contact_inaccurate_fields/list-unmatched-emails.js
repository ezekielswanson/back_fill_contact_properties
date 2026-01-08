const xlsx = require('xlsx');
const path = require('path');

/**
 * List emails from Sheet 2 that don't have matches in Sheet 1
 * These are the emails that were excluded from the output
 */

// Configuration
const INPUT_FILE = path.join(__dirname, 'input_files/subscriptions_contact_map_fields.xlsx');
const SHEET1_NAME = 'Export For Stripe Subs Field Up';
const SHEET2_NAME = 'Export For Contact Field Update';

function listUnmatchedEmails() {
    console.log('🔍 Finding unmatched emails (excluded from output)');
    console.log('='.repeat(80));
    
    try {
        // Load workbook
        const workbook = xlsx.readFile(INPUT_FILE);
        
        // Load Sheet 1 and create email map
        const sheet1 = xlsx.utils.sheet_to_json(workbook.Sheets[SHEET1_NAME]);
        const sheet1Emails = new Set();
        
        sheet1.forEach((row) => {
            const email = row['Stripe Customer Email'];
            if (email && typeof email === 'string') {
                sheet1Emails.add(email.toLowerCase().trim());
            }
        });
        
        console.log(`\n📊 Sheet 1: ${sheet1.length} rows, ${sheet1Emails.size} unique emails`);
        
        // Load Sheet 2
        const sheet2 = xlsx.utils.sheet_to_json(workbook.Sheets[SHEET2_NAME]);
        console.log(`📊 Sheet 2: ${sheet2.length} rows`);
        
        // Find unmatched emails
        const unmatchedEmails = [];
        
        sheet2.forEach((row, index) => {
            const email = row['Email'];
            if (!email || typeof email !== 'string') {
                // Also track rows without email
                unmatchedEmails.push({
                    rowIndex: index + 2, // Excel row number
                    email: '(no email)',
                    reason: 'Missing email field'
                });
                return;
            }
            
            const normalizedEmail = email.toLowerCase().trim();
            if (!sheet1Emails.has(normalizedEmail)) {
                unmatchedEmails.push({
                    rowIndex: index + 2, // Excel row number
                    email: email,
                    normalizedEmail: normalizedEmail
                });
            }
        });
        
        console.log(`\n❌ Unmatched emails: ${unmatchedEmails.length}`);
        console.log(`\n📋 List of unmatched emails:\n`);
        
        // Group by reason
        const missingEmail = unmatchedEmails.filter(e => e.email === '(no email)');
        const noMatch = unmatchedEmails.filter(e => e.email !== '(no email)');
        
        if (noMatch.length > 0) {
            console.log(`Emails not found in Sheet 1 (${noMatch.length}):`);
            console.log('-'.repeat(80));
            noMatch.forEach((item, index) => {
                console.log(`${index + 1}. Row ${item.rowIndex}: ${item.email}`);
            });
        }
        
        if (missingEmail.length > 0) {
            console.log(`\nRows with missing email field (${missingEmail.length}):`);
            console.log('-'.repeat(80));
            missingEmail.forEach((item, index) => {
                console.log(`${index + 1}. Row ${item.rowIndex}: ${item.reason}`);
            });
        }
        
        // Summary
        console.log(`\n📊 Summary:`);
        console.log(`   Total rows in Sheet 2: ${sheet2.length}`);
        console.log(`   Matched emails: ${sheet2.length - unmatchedEmails.length}`);
        console.log(`   Unmatched emails: ${unmatchedEmails.length}`);
        console.log(`     - Not found in Sheet 1: ${noMatch.length}`);
        console.log(`     - Missing email field: ${missingEmail.length}`);
        
        // Export to text file
        const fs = require('fs');
        const outputPath = path.join(__dirname, 'logging_files', 'unmatched-emails.txt');
        const outputDir = path.dirname(outputPath);
        
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }
        
        let outputText = 'UNMATCHED EMAILS FROM SHEET 2\n';
        outputText += '='.repeat(80) + '\n\n';
        outputText += `Total unmatched: ${unmatchedEmails.length}\n\n`;
        
        if (noMatch.length > 0) {
            outputText += `Emails not found in Sheet 1 (${noMatch.length}):\n`;
            outputText += '-'.repeat(80) + '\n';
            noMatch.forEach((item, index) => {
                outputText += `${index + 1}. Row ${item.rowIndex}: ${item.email}\n`;
            });
            outputText += '\n';
        }
        
        if (missingEmail.length > 0) {
            outputText += `Rows with missing email field (${missingEmail.length}):\n`;
            outputText += '-'.repeat(80) + '\n';
            missingEmail.forEach((item, index) => {
                outputText += `${index + 1}. Row ${item.rowIndex}: ${item.reason}\n`;
            });
        }
        
        fs.writeFileSync(outputPath, outputText);
        console.log(`\n📁 Full list saved to: ${outputPath}`);
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        process.exit(1);
    }
}

// Run the script
listUnmatchedEmails();

