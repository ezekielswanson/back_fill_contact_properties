const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const INPUT_FILE = path.join(__dirname, 'input_files/subscriptions_contact_map_fields.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'contacts_with_updated_fields.xlsx');

const SHEET1_NAME = 'Export For Stripe Subs Field Up';
const SHEET2_NAME = 'Export For Contact Field Update';

const COLUMNS_TO_MAP = [
  'Billing Start Date',
  'Billing End Date',
  'Status',
  'Products',
  'Discount',
  'Coupon',
];

function decodeXmlEntities(value) {
  if (!value) return '';
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, num) => String.fromCharCode(parseInt(num, 10)));
}

function readZipEntries(filePath) {
  const buffer = fs.readFileSync(filePath);
  const eocdSignature = 0x06054b50;
  let eocdOffset = -1;

  for (let i = buffer.length - 22; i >= 0; i -= 1) {
    if (buffer.readUInt32LE(i) === eocdSignature) {
      eocdOffset = i;
      break;
    }
  }

  if (eocdOffset === -1) {
    throw new Error(`Unable to locate end of central directory in ${filePath}`);
  }

  const centralDirectoryOffset = buffer.readUInt32LE(eocdOffset + 16);
  const totalEntries = buffer.readUInt16LE(eocdOffset + 10);

  let offset = centralDirectoryOffset;
  const entries = new Map();

  for (let i = 0; i < totalEntries; i += 1) {
    const signature = buffer.readUInt32LE(offset);
    if (signature !== 0x02014b50) {
      throw new Error(`Invalid central directory entry at offset ${offset}`);
    }

    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const uncompressedSize = buffer.readUInt32LE(offset + 24);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const fileName = buffer.slice(offset + 46, offset + 46 + fileNameLength).toString('utf8');

    entries.set(fileName, {
      compressionMethod,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
    });

    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return { buffer, entries };
}

function extractZipEntry(zipData, entryName) {
  const entry = zipData.entries.get(entryName);
  if (!entry) {
    throw new Error(`Entry not found in archive: ${entryName}`);
  }

  const headerOffset = entry.localHeaderOffset;
  const signature = zipData.buffer.readUInt32LE(headerOffset);
  if (signature !== 0x04034b50) {
    throw new Error(`Invalid local file header for ${entryName}`);
  }

  const fileNameLength = zipData.buffer.readUInt16LE(headerOffset + 26);
  const extraLength = zipData.buffer.readUInt16LE(headerOffset + 28);
  const dataStart = headerOffset + 30 + fileNameLength + extraLength;
  const compressedData = zipData.buffer.slice(dataStart, dataStart + entry.compressedSize);

  if (entry.compressionMethod === 0) {
    return compressedData;
  }
  if (entry.compressionMethod === 8) {
    return zlib.inflateRawSync(compressedData);
  }

  throw new Error(`Unsupported compression method ${entry.compressionMethod} for ${entryName}`);
}

function parseWorkbook(zipData) {
  const workbookXml = extractZipEntry(zipData, 'xl/workbook.xml').toString('utf8');
  const sheetRegex = /<sheet[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?>/g;
  const sheets = new Map();
  let match;
  while ((match = sheetRegex.exec(workbookXml)) !== null) {
    sheets.set(match[1], match[2]);
  }

  const relsXml = extractZipEntry(zipData, 'xl/_rels/workbook.xml.rels').toString('utf8');
  const relRegex = /<Relationship[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?>/g;
  const rels = new Map();
  while ((match = relRegex.exec(relsXml)) !== null) {
    rels.set(match[1], match[2]);
  }

  return { sheets, rels };
}

function parseSharedStrings(zipData) {
  let xml;
  try {
    xml = extractZipEntry(zipData, 'xl/sharedStrings.xml').toString('utf8');
  } catch (error) {
    return [];
  }

  const strings = [];
  const siRegex = /<si>([\s\S]*?)<\/si>/g;
  let match;
  while ((match = siRegex.exec(xml)) !== null) {
    const siContent = match[1];
    const textMatches = [...siContent.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)];
    const combined = textMatches.map(textMatch => decodeXmlEntities(textMatch[1])).join('');
    strings.push(combined);
  }

  return strings;
}

function columnLettersToIndex(letters) {
  let index = 0;
  for (const char of letters) {
    index = index * 26 + (char.toUpperCase().charCodeAt(0) - 64);
  }
  return index;
}

function parseCellValue(cellXml, sharedStrings) {
  const typeMatch = cellXml.match(/t="([^"]+)"/);
  const type = typeMatch ? typeMatch[1] : null;
  const valueMatch = cellXml.match(/<v[^>]*>([\s\S]*?)<\/v>/);

  if (type === 'inlineStr') {
    const inlineMatch = cellXml.match(/<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/);
    return inlineMatch ? decodeXmlEntities(inlineMatch[1]) : '';
  }

  if (!valueMatch) {
    return '';
  }

  const rawValue = valueMatch[1];
  if (rawValue === '') {
    return '';
  }
  if (type === 's') {
    return sharedStrings[Number(rawValue)] || '';
  }
  if (type === 'b') {
    return rawValue === '1';
  }
  if (type === 'str') {
    return decodeXmlEntities(rawValue);
  }

  const numeric = Number(rawValue);
  if (!Number.isNaN(numeric)) {
    return Number.isInteger(numeric) ? Number(rawValue) : numeric;
  }

  return decodeXmlEntities(rawValue);
}

function readSheetRows(zipData, sheetName) {
  const { sheets, rels } = parseWorkbook(zipData);
  if (!sheets.has(sheetName)) {
    throw new Error(`Sheet '${sheetName}' not found`);
  }
  const relId = sheets.get(sheetName);
  const target = rels.get(relId);
  if (!target) {
    throw new Error(`Relationship target missing for sheet '${sheetName}'`);
  }

  const sheetPath = `xl/${target}`.replace(/^xl\//, 'xl/');
  const sheetXml = extractZipEntry(zipData, sheetPath).toString('utf8');
  const sharedStrings = parseSharedStrings(zipData);

  const rows = [];
  let headers = null;
  const rowRegex = /<row[^>]*?>([\s\S]*?)<\/row>/g;
  let rowMatch;
  while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
    const rowContent = rowMatch[1];
    const cellRegex = /<c[^>]*?(?:\/>|>\s*[\s\S]*?<\/c>)/g;
    const rowMap = new Map();
    let cellMatch;
    while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
      const cellXml = cellMatch[0];
      const refMatch = cellXml.match(/r="([A-Z]+)\d+"/);
      if (!refMatch) continue;
      const columnIndex = columnLettersToIndex(refMatch[1]);
      rowMap.set(columnIndex, parseCellValue(cellXml, sharedStrings));
    }

    if (!headers) {
      headers = [];
      for (const key of Array.from(rowMap.keys()).sort((a, b) => a - b)) {
        headers.push(rowMap.get(key));
      }
      continue;
    }

    const rowData = {};
    headers.forEach((header, index) => {
      if (!header) return;
      rowData[String(header)] = rowMap.get(index + 1) ?? '';
    });
    if (Object.keys(rowData).length) {
      rows.push(rowData);
    }
  }

  return rows;
}

function listSheetNames(zipData) {
  const { sheets } = parseWorkbook(zipData);
  return Array.from(sheets.keys());
}

function excelSerialDateToJSDate(serial) {
  const serialNumber = Number(serial);
  if (Number.isNaN(serialNumber) || serialNumber <= 0) {
    if (typeof serial === 'string' && !Number.isNaN(Date.parse(serial))) {
      return new Date(serial);
    }
    return null;
  }

  const excelEpochOffsetDays = 25569;
  const daysInMilliseconds = (serialNumber - excelEpochOffsetDays) * 86400000;
  const jsDate = new Date(daysInMilliseconds);

  if (jsDate.getFullYear() < 1900 || jsDate.getFullYear() > 2100) {
    if (typeof serial === 'string' && !Number.isNaN(Date.parse(serial))) {
      return new Date(serial);
    }
    return null;
  }

  return jsDate;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === 'number') return excelSerialDateToJSDate(value);
  if (typeof value === 'string') {
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

function formatDateInTimezone(date, timeZone) {
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  });
  const parts = formatter.formatToParts(date);
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.month}/${values.day}/${values.year} ${values.hour}:${values.minute}:${values.second}`;
}

function formatDate(value, timeZone) {
  if (!value) return '';
  let date = value;
  if (!(value instanceof Date)) {
    const parsed = parseDate(value);
    if (!parsed) return value;
    date = parsed;
  }

  if (timeZone) {
    return formatDateInTimezone(date, timeZone);
  }

  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');

  return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
}

function normalizeValue(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    return Number.isInteger(value) ? String(value) : String(value);
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return '';
    const numeric = Number(trimmed);
    if (!Number.isNaN(numeric)) {
      return Number.isInteger(numeric) ? String(numeric) : String(numeric);
    }
    return trimmed;
  }
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  return String(value);
}

function buildLatestRowMap(sheet1Rows) {
  const grouped = new Map();
  sheet1Rows.forEach(row => {
    const email = row['Stripe Customer Email'];
    if (!email || typeof email !== 'string') return;
    const normalized = email.toLowerCase().trim();
    const createDate = parseDate(row['Most Recent Create Date']);
    if (!grouped.has(normalized)) grouped.set(normalized, []);
    grouped.get(normalized).push({ row, createDate });
  });

  const latestMap = new Map();
  grouped.forEach((rows, email) => {
    rows.sort((a, b) => {
      if (!a.createDate && !b.createDate) return 0;
      if (!a.createDate) return 1;
      if (!b.createDate) return -1;
      return b.createDate.getTime() - a.createDate.getTime();
    });
    latestMap.set(email, rows[0].row);
  });
  return latestMap;
}

function detectTimezone(sheet1Rows, outputRows) {
  const candidates = [
    'America/Chicago',
    'America/New_York',
    'America/Denver',
    'America/Los_Angeles',
    'UTC',
  ];
  const latestMap = buildLatestRowMap(sheet1Rows);

  const samples = [];
  for (const outputRow of outputRows) {
    const email = outputRow.Email;
    if (!email || typeof email !== 'string') continue;
    const sourceRow = latestMap.get(email.toLowerCase().trim());
    if (!sourceRow) continue;
    const billingStart = sourceRow['Billing Start Date'];
    if (typeof billingStart === 'number') {
      samples.push({ sourceRow, outputRow });
    }
    if (samples.length >= 100) break;
  }

  if (!samples.length) return null;

  let bestTimeZone = null;
  let bestMatches = -1;

  for (const timeZone of candidates) {
    let matches = 0;
    let total = 0;
    for (const sample of samples) {
      for (const column of ['Billing Start Date', 'Billing End Date']) {
        const sourceValue = sample.sourceRow[column];
        if (typeof sourceValue !== 'number') continue;
        const expected = formatDate(sourceValue, timeZone);
        const actual = sample.outputRow[column];
        if (expected === actual) matches += 1;
        total += 1;
      }
    }
    if (total === 0) continue;
    if (matches > bestMatches) {
      bestMatches = matches;
      bestTimeZone = timeZone;
    }
  }

  return bestTimeZone;
}

function buildExpectedRows(sheet1Rows, sheet2Rows, timeZone) {
  const latestMap = buildLatestRowMap(sheet1Rows);
  const expected = [];

  sheet2Rows.forEach(row => {
    const email = row.Email;
    if (!email || typeof email !== 'string') return;
    const normalized = email.toLowerCase().trim();
    const sourceRow = latestMap.get(normalized);
    if (!sourceRow) return;

    const expectedRow = { Email: email };
    COLUMNS_TO_MAP.forEach(column => {
      const value = sourceRow[column];
      if (column === 'Billing Start Date' || column === 'Billing End Date') {
        expectedRow[column] = formatDate(value, timeZone);
      } else {
        expectedRow[column] = value !== undefined && value !== null ? value : '';
      }
    });
    expected.push(expectedRow);
  });

  return expected;
}

function compareRows(expectedRows, outputRows) {
  const issues = [];
  if (expectedRows.length !== outputRows.length) {
    issues.push(`Row count mismatch: expected ${expectedRows.length}, found ${outputRows.length}.`);
  }

  expectedRows.forEach((expectedRow, index) => {
    const actualRow = outputRows[index];
    if (!actualRow) return;
    ['Email', ...COLUMNS_TO_MAP].forEach(column => {
      const expectedValue = normalizeValue(expectedRow[column]);
      const actualValue = normalizeValue(actualRow[column]);
      if (expectedValue !== actualValue) {
        issues.push(
          `Row ${index + 2} mismatch for '${column}': expected '${expectedValue}', got '${actualValue}'. Email: ${expectedRow.Email || '(missing)'}`
        );
      }
    });
  });

  return issues;
}

function main() {
  console.log('🔎 VERIFYING CONTACT FIELD MAPPING OUTPUT');
  console.log('='.repeat(60));

  if (!fs.existsSync(INPUT_FILE)) {
    console.error(`❌ Missing input file: ${INPUT_FILE}`);
    process.exit(1);
  }

  if (!fs.existsSync(OUTPUT_FILE)) {
    console.error(`❌ Missing output file: ${OUTPUT_FILE}`);
    process.exit(1);
  }

  const inputZip = readZipEntries(INPUT_FILE);
  const outputZip = readZipEntries(OUTPUT_FILE);

  let sheet1Rows;
  let sheet2Rows;
  try {
    sheet1Rows = readSheetRows(inputZip, SHEET1_NAME);
    sheet2Rows = readSheetRows(inputZip, SHEET2_NAME);
  } catch (error) {
    console.error(`❌ ${error.message}`);
    process.exit(1);
  }

  let outputRows;
  let outputSheetName = SHEET2_NAME;
  try {
    outputRows = readSheetRows(outputZip, outputSheetName);
  } catch (error) {
    const sheetNames = listSheetNames(outputZip);
    if (!sheetNames.length) {
      console.error(`❌ No sheets found in output file: ${OUTPUT_FILE}`);
      process.exit(1);
    }
    outputSheetName = sheetNames[0];
    outputRows = readSheetRows(outputZip, outputSheetName);
  }

  const timeZone = detectTimezone(sheet1Rows, outputRows) || undefined;
  if (timeZone) {
    console.log(`Detected output timezone: ${timeZone}`);
  }

  const expectedRows = buildExpectedRows(sheet1Rows, sheet2Rows, timeZone);
  const issues = compareRows(expectedRows, outputRows);

  console.log(`Input sheet 1 rows: ${sheet1Rows.length}`);
  console.log(`Input sheet 2 rows: ${sheet2Rows.length}`);
  console.log(`Expected output rows: ${expectedRows.length}`);
  console.log(`Actual output rows: ${outputRows.length}`);

  if (issues.length) {
    console.log('\n❌ VERIFICATION FAILED');
    console.log('Issues:');
    issues.slice(0, 20).forEach(issue => console.log(` - ${issue}`));
    if (issues.length > 20) {
      console.log(`...and ${issues.length - 20} more issues.`);
    }
    process.exit(1);
  }

  console.log('\n✅ VERIFICATION PASSED: Output matches expected mapping.');
  process.exit(0);
}

if (require.main === module) {
  main();
}

module.exports = {
  readZipEntries,
  readSheetRows,
};
