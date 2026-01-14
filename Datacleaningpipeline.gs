/**
 * =============================================================================
 * DATA CLEANING PIPELINE V3 - HVAC VERSION - Production-Ready Google Sheets Apps Script
 * =============================================================================
 * 
 * HVAC VERSION:
 * - Integrated with HVAC Categorization System
 * - Menu calls: categorizeHVACCompanies() and uploadCheckedEntriesHVAC()
 * - For Home Builders version, use DataCleaningPipeline_V3.gs
 * 
 * VERSION 3 UPDATES (Critical Fix):
 * - Row-level duplicate detection: Uses COMPOSITE KEY (Name + Company + Website)
 * - Column-level deduplication: ONLY applies to Email and Phone columns
 * - Organization, Website, City, Description: NO duplicate removal (preserved as-is)
 * 
 * DUPLICATE DETECTION LOGIC:
 * A row is considered a duplicate ONLY if ALL THREE match:
 * 1. Contact Full Name
 * 2. Company Name - Cleaned
 * 3. Website
 * 
 * Example - NOT duplicates (same name, different companies):
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Whitestone Builders | whitestone.com
 * 
 * Example - ARE duplicates (all three fields match):
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Aspire Fine Homes | aspire.com  ← DUPLICATE ROW
 * 
 * IMPORTANT: Source sheet must have "Company Name - Cleaned" column
 *            Output sheet will show this as "Organization"
 * 
 * Previous versions (V1/V2) incorrectly removed duplicates within the same row
 * across ALL columns. V3 fixes this by:
 * 1. Removing duplicate contacts using composite key (Name + Company + Website)
 * 2. Only deduplicating emails/phones within each row
 * 
 * @author: Claude
 * @version: 3.0 - HVAC Categorization Version
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const CONFIG = {
  // Sheet names
  OUTPUT_SHEET_NAME: 'Cleaned Data V3',
  
  // Columns to process (by header name)
  COLUMNS: {
    FULL_NAME: 'Contact Full Name',
    ORGANIZATION: 'Company Name - Cleaned',  // Source column name
    WEBSITE: 'Website',
    PRIMARY_EMAIL: 'Primary Email',
    EMAIL_1: 'Email 1',
    EMAIL_2: 'Email 2',
    PERSONAL_EMAIL: 'Personal Email',
    CONTACT_PHONE_1: 'Contact Phone 1',
    COMPANY_PHONE_1: 'Company Phone 1',
    COMPANY_PHONE_2: 'Company Phone 2',
    CONTACT_MOBILE: 'Contact Mobile Phone',
    COMPANY_CITY: 'Company City',
    COMPANY_DESC: 'Company Description'
  },
  
  // Performance settings
  BATCH_SIZE: 1000
};

// =============================================================================
// MAIN EXECUTION
// =============================================================================

/**
 * Main entry point for the data cleaning pipeline V3
 */
function runDataCleaningPipelineV3() {
  try {
    const startTime = new Date();
    Logger.log('=== Starting Data Cleaning Pipeline V3 ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getActiveSheet();
    
    // Validate source sheet
    if (!validateSourceSheet(sourceSheet)) {
      throw new Error('Source sheet validation failed. Check required columns.');
    }
    
    // Step 1: Load and parse data
    Logger.log('Step 1: Loading data...');
    const rawData = loadSheetData(sourceSheet);
    const originalRowCount = rawData.data.length;
    
    // Step 2: V3 UPDATED - Remove duplicate ROWS using composite key
    Logger.log('Step 2: Removing duplicate rows (by Name + Company + Website)...');
    const uniqueRowsData = removeDuplicateRows(rawData);
    const removedRows = originalRowCount - uniqueRowsData.data.length;
    Logger.log(`Removed ${removedRows} duplicate rows`);
    
    // Step 3: V3 UPDATED - Clean duplicates ONLY within Email and Phone columns
    Logger.log('Step 3: Cleaning duplicates within Email/Phone columns...');
    const cleanedData = cleanEmailPhoneDuplicates(uniqueRowsData);
    
    // Step 4: Fill Primary Email from Email 1 if empty
    Logger.log('Step 4: Applying Primary Email fill rule...');
    const emailFilledData = fillPrimaryEmail(cleanedData);
    
    // Step 5: Deduplicate and normalize emails
    Logger.log('Step 5: Deduplicating emails...');
    const emailCleanedData = deduplicateEmails(emailFilledData);
    
    // Step 6: Deduplicate and normalize phones with priority
    Logger.log('Step 6: Deduplicating phones (Mobile priority)...');
    const phoneCleanedData = deduplicatePhonesWithPriority(emailCleanedData);
    
    // Step 7: Parse names
    Logger.log('Step 7: Parsing names...');
    const finalData = parseNames(phoneCleanedData);
    
    // Step 8: Create output sheet
    Logger.log('Step 8: Creating output sheet...');
    const outputSheet = createOutputSheet(ss, sourceSheet);
    
    // Step 9: Write data to output
    Logger.log('Step 9: Writing cleaned data...');
    writeOutputData(outputSheet, finalData, sourceSheet.getName());
    
    // Step 10: Add validation formula
    Logger.log('Step 10: Adding validation formula...');
    addValidationFormula(outputSheet, sourceSheet.getName());
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`=== Pipeline V3 Complete in ${duration}s ===`);
    Logger.log(`Original rows: ${originalRowCount}`);
    Logger.log(`Removed duplicate rows: ${removedRows}`);
    Logger.log(`Final rows: ${finalData.data.length}`);
    
    SpreadsheetApp.getUi().alert(
      'Success! (Version 3)',
      `Data cleaning complete!\n\n` +
      `Original rows: ${originalRowCount}\n` +
      `Duplicate rows removed: ${removedRows}\n` +
      `Final rows: ${finalData.data.length}\n` +
      `Duration: ${duration.toFixed(2)}s\n` +
      `Output: "${CONFIG.OUTPUT_SHEET_NAME}" tab\n\n` +
      `V3 Updates Applied:\n` +
      `✓ Duplicate rows removed (Name + Company + Website)\n` +
      `✓ Organization/Website preserved\n` +
      `✓ Only Emails/Phones deduplicated`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Pipeline failed: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

// =============================================================================
// STEP 1: DATA LOADING
// =============================================================================

/**
 * Validates that the source sheet contains all required columns
 */
function validateSourceSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const requiredColumns = Object.values(CONFIG.COLUMNS);
  
  const missingColumns = requiredColumns.filter(col => !headers.includes(col));
  
  if (missingColumns.length > 0) {
    Logger.log('Missing columns: ' + missingColumns.join(', '));
    return false;
  }
  
  return true;
}

/**
 * Loads data from sheet and creates a structured object
 */
function loadSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error('No data rows found in sheet');
  }
  
  // Get all data including headers
  const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  // Create column index map
  const columnMap = {};
  headers.forEach((header, index) => {
    columnMap[header] = index;
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: dataRows
  };
}

// =============================================================================
// STEP 2: ROW-LEVEL DUPLICATE REMOVAL (V3 NEW)
// =============================================================================

/**
 * V3 UPDATED: Removes duplicate ROWS based on composite key:
 * - Contact Full Name + Company Name - Cleaned + Website
 * 
 * A row is considered a duplicate ONLY if ALL THREE match:
 * 1. Same Contact Full Name
 * 2. Same Company Name - Cleaned
 * 3. Same Website
 * 
 * Example - These are NOT duplicates:
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Whitestone Builders | whitestone.com
 * 
 * Example - These ARE duplicates:
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Aspire Fine Homes | aspire.com  ← DUPLICATE
 */
function removeDuplicateRows(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const fullNameIdx = columnMap[CONFIG.COLUMNS.FULL_NAME];
  const companyNameIdx = columnMap[CONFIG.COLUMNS.ORGANIZATION];
  const websiteIdx = columnMap[CONFIG.COLUMNS.WEBSITE];
  
  const seenCombinations = new Set();
  const uniqueRows = [];
  
  data.forEach((row, index) => {
    const fullName = row[fullNameIdx];
    const companyName = row[companyNameIdx];
    const website = row[websiteIdx];
    
    // Create composite key from all three fields
    // Normalize: trim and lowercase for comparison
    const nameKey = fullName ? fullName.toString().trim().toLowerCase() : '';
    const companyKey = companyName ? companyName.toString().trim().toLowerCase() : '';
    const websiteKey = website ? website.toString().trim().toLowerCase() : '';
    
    // Composite key: name|company|website
    const compositeKey = `${nameKey}|${companyKey}|${websiteKey}`;
    
    if (seenCombinations.has(compositeKey)) {
      // Duplicate found - skip this entire row
      Logger.log(`Duplicate row removed (row ${index + 2}): Name="${fullName}", Company="${companyName}", Website="${website}"`);
    } else {
      // First occurrence - keep this row
      seenCombinations.add(compositeKey);
      uniqueRows.push(row);
    }
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: uniqueRows
  };
}

// =============================================================================
// STEP 3: COLUMN-LEVEL DUPLICATE CLEANING (V3 UPDATED)
// =============================================================================

/**
 * V3 UPDATED: Removes duplicate values ONLY within Email and Phone columns
 * Does NOT touch: Organization, Website, Company Description, Company City
 * 
 * This fixes the V2 bug where "John Galli" in Contact Full Name and Organization
 * would incorrectly remove Organization
 */
function cleanEmailPhoneDuplicates(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  // V3: ONLY clean duplicates in Email and Phone columns
  const columnsToClean = [
    CONFIG.COLUMNS.PRIMARY_EMAIL,
    CONFIG.COLUMNS.EMAIL_1,
    CONFIG.COLUMNS.EMAIL_2,
    CONFIG.COLUMNS.PERSONAL_EMAIL,
    CONFIG.COLUMNS.CONTACT_PHONE_1,
    CONFIG.COLUMNS.COMPANY_PHONE_1,
    CONFIG.COLUMNS.COMPANY_PHONE_2,
    CONFIG.COLUMNS.CONTACT_MOBILE
  ];
  
  const cleanedData = data.map(row => {
    const newRow = [...row];
    const seenValues = new Set();
    
    // Get indices for columns to clean (only emails and phones)
    const indicesToClean = columnsToClean
      .map(col => columnMap[col])
      .filter(idx => idx !== undefined);
    
    indicesToClean.forEach(colIdx => {
      const value = row[colIdx];
      
      // Skip empty values
      if (!value || value.toString().trim() === '') {
        return;
      }
      
      // Normalize for comparison: trim and lowercase
      const normalizedValue = value.toString().trim().toLowerCase();
      
      if (seenValues.has(normalizedValue)) {
        // Duplicate found within this row - clear it
        newRow[colIdx] = '';
      } else {
        // First occurrence - keep it and add to set
        seenValues.add(normalizedValue);
        // Keep original value but trimmed
        newRow[colIdx] = value.toString().trim();
      }
    });
    
    return newRow;
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: cleanedData
  };
}

// =============================================================================
// STEP 4: PRIMARY EMAIL FILL RULE
// =============================================================================

/**
 * If Primary Email is empty and Email 1 has value, move Email 1 to Primary Email
 */
function fillPrimaryEmail(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const primaryIdx = columnMap[CONFIG.COLUMNS.PRIMARY_EMAIL];
  const email1Idx = columnMap[CONFIG.COLUMNS.EMAIL_1];
  
  const filledData = data.map(row => {
    const newRow = [...row];
    
    const primary = row[primaryIdx] ? row[primaryIdx].toString().trim() : '';
    const email1 = row[email1Idx] ? row[email1Idx].toString().trim() : '';
    
    // If Primary Email is empty AND Email 1 has a value
    if (!primary && email1) {
      // Move Email 1 → Primary Email
      newRow[primaryIdx] = email1;
      // Clear Email 1
      newRow[email1Idx] = '';
    }
    
    return newRow;
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: filledData
  };
}

// =============================================================================
// STEP 5: EMAIL DEDUPLICATION
// =============================================================================

/**
 * Deduplicates emails according to business rules and normalizes them
 */
function deduplicateEmails(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const primaryIdx = columnMap[CONFIG.COLUMNS.PRIMARY_EMAIL];
  const email1Idx = columnMap[CONFIG.COLUMNS.EMAIL_1];
  const email2Idx = columnMap[CONFIG.COLUMNS.EMAIL_2];
  const personalIdx = columnMap[CONFIG.COLUMNS.PERSONAL_EMAIL];
  
  const cleanedData = data.map(row => {
    const newRow = [...row];
    
    // Normalize and get emails
    let primary = normalizeEmail(row[primaryIdx]);
    let email1 = normalizeEmail(row[email1Idx]);
    let email2 = normalizeEmail(row[email2Idx]);
    let personal = normalizeEmail(row[personalIdx]);
    
    // Apply deduplication logic
    const result = deduplicateEmailSet(primary, email1, email2);
    
    // Update row with deduplicated values
    newRow[primaryIdx] = result.primary;
    newRow[email1Idx] = result.email1;
    newRow[email2Idx] = result.email2;
    newRow[personalIdx] = personal; // Personal email stays separate
    
    return newRow;
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: cleanedData
  };
}

/**
 * Normalizes an email: trim and lowercase
 */
function normalizeEmail(email) {
  if (!email || email.toString().trim() === '') {
    return '';
  }
  return email.toString().trim().toLowerCase();
}

/**
 * Implements email deduplication business logic
 */
function deduplicateEmailSet(primary, email1, email2) {
  // All empty - return empty
  if (!primary && !email1 && !email2) {
    return { primary: '', email1: '', email2: '' };
  }
  
  // All three are the same
  if (primary && primary === email1 && primary === email2) {
    return { primary: '', email1: primary, email2: '' };
  }
  
  // Primary = Email1 (keep Email1 only)
  if (primary && primary === email1) {
    return { primary: '', email1: email1, email2: email2 };
  }
  
  // Primary = Email2 (keep Email2 only)
  if (primary && primary === email2) {
    return { primary: '', email1: email1, email2: email2 };
  }
  
  // Email1 = Email2 (keep Email1 only)
  if (email1 && email1 === email2) {
    return { primary: primary, email1: email1, email2: '' };
  }
  
  // All different - keep all
  return { primary: primary, email1: email1, email2: email2 };
}

// =============================================================================
// STEP 6: PHONE DEDUPLICATION WITH MOBILE PRIORITY
// =============================================================================

/**
 * Deduplicates phones with Contact Mobile Phone priority
 * - Removes dots from phone numbers
 * - If duplicate exists, always keep Contact Mobile Phone
 */
function deduplicatePhonesWithPriority(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const contactPhone1Idx = columnMap[CONFIG.COLUMNS.CONTACT_PHONE_1];
  const companyPhone1Idx = columnMap[CONFIG.COLUMNS.COMPANY_PHONE_1];
  const companyPhone2Idx = columnMap[CONFIG.COLUMNS.COMPANY_PHONE_2];
  const contactMobileIdx = columnMap[CONFIG.COLUMNS.CONTACT_MOBILE];
  
  const cleanedData = data.map(row => {
    const newRow = [...row];
    
    // Get and normalize all phone numbers (remove dots)
    let contactPhone1 = normalizePhone(row[contactPhone1Idx]);
    let companyPhone1 = normalizePhone(row[companyPhone1Idx]);
    let companyPhone2 = normalizePhone(row[companyPhone2Idx]);
    let contactMobile = normalizePhone(row[contactMobileIdx]);
    
    // Create map for comparison (digits only)
    const phoneMap = new Map();
    
    // Priority order: Contact Mobile Phone is highest priority
    let mobileDigits = '';
    if (contactMobile) {
      mobileDigits = extractDigitsOnly(contactMobile);
      phoneMap.set(mobileDigits, 'mobile');
    }
    
    // Now check other phones - if they match mobile, clear them
    if (contactPhone1) {
      const digits = extractDigitsOnly(contactPhone1);
      if (digits === mobileDigits) {
        contactPhone1 = ''; // Clear duplicate
      } else if (!phoneMap.has(digits)) {
        phoneMap.set(digits, 'contact1');
      } else {
        contactPhone1 = ''; // Clear duplicate
      }
    }
    
    if (companyPhone1) {
      const digits = extractDigitsOnly(companyPhone1);
      if (digits === mobileDigits) {
        companyPhone1 = ''; // Clear duplicate
      } else if (!phoneMap.has(digits)) {
        phoneMap.set(digits, 'company1');
      } else {
        companyPhone1 = ''; // Clear duplicate
      }
    }
    
    if (companyPhone2) {
      const digits = extractDigitsOnly(companyPhone2);
      if (digits === mobileDigits) {
        companyPhone2 = ''; // Clear duplicate
      } else if (!phoneMap.has(digits)) {
        phoneMap.set(digits, 'company2');
      } else {
        companyPhone2 = ''; // Clear duplicate
      }
    }
    
    // Update row with cleaned values
    newRow[contactPhone1Idx] = contactPhone1;
    newRow[companyPhone1Idx] = companyPhone1;
    newRow[companyPhone2Idx] = companyPhone2;
    newRow[contactMobileIdx] = contactMobile;
    
    return newRow;
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: cleanedData
  };
}

/**
 * Normalizes phone by removing dots
 */
function normalizePhone(phone) {
  if (!phone || phone.toString().trim() === '') {
    return '';
  }
  
  const phoneStr = phone.toString().trim();
  
  // Remove dots from phone number
  return phoneStr.replace(/\./g, '');
}

/**
 * Extracts only digits from phone for comparison
 */
function extractDigitsOnly(phone) {
  if (!phone) return '';
  return phone.toString().replace(/\D/g, '');
}

// =============================================================================
// STEP 7: NAME PARSING
// =============================================================================

/**
 * Parses full names into first and last names intelligently
 */
function parseNames(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const fullNameIdx = columnMap[CONFIG.COLUMNS.FULL_NAME];
  
  const parsedData = data.map(row => {
    const fullName = row[fullNameIdx] ? row[fullNameIdx].toString().trim() : '';
    const { firstName, lastName } = parseFullName(fullName);
    
    return {
      originalRow: row,
      firstName: firstName,
      lastName: lastName
    };
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: parsedData
  };
}

/**
 * Intelligently parses a full name into first and last name
 */
function parseFullName(fullName) {
  if (!fullName || fullName.trim() === '') {
    return { firstName: '', lastName: '' };
  }
  
  const parts = fullName.trim().split(/\s+/).filter(p => p.length > 0);
  
  if (parts.length === 0) {
    return { firstName: '', lastName: '' };
  }
  
  if (parts.length === 1) {
    return { firstName: parts[0], lastName: '' };
  }
  
  // For 2+ words: split in the middle
  const midPoint = Math.ceil(parts.length / 2);
  const firstName = parts.slice(0, midPoint).join(' ');
  const lastName = parts.slice(midPoint).join(' ');
  
  return { firstName, lastName };
}

// =============================================================================
// STEP 8: OUTPUT SHEET CREATION
// =============================================================================

/**
 * Creates or clears the output sheet
 */
function createOutputSheet(spreadsheet, sourceSheet) {
  let outputSheet = spreadsheet.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);
  
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = spreadsheet.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
  }
  
  return outputSheet;
}

/**
 * Writes cleaned data with specified column order
 * Note: Reads "Company Name - Cleaned" from source, outputs as "Organization"
 */
function writeOutputData(outputSheet, dataObject, sourceSheetName) {
  const { headers, columnMap, data } = dataObject;
  
  // Define output columns in required order
  const outputHeaders = [
    'Contact Full Name',
    'First Name',
    'Last Name',
    'Organization',
    'Primary Email',
    'Email 1',
    'Email 2',
    'Personal Email',
    'Contact Phone 1',
    'Company Phone 1',
    'Company Phone 2',
    'Contact Mobile Phone',
    'Company Description',
    'Website',
    'Company City',
    'Validation Status'
  ];
  
  // Write headers
  outputSheet.getRange(1, 1, 1, outputHeaders.length)
    .setValues([outputHeaders])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // Map of output header to source column
  const headerMapping = {
    'Contact Full Name': CONFIG.COLUMNS.FULL_NAME,
    'Organization': CONFIG.COLUMNS.ORGANIZATION,
    'Primary Email': CONFIG.COLUMNS.PRIMARY_EMAIL,
    'Email 1': CONFIG.COLUMNS.EMAIL_1,
    'Email 2': CONFIG.COLUMNS.EMAIL_2,
    'Personal Email': CONFIG.COLUMNS.PERSONAL_EMAIL,
    'Contact Phone 1': CONFIG.COLUMNS.CONTACT_PHONE_1,
    'Company Phone 1': CONFIG.COLUMNS.COMPANY_PHONE_1,
    'Company Phone 2': CONFIG.COLUMNS.COMPANY_PHONE_2,
    'Contact Mobile Phone': CONFIG.COLUMNS.CONTACT_MOBILE,
    'Company Description': CONFIG.COLUMNS.COMPANY_DESC,
    'Website': CONFIG.COLUMNS.WEBSITE,
    'Company City': CONFIG.COLUMNS.COMPANY_CITY
  };
  
  // Prepare output data
  const outputData = data.map(item => {
    const row = item.originalRow;
    const outputRow = [];
    
    // Add columns in exact order
    outputHeaders.forEach(header => {
      if (header === 'First Name') {
        outputRow.push(item.firstName);
      } else if (header === 'Last Name') {
        outputRow.push(item.lastName);
      } else if (header === 'Validation Status') {
        outputRow.push('');
      } else {
        const sourceColumn = headerMapping[header];
        const colIdx = columnMap[sourceColumn];
        outputRow.push(colIdx !== undefined ? row[colIdx] : '');
      }
    });
    
    return outputRow;
  });
  
  // Write data in batches
  if (outputData.length > 0) {
    const batchSize = CONFIG.BATCH_SIZE;
    for (let i = 0; i < outputData.length; i += batchSize) {
      const batch = outputData.slice(i, i + batchSize);
      const startRow = i + 2;
      outputSheet.getRange(startRow, 1, batch.length, outputHeaders.length)
        .setValues(batch);
    }
  }
  
  // Format sheet
  outputSheet.setFrozenRows(1);
  outputSheet.autoResizeColumns(1, outputHeaders.length);
}

// =============================================================================
// STEP 9: VALIDATION FORMULA
// =============================================================================

/**
 * Adds validation formula to check data integrity
 */
function addValidationFormula(outputSheet, sourceSheetName) {
  const lastRow = outputSheet.getLastRow();
  
  if (lastRow < 2) {
    return;
  }
  
  const validationCol = 16;
  
  const formulaTemplate = `=IFERROR(
 IF(
  AND(
   A{ROW} = INDEX('${sourceSheetName}'!A:A, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   D{ROW} = INDEX('${sourceSheetName}'!B:B, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   N{ROW} = INDEX('${sourceSheetName}'!C:C, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   F{ROW} = INDEX('${sourceSheetName}'!E:E, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   TRIM(SUBSTITUTE(M{ROW},CHAR(10)," ")) =
     TRIM(SUBSTITUTE(INDEX('${sourceSheetName}'!M:M, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),CHAR(10)," ")),
   O{ROW} = INDEX('${sourceSheetName}'!L:L, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   L{ROW} = INDEX('${sourceSheetName}'!K:K, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0))
  ),
  "MATCH OK",
  "DATA MISMATCH"
 ),
 "NOT FOUND"
)`;
  
  // Apply formula to each row
  for (let row = 2; row <= lastRow; row++) {
    const formula = formulaTemplate.replace(/{ROW}/g, row);
    outputSheet.getRange(row, validationCol).setFormula(formula);
  }
  
  // Format validation column
  const validationRange = outputSheet.getRange(2, validationCol, lastRow - 1, 1);
  
  const matchRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MATCH OK')
    .setBackground('#d9ead3')
    .setFontColor('#38761d')
    .setRanges([validationRange])
    .build();
  
  const mismatchRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('DATA MISMATCH')
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([validationRange])
    .build();
  
  const notFoundRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('NOT FOUND')
    .setBackground('#fff2cc')
    .setFontColor('#bf9000')
    .setRanges([validationRange])
    .build();
  
  const rules = outputSheet.getConditionalFormatRules();
  rules.push(matchRule, mismatchRule, notFoundRule);
  outputSheet.setConditionalFormatRules(rules);
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Creates ONE unified menu in the Google Sheets UI
 * Works on any Google Sheet when this Apps Script project is attached
 * 
 * MENU STRUCTURE:
 * 🔧 Data Tools
 * ├── 📊 Data Cleanup
 * │   └── Run Pipeline V3
 * ├── 🏠 Categorize Companies
 * │   ├── Categorize All Companies
 * │   └── Upload Checked Entries
 * ├── 📧 Email Validation
 * │   ├── Validate Primary Email
 * │   ├── Validate Email 1
 * │   ├── Validate Email 2
 * │   └── Validate Personal Email
 * ├── 📞 Phone Validation
 * │   ├── Validate Contact Phone 1
 * │   ├── Validate Company Phone 1
 * │   ├── Validate Company Phone 2
 * │   └── Validate Contact Mobile Phone
 * └── 🎯 Filter by Criteria
 *     ├── Apply Filter
 *     ├── Check Validation Status
 *     └── View Criteria Settings
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create main menu with submenus
  ui.createMenu('🔧 Data Tools')
    
    // Data Cleanup submenu
    .addSubMenu(ui.createMenu('📊 Data Cleanup')
      .addItem('▶️ Run Pipeline V3', 'runDataCleaningPipelineV3')
      .addSeparator()
      .addItem('📋 View Logs', 'showLogs'))
    
    // Categorize Companies submenu
    .addSubMenu(ui.createMenu('🔧 Categorize HVAC Companies')
      .addItem('▶️ Categorize HVAC Companies', 'categorizeHVACCompanies')
      .addSeparator()
      .addItem('📤 Upload Checked Entries', 'uploadCheckedEntriesHVAC'))
    
    // Email Validation submenu
    .addSubMenu(ui.createMenu('📧 Email Validation')
      .addItem('📬 Validate Primary Email', 'validatePrimaryEmail')
      .addItem('📧 Validate Email 1', 'validateEmail1')
      .addItem('📧 Validate Email 2', 'validateEmail2')
      .addItem('📨 Validate Personal Email', 'validatePersonalEmail')
      .addSeparator()
      .addItem('⚙️ Configure Settings', 'showEmailConfigDialog')
      .addItem('📋 View Validation Logs', 'showEmailValidationLogs'))
    
    // Phone Validation submenu
    .addSubMenu(ui.createMenu('📞 Phone Validation')
      .addItem('📱 Validate Contact Phone 1', 'validateContactPhone1')
      .addItem('🏢 Validate Company Phone 1', 'validateCompanyPhone1')
      .addItem('🏢 Validate Company Phone 2', 'validateCompanyPhone2')
      .addItem('📱 Validate Contact Mobile Phone', 'validateContactMobilePhone')
      .addSeparator()
      .addItem('⚙️ Configure Settings', 'showConfigDialog')
      .addItem('📋 View Validation Logs', 'showValidationLogs'))
    
    // Filter by Criteria submenu (NEW!)
    .addSubMenu(ui.createMenu('🎯 Filter by Criteria')
      .addItem('▶️ Apply Filter', 'filterByCriteria')
      .addSeparator()
      .addItem('📊 Check Validation Status', 'showValidationInfo')
      .addItem('⚙️ View Criteria Settings', 'showCriteriaConfig'))
    
    .addToUi();
}

/**
 * Shows execution logs
 */
function showLogs() {
  const logs = Logger.getLog();
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Execution Logs',
    logs || 'No logs available',
    ui.ButtonSet.OK
  );
}

// =============================================================================
// END OF SCRIPT V3
// =============================================================================
