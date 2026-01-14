/**
 * HVAC CATEGORIZATION SYSTEM - V1.0
 * ==================================
 * Based on Home Builder Categorization V7.0
 * 
 * FEATURES:
 * - Works on the ACTIVE spreadsheet (no hardcoded ID)
 * - Integrated with unified menu system
 * - Simplified to 3 categories: HVAC Companies & Related, Other Companies, No Description
 * - Added repair/maintenance keywords (categorized as HVAC)
 * - No popup messages - runs silently
 * - Excludes non-HVAC trades even if they mention installation
 * - Better keyword-based analysis
 * 
 * CATEGORIES:
 * 1. HVAC Companies & Related (includes heating, cooling, ventilation, AC)
 * 2. Other Companies (plumbing, electrical, roofing, everything else)
 * 3. No Description
 * 
 * @author: Claude
 * @version: 1.0 - HVAC-focused categorization
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const HVAC_CATEGORIZATION_CONFIG = {
  // HEADER NAMES TO FIND (case-insensitive)
  HEADER_NAMES: {
    contactFullName: 'Contact Full Name',
    firstName: 'First Name',
    lastName: 'Last Name',
    companyName: 'Organization',
    description: 'Company Description',
    primaryEmail: 'Primary Email',
    email1: 'Email 1',
    email2: 'Email 2',
    personalEmail: 'Personal Email',
    contactPhone1: 'Contact Phone 1',
    companyPhone1: 'Company Phone 1',
    companyPhone2: 'Company Phone 2',
    contactMobilePhone: 'Contact Mobile Phone',
    website: 'Website',
    city: 'Company City'
  },
  
  // Output columns (in order) - all columns from cleaned data
  OUTPUT_COLUMNS: [
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
    'Company City'
  ],
  
  // Starting row
  FIRST_DATA_ROW: 2,
  
  // Fuzzy matching threshold
  FUZZY_THRESHOLD: 0.80,
  
  // Keywords for companies with no description
  NO_DESC_KEYWORDS: [
    ['hvac', 5.0],
    ['heating', 3.5],
    ['cooling', 3.5],
    ['air conditioning', 4.0],
    ['ventilation', 3.0],
    ['climate control', 3.5],
    ['ac', 3.0],
    ['furnace', 3.0],
    ['heat pump', 3.5],
  ],
  
  // EXCLUSION KEYWORDS - Other trades (categorize as OTHER)
  PLUMBING_KEYWORDS: [
    'plumbing contractor',
    'plumbing company',
    'plumber',
    'plumbing service',
    'water heater only',
    'drain cleaning',
    'pipe installation',
    'septic',
    'sewer',
  ],
  
  ELECTRICAL_KEYWORDS: [
    'electrical contractor',
    'electrical company',
    'electrician',
    'electrical service',
    'wiring',
    'electrical panel',
    'generator installation',
    'solar panel',
  ],
  
  // OTHER TRADES KEYWORDS - Categorize as OTHER
  OTHER_TRADES_KEYWORDS: [
    'roofing contractor',
    'roofing company',
    'home builder',
    'construction company',
    'landscaping company',
    'flooring company',
    'painting contractor',
    'carpet installation',
    'tile contractor',
    'cabinet maker',
    'window installation',
    'garage door',
    'concrete contractor',
    'asphalt paving',
    'excavation',
    'tree service',
    'pool contractor',
    'fence contractor',
  ],
  
  // HVAC KEYWORDS - Strong indicators
  HVAC_KEYWORDS: [
    // HVAC - very strong
    'hvac',
    'hvac contractor',
    'hvac company',
    'hvac service',
    'hvac installation',
    'hvac repair',
    'hvac maintenance',
    'hvac technician',
    'hvac specialist',
    'hvac systems',
    
    // Heating
    'heating',
    'heating contractor',
    'heating company',
    'heating service',
    'heating installation',
    'heating repair',
    'heating maintenance',
    'heating system',
    'furnace',
    'furnace installation',
    'furnace repair',
    'boiler',
    'boiler installation',
    'boiler repair',
    'heat pump',
    'heat pump installation',
    'heat pump repair',
    'radiant heat',
    'gas heating',
    'oil heating',
    'electric heating',
    
    // Cooling / Air Conditioning
    'cooling',
    'cooling contractor',
    'cooling service',
    'cooling system',
    'air conditioning',
    'air conditioner',
    'ac',
    'ac contractor',
    'ac company',
    'ac service',
    'ac installation',
    'ac repair',
    'ac maintenance',
    'central air',
    'central ac',
    'ductless ac',
    'mini split',
    'chiller',
    'refrigeration',
    
    // Ventilation
    'ventilation',
    'ventilation contractor',
    'ventilation service',
    'ventilation system',
    'indoor air quality',
    'air quality',
    'ductwork',
    'duct installation',
    'duct cleaning',
    'duct repair',
    'air handler',
    'exhaust fan',
    'whole house fan',
    
    // Climate Control
    'climate control',
    'temperature control',
    'comfort systems',
    'comfort solutions',
    
    // Thermostat
    'thermostat',
    'thermostat installation',
    'smart thermostat',
    
    // Service types
    'hvac emergency',
    'hvac 24/7',
    'hvac tune-up',
    'hvac inspection',
    'hvac replacement',
    'hvac upgrade',
    
    // Commercial HVAC
    'commercial hvac',
    'commercial air conditioning',
    'commercial heating',
    'commercial refrigeration',
    
    // Residential HVAC
    'residential hvac',
    'residential air conditioning',
    'residential heating',
    'home hvac',
    'home air conditioning',
    'home heating',
    
    // Energy efficiency
    'energy efficient hvac',
    'high efficiency hvac',
    'energy star hvac',
    
    // General HVAC terms
    'heating and cooling',
    'heating & cooling',
    'heating and air',
    'heating & air',
    'heating cooling',
    'air conditioning heating',
  ],
  
  // Generic HVAC terms (lower weight)
  GENERIC_HVAC_KEYWORDS: [
    'contractor',
    'contracting',
    'service',
    'installation',
    'repair',
    'maintenance',
  ],
};

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Main categorization function - works on ACTIVE spreadsheet
 */
function categorizeHVACCompanies() {
  const startTime = new Date();
  
  // Get ACTIVE spreadsheet (not hardcoded)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getActiveSheet();
  
  if (!spreadsheet || !sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'No active spreadsheet or sheet found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log('=== STARTING HVAC CATEGORIZATION V1.0 ===');
  Logger.log(`Spreadsheet: ${spreadsheet.getName()}`);
  Logger.log(`Active Sheet: ${sourceSheet.getName()}`);
  
  // Find columns
  const columnMap = findColumnsByHeadersHVAC(sourceSheet);
  if (!columnMap) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Could not find required headers in the active sheet.\n\n' +
      'Required headers:\n' +
      '- Organization (or Company Name)\n' +
      '- Company Description\n\n' +
      'These are the minimum columns needed for categorization.\n' +
      'Other columns are optional but recommended.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Show progress
  SpreadsheetApp.getUi().alert(
    'HVAC Categorization Started',
    'Processing companies... This may take a few minutes.\n\n' +
    'Check the script logs for progress.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  // Process companies
  const results = processAllCompaniesHVAC(sourceSheet, columnMap);
  
  // Create category sheets
  const categoryStats = createCategorySheetsHVAC(spreadsheet, sourceSheet, results, columnMap);
  
  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(1);
  
  Logger.log('=== CATEGORIZATION COMPLETE ===');
  Logger.log(`Duration: ${duration} seconds`);
  Logger.log(`HVAC Companies: ${categoryStats.hvac}`);
  Logger.log(`Other Companies: ${categoryStats.other}`);
  Logger.log(`No Description: ${categoryStats.noDesc}`);
  
  // Show completion message
  SpreadsheetApp.getUi().alert(
    'HVAC Categorization Complete',
    `Processing complete in ${duration} seconds!\n\n` +
    `Results:\n` +
    `- HVAC Companies & Related: ${categoryStats.hvac}\n` +
    `- Other Companies: ${categoryStats.other}\n` +
    `- No Description: ${categoryStats.noDesc}\n\n` +
    `Total: ${categoryStats.hvac + categoryStats.other + categoryStats.noDesc}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// PROCESSING FUNCTIONS
// ============================================================================

function processAllCompaniesHVAC(sheet, columnMap) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  Logger.log(`Processing rows ${HVAC_CATEGORIZATION_CONFIG.FIRST_DATA_ROW} to ${lastRow}`);
  
  const results = [];
  
  for (let rowNum = HVAC_CATEGORIZATION_CONFIG.FIRST_DATA_ROW; rowNum <= lastRow; rowNum++) {
    const rowData = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    
    const companyName = columnMap.nameCol !== -1 ? String(rowData[columnMap.nameCol] || '').trim() : '';
    const description = columnMap.descCol !== -1 ? String(rowData[columnMap.descCol] || '').trim() : '';
    const website = columnMap.webCol !== -1 ? String(rowData[columnMap.webCol] || '').trim() : '';
    
    if (!companyName && !description) {
      continue; // Skip completely empty rows
    }
    
    const result = categorizeCompanyHVAC(companyName, description, website, rowNum);
    results.push(result);
    
    if (rowNum % 100 === 0) {
      Logger.log(`  Processed ${rowNum - HVAC_CATEGORIZATION_CONFIG.FIRST_DATA_ROW + 1} rows...`);
    }
  }
  
  Logger.log(`Total processed: ${results.length}`);
  return results;
}

function categorizeCompanyHVAC(companyName, description, website, rowNum) {
  const name = companyName.toLowerCase();
  const desc = description.toLowerCase();
  const web = extractDomainHVAC(website);
  
  const combinedText = `${name} ${desc} ${web || ''}`.toLowerCase();
  
  // NO DESCRIPTION category
  if (!desc || desc.length < 5) {
    let score = 0;
    HVAC_CATEGORIZATION_CONFIG.NO_DESC_KEYWORDS.forEach(([keyword, weight]) => {
      if (name.includes(keyword) || (web && web.includes(keyword))) {
        score += weight;
      }
    });
    
    if (score >= 3.0) {
      return {
        rowNum: rowNum,
        category: 'HVAC Companies & Related',
        reason: 'Name suggests HVAC (no description)',
        confidence: 'Low',
        score: score
      };
    }
    
    return {
      rowNum: rowNum,
      category: 'No Description',
      reason: 'Missing description',
      confidence: 'N/A',
      score: 0
    };
  }
  
  // Check for EXCLUSIONS first (plumbing, electrical, etc.)
  for (const keyword of HVAC_CATEGORIZATION_CONFIG.PLUMBING_KEYWORDS) {
    if (fuzzyMatchHVAC(combinedText, keyword, HVAC_CATEGORIZATION_CONFIG.FUZZY_THRESHOLD)) {
      return {
        rowNum: rowNum,
        category: 'Other Companies',
        reason: 'Plumbing company',
        confidence: 'High',
        score: 0
      };
    }
  }
  
  for (const keyword of HVAC_CATEGORIZATION_CONFIG.ELECTRICAL_KEYWORDS) {
    if (fuzzyMatchHVAC(combinedText, keyword, HVAC_CATEGORIZATION_CONFIG.FUZZY_THRESHOLD)) {
      return {
        rowNum: rowNum,
        category: 'Other Companies',
        reason: 'Electrical company',
        confidence: 'High',
        score: 0
      };
    }
  }
  
  for (const keyword of HVAC_CATEGORIZATION_CONFIG.OTHER_TRADES_KEYWORDS) {
    if (fuzzyMatchHVAC(combinedText, keyword, HVAC_CATEGORIZATION_CONFIG.FUZZY_THRESHOLD)) {
      return {
        rowNum: rowNum,
        category: 'Other Companies',
        reason: 'Other trade company',
        confidence: 'High',
        score: 0
      };
    }
  }
  
  // Check for HVAC keywords
  let hvacScore = 0;
  let matchedKeywords = [];
  
  for (const keyword of HVAC_CATEGORIZATION_CONFIG.HVAC_KEYWORDS) {
    if (fuzzyMatchHVAC(combinedText, keyword, HVAC_CATEGORIZATION_CONFIG.FUZZY_THRESHOLD)) {
      // Weight by keyword specificity
      if (keyword.includes('hvac')) {
        hvacScore += 5.0;
      } else if (keyword.length > 15) {
        hvacScore += 3.0;
      } else {
        hvacScore += 1.0;
      }
      matchedKeywords.push(keyword);
    }
  }
  
  // Strong HVAC match
  if (hvacScore >= 5.0) {
    return {
      rowNum: rowNum,
      category: 'HVAC Companies & Related',
      reason: `Strong HVAC keywords: ${matchedKeywords.slice(0, 3).join(', ')}`,
      confidence: 'High',
      score: hvacScore
    };
  }
  
  // Medium HVAC match
  if (hvacScore >= 2.0) {
    return {
      rowNum: rowNum,
      category: 'HVAC Companies & Related',
      reason: `HVAC keywords: ${matchedKeywords.slice(0, 3).join(', ')}`,
      confidence: 'Medium',
      score: hvacScore
    };
  }
  
  // Generic keywords only
  let genericScore = 0;
  for (const keyword of HVAC_CATEGORIZATION_CONFIG.GENERIC_HVAC_KEYWORDS) {
    if (combinedText.includes(keyword)) {
      genericScore += 0.5;
    }
  }
  
  if (genericScore >= 1.0 && matchedKeywords.length > 0) {
    return {
      rowNum: rowNum,
      category: 'HVAC Companies & Related',
      reason: 'Generic HVAC terms',
      confidence: 'Low',
      score: hvacScore
    };
  }
  
  // Default: Other Companies
  return {
    rowNum: rowNum,
    category: 'Other Companies',
    reason: 'No HVAC keywords found',
    confidence: 'High',
    score: 0
  };
}

// ============================================================================
// SHEET CREATION
// ============================================================================

function createCategorySheetsHVAC(spreadsheet, sourceSheet, results, columnMap) {
  const categoryStats = {
    hvac: 0,
    other: 0,
    noDesc: 0
  };
  
  // Group by category
  const categorized = {
    'HVAC Companies & Related': [],
    'Other Companies': [],
    'No Description': []
  };
  
  results.forEach(result => {
    categorized[result.category].push(result);
    
    if (result.category === 'HVAC Companies & Related') {
      categoryStats.hvac++;
    } else if (result.category === 'Other Companies') {
      categoryStats.other++;
    } else {
      categoryStats.noDesc++;
    }
  });
  
  // Create sheets for each category
  Object.keys(categorized).forEach(category => {
    const items = categorized[category];
    if (items.length === 0) return;
    
    createCategorySheetHVAC(spreadsheet, sourceSheet, category, items, columnMap);
  });
  
  return categoryStats;
}

function createCategorySheetHVAC(spreadsheet, sourceSheet, categoryName, items, columnMap) {
  Logger.log(`Creating sheet: ${categoryName} (${items.length} items)`);
  
  // Delete existing sheet
  let existingSheet = spreadsheet.getSheetByName(categoryName);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // Create new sheet
  const newSheet = spreadsheet.insertSheet(categoryName);
  
  // Add analysis columns
  const analysisHeaders = ['Category Reason', 'Confidence', 'Score', 'Original Row', 'Check to Upload'];
  const outputHeaders = HVAC_CATEGORIZATION_CONFIG.OUTPUT_COLUMNS.concat(analysisHeaders);
  
  // Set headers
  newSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  newSheet.getRange(1, 1, 1, outputHeaders.length).setFontWeight('bold');
  newSheet.getRange(1, 1, 1, outputHeaders.length).setBackground('#4285f4');
  newSheet.getRange(1, 1, 1, outputHeaders.length).setFontColor('#ffffff');
  
  // Get source data
  const sourceLastCol = sourceSheet.getLastColumn();
  const sourceData = sourceSheet.getDataRange().getValues();
  
  if (items.length > 0) {
    const dataRows = items.map(item => {
      const rowNum = item.rowNum;
      const sourceRow = sourceData[rowNum - 1];
      
      // Extract data for output columns
      const selectedData = HVAC_CATEGORIZATION_CONFIG.OUTPUT_COLUMNS.map(colName => {
        const headerRow = sourceData[0];
        const colIndex = headerRow.findIndex(h => 
          String(h).trim().toLowerCase() === colName.toLowerCase()
        );
        return colIndex !== -1 ? sourceRow[colIndex] : '';
      });
      
      // Add analysis data
      const analysisData = [
        item.reason,
        item.confidence,
        item.score,
        rowNum,
        false
      ];
      
      return selectedData.concat(analysisData);
    });
    
    newSheet.getRange(2, 1, dataRows.length, outputHeaders.length).setValues(dataRows);
    
    // Format checkbox column
    const checkboxCol = outputHeaders.length;
    newSheet.getRange(2, checkboxCol, dataRows.length, 1).insertCheckboxes();
    
    // Format
    newSheet.setFrozenRows(1);
    newSheet.autoResizeColumns(1, outputHeaders.length);
  }
  
  return categoryStats;
}

// ============================================================================
// UPLOAD ENTRIES FUNCTION
// ============================================================================

function uploadCheckedEntriesHVAC() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  
  Logger.log('=== UPLOADING CHECKED HVAC ENTRIES ===');
  
  const allSheets = spreadsheet.getSheets();
  const categorySheets = allSheets.filter(sheet => {
    const name = sheet.getName();
    return name === 'HVAC Companies & Related' || 
           name === 'Other Companies' || 
           name === 'No Description';
  });
  
  Logger.log(`Found ${categorySheets.length} category sheets`);
  
  const checkedRows = [];
  let totalChecked = 0;
  
  categorySheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return;
    
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    let checkboxCol = -1;
    let originalRowCol = -1;
    
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].includes('Check')) checkboxCol = i;
      if (headers[i] === 'Original Row') originalRowCol = i;
    }
    
    if (checkboxCol === -1 || originalRowCol === -1) return;
    
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    data.forEach((row, idx) => {
      if (row[checkboxCol] === true) {
        checkedRows.push({
          originalRowNum: Number(row[originalRowCol]),
          categorySheet: sheetName
        });
        totalChecked++;
      }
    });
  });
  
  if (totalChecked === 0) {
    SpreadsheetApp.getUi().alert(
      'No Checked Entries',
      'No entries are checked. Please check some entries first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log(`Total checked: ${totalChecked}`);
  
  // Find the source sheet (first sheet that's not a category sheet)
  let sourceSheet = null;
  for (const sheet of allSheets) {
    const name = sheet.getName();
    if (name !== 'HVAC Companies & Related' && 
        name !== 'Other Companies' && 
        name !== 'No Description' &&
        name !== 'Final Results') {
      sourceSheet = sheet;
      break;
    }
  }
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Could not find source sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const sourceLastRow = sourceSheet.getLastRow();
  const sourceLastCol = sourceSheet.getLastColumn();
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastCol).getValues()[0];
  const sourceData = sourceSheet.getRange(1, 1, sourceLastRow, sourceLastCol).getValues();
  
  let finalSheet = spreadsheet.getSheetByName('Final Results');
  if (!finalSheet) {
    finalSheet = spreadsheet.insertSheet('Final Results');
  } else {
    finalSheet.clear();
  }
  
  const finalHeaders = sourceHeaders.concat(['Source Category']);
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setFontWeight('bold');
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setBackground('#34a853');
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setFontColor('#ffffff');
  
  const finalData = [];
  let successCount = 0;
  
  checkedRows.forEach(item => {
    const rowNum = item.originalRowNum;
    const category = item.categorySheet;
    
    if (!rowNum || isNaN(rowNum) || rowNum < 1 || rowNum > sourceLastRow) {
      Logger.log(`ERROR: Invalid row ${rowNum}`);
      return;
    }
    
    const arrayIndex = rowNum - 1;
    const originalRow = sourceData[arrayIndex];
    
    if (originalRow) {
      finalData.push(originalRow.concat([category]));
      successCount++;
    }
  });
  
  if (finalData.length > 0) {
    finalSheet.getRange(2, 1, finalData.length, finalHeaders.length).setValues(finalData);
  }
  
  finalSheet.setFrozenRows(1);
  finalSheet.autoResizeColumns(1, finalHeaders.length);
  
  Logger.log(`=== UPLOAD COMPLETE: ${successCount} rows ===`);
  
  SpreadsheetApp.getUi().alert(
    'Upload Complete',
    `Successfully uploaded ${successCount} checked entries to "Final Results" sheet.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function findColumnsByHeadersHVAC(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const columnMap = {
    contactFullNameCol: -1,
    firstNameCol: -1,
    lastNameCol: -1,
    nameCol: -1,
    descCol: -1,
    primaryEmailCol: -1,
    email1Col: -1,
    email2Col: -1,
    personalEmailCol: -1,
    contactPhone1Col: -1,
    companyPhone1Col: -1,
    companyPhone2Col: -1,
    contactMobilePhoneCol: -1,
    webCol: -1,
    cityCol: -1
  };
  
  for (let i = 0; i < headerRow.length; i++) {
    const header = String(headerRow[i]).trim();
    
    if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.contactFullName)) {
      columnMap.contactFullNameCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.firstName)) {
      columnMap.firstNameCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.lastName)) {
      columnMap.lastNameCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.companyName)) {
      columnMap.nameCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.description)) {
      columnMap.descCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.primaryEmail)) {
      columnMap.primaryEmailCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.email1)) {
      columnMap.email1Col = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.email2)) {
      columnMap.email2Col = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.personalEmail)) {
      columnMap.personalEmailCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.contactPhone1)) {
      columnMap.contactPhone1Col = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.companyPhone1)) {
      columnMap.companyPhone1Col = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.companyPhone2)) {
      columnMap.companyPhone2Col = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.contactMobilePhone)) {
      columnMap.contactMobilePhoneCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.website)) {
      columnMap.webCol = i;
    } else if (matchesHeaderHVAC(header, HVAC_CATEGORIZATION_CONFIG.HEADER_NAMES.city)) {
      columnMap.cityCol = i;
    }
  }
  
  // Only require the essential columns for categorization
  if (columnMap.nameCol === -1 || columnMap.descCol === -1) {
    return null;
  }
  
  return columnMap;
}

function matchesHeaderHVAC(actual, expected) {
  const normalize = (str) => str.toLowerCase().replace(/\s+/g, ' ').trim();
  return normalize(actual) === normalize(expected);
}

function extractDomainHVAC(input) {
  if (!input) return null;
  input = input.toLowerCase().trim();
  input = input.replace(/^(https?:\/\/)?(www\.)?/, '');
  if (input.includes('@')) input = input.split('@')[1];
  input = input.split('/')[0].split('?')[0];
  input = input.replace(/\.(com|net|org|co|io|biz|info|us)$/i, '');
  return input || null;
}

function fuzzyMatchHVAC(text, keyword, threshold) {
  if (text.includes(keyword)) return true;
  
  const variations = generateVariationsHVAC(keyword);
  for (const variation of variations) {
    if (text.includes(variation)) return true;
  }
  
  return false;
}

function generateVariationsHVAC(keyword) {
  const variations = [];
  if (keyword.endsWith('s')) variations.push(keyword.slice(0, -1));
  else variations.push(keyword + 's');
  if (keyword.includes('-')) {
    variations.push(keyword.replace(/-/g, ' '));
    variations.push(keyword.replace(/-/g, ''));
  }
  if (keyword.includes(' ')) {
    variations.push(keyword.replace(/\s+/g, '-'));
    variations.push(keyword.replace(/\s+/g, ''));
  }
  return variations;
}
