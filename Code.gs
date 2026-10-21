
/***************************************************************
 * PROCUREMENT AUTOMATION - COMPLETE CODE.GS
 *
 * FEATURES
 * 1. Project-description dropdown from external Entries sheet.
 * 2. Autofill Classification, Category, UOM and Unit Cost.
 * 3. Calculate Estimated Budget = Unit Cost x Quantity.
 * 4. Procurement-mode dropdown from external Modes sheet.
 * 5. Clear related fields when project description is deleted.
 * 6. Automatically summarize amounts by Category.
 * 7. Support headers in different column positions.
 ***************************************************************/


/***************************************************************
 * CONFIGURATION
 ***************************************************************/

const ENTRY_CONFIG = {
  SOURCE_SPREADSHEET_ID:
    '18U_xk2kJaOaDYizro5ZlyymCLQiLgiuCxjtKFr-W7S8',

  SOURCE_SHEET_NAME: 'Entries',
  HELPER_SHEET_NAME: '_EntryDropdownOptions',
  MODES_SHEET_NAME: 'Modes',
  MODES_COLUMN: 1,
  MODES_START_ROW: 1,
  // The PPMP template has title and multi-row table headings above
  // the actual field labels (the entry area starts around row 15).
  // Search deep enough to discover those labels reliably.
  HEADER_SEARCH_ROWS: 30,

  SOURCE_HEADERS: {
    description: 'General Description/Generic Name',
    specifications: 'Technical Specifications',
    uom: 'UOM',
    unitCost: 'Unit Cost',
    classification: 'Classification',
    category: 'Category'
  },

  DESTINATION_HEADERS: {
    description:
      'General Description and Objective of the Project to be Procured',

    classification:
      'Type of the Project to be Procured',

    category: 'Category',
    uom: 'Unit or Measurement/Size',
    unitCost: 'Unit Cost',
    quantity: 'Quantity',

    estimatedBudget:
      'Estimated Budget / Authorized Budgetary Allocation (PhP)',

    recommendedMode:
      'Recommended Mode of Procurement'
  },

  PPMP_FORM_SHEET: 'PPMP Form',
  PPMP_SUMMARY_SHEET: 'PPMP Summary'
};


/***************************************************************
 * PPMP FORM MENU
 *
 * Google Apps Script cannot create a sheet drawing/button directly.
 * This menu provides an immediately usable Add Entry command.
 ***************************************************************/
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('PPMP Form')
    .addItem('Refresh End-User Dropdown (B9)', 'refreshEndUserDropdownB9')
    .addToUi();
}


/***************************************************************
 * MAIN SETUP
 *
 * Run this function manually once after saving the code.
 ***************************************************************/

function setupEntryDropdownAndAutofill() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('No active spreadsheet was found.');
  }

  const sourceSpreadsheet =
    SpreadsheetApp.openById(
      ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
    );

  const sourceSheet =
    sourceSpreadsheet.getSheetByName(
      ENTRY_CONFIG.SOURCE_SHEET_NAME
    );

  if (!sourceSheet) {
    throw new Error(
      'Source worksheet "' +
      ENTRY_CONFIG.SOURCE_SHEET_NAME +
      '" was not found.'
    );
  }

  const sourceData =
    loadSourceLookup_(sourceSheet);

  if (sourceData.options.length === 0) {
    throw new Error(
      'No project descriptions were found in the Entries worksheet.'
    );
  }

  // Remember which destination spreadsheet this script serves so
  // edits made in the separate Entries spreadsheet can refresh it.
  PropertiesService.getScriptProperties().setProperty(
    'PPMP_DESTINATION_SPREADSHEET_ID',
    spreadsheet.getId()
  );

  const descriptionDropdownsApplied =
    refreshEntryDropdownOptions_(spreadsheet, sourceData);

  // Apply procurement-mode dropdowns.
  const modeDropdownsApplied =
    applyRecommendedModeDropdown_(spreadsheet);

  // Enforce PPMP Form validation boundaries after all dropdown routines:
  // Row 15 is the first entry row. Never leave validations above it.
  const ppmpFormSheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (ppmpFormSheet && ppmpFormSheet.getMaxRows() >= 14) {
    // Clear unwanted validations above the line-item table, but B9 is
    // an intentional End-User dropdown and must be applied afterward.
    ppmpFormSheet.getRange(1, 1, 14, ppmpFormSheet.getMaxColumns()).clearDataValidations();
  }

  // Apply B9 last: the helper list must be written after the general
  // dropdown refresh clears the helper sheet, and after row 1–14 rules
  // are cleared so the End-User validation is not accidentally removed.
  applyEndUserDropdownB9_(spreadsheet, sourceSpreadsheet);

  // Recreate both installable edit triggers:
  // 1) edits inside PPMP Form; 2) edits inside the external Entries sheet.
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (
      handler === 'handleEntrySelection' ||
      handler === 'handleEntriesSourceEdit'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleEntrySelection')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  ScriptApp.newTrigger('handleEntriesSourceEdit')
    .forSpreadsheet(sourceSpreadsheet)
    .onEdit()
    .create();

  // Place the functional Add Entry Row control immediately below the
  // last populated line item in the PPMP Form table.
  createAddEntryRowControl_(ppmpFormSheet);

  // Populate PPMP Summary immediately.
  updatePPMPSummary_(spreadsheet);

  SpreadsheetApp.flush();

  spreadsheet.toast(
    'Setup complete. Description dropdowns: ' +
    descriptionDropdownsApplied +
    '. Procurement-mode dropdowns: ' +
    modeDropdownsApplied +
    '. PPMP Summary updated.',
    'Procurement Automation',
    8
  );
}



/***************************************************************
 * ADD A NEW PPMP FORM ENTRY ROW
 *
 * Assign the sheet's "Add Entry" button to addPPMPEntryRow.
 * The new row inherits formatting and validations from the previous
 * entry row, then receives fresh description and procurement-mode
 * dropdowns. Existing rows and values are preserved.
 ***************************************************************/
function addPPMPEntryRow() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('No active spreadsheet was found.');

  const sheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (!sheet) throw new Error('The PPMP Form sheet was not found.');

  const info = findDestinationHeaders_(sheet);
  if (!info) throw new Error('Could not identify the PPMP Form entry columns.');

  const descriptionColumn = info.headerMap[
    normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.description)
  ];
  const modeColumn = info.headerMap[
    normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.recommendedMode)
  ];
  if (descriptionColumn === undefined || modeColumn === undefined) {
    throw new Error('Could not identify the Description and Recommended Mode columns.');
  }

  const firstEntryRow = 15;
  const lastDescriptionRow = findLastDescriptionEntryRow_(
    sheet, descriptionColumn + 1, firstEntryRow
  );
  const newRow = Math.max(firstEntryRow, lastDescriptionRow + 1);

  if (newRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), newRow - sheet.getMaxRows() + 1);
  }

  // Copy the previous row's visual layout and cell validation rules.
  // For the first entry, copy the template's first entry row to itself.
  const templateRow = Math.max(firstEntryRow, newRow - 1);
  if (newRow !== templateRow) {
    sheet.getRange(templateRow, 1, 1, sheet.getMaxColumns())
      .copyTo(sheet.getRange(newRow, 1, 1, sheet.getMaxColumns()),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.getRange(templateRow, 1, 1, sheet.getMaxColumns())
      .copyTo(sheet.getRange(newRow, 1, 1, sheet.getMaxColumns()),
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    sheet.setRowHeight(newRow, sheet.getRowHeight(templateRow));
  }

  // New entry should be blank, without carrying forward the previous item's data.
  sheet.getRange(newRow, 1, 1, sheet.getMaxColumns()).clearContent();

  const sourceSpreadsheet = SpreadsheetApp.openById(
    ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
  );
  const sourceSheet = sourceSpreadsheet.getSheetByName(
    ENTRY_CONFIG.SOURCE_SHEET_NAME
  );
  if (!sourceSheet) throw new Error('The source Entries worksheet was not found.');

  const sourceData = loadSourceLookup_(sourceSheet);
  if (!sourceData.options.length) {
    throw new Error('No project descriptions were found in the Entries worksheet.');
  }

  let helperSheet = spreadsheet.getSheetByName(ENTRY_CONFIG.HELPER_SHEET_NAME);
  if (!helperSheet) helperSheet = spreadsheet.insertSheet(ENTRY_CONFIG.HELPER_SHEET_NAME);
  // Column A is reserved for project descriptions. Do not clear column B,
  // which stores the End-User choices used by PPMP Form!B9:D9 validation.
  helperSheet.getRange(1, 1, helperSheet.getMaxRows(), 1).clearContent();
  helperSheet.getRange(1, 1, sourceData.options.length, 1)
    .setValues(sourceData.options.map(function(value) { return [value]; }));
  if (!helperSheet.isSheetHidden()) helperSheet.hideSheet();

  const descriptionRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(helperSheet.getRange(1, 1, sourceData.options.length, 1), true)
    .setAllowInvalid(false)
    .setHelpText('Select a project description from the Entries source.')
    .build();
  sheet.getRange(newRow, descriptionColumn + 1).setDataValidation(descriptionRule);

  const modesSheet = sourceSpreadsheet.getSheetByName(ENTRY_CONFIG.MODES_SHEET_NAME);
  if (!modesSheet || modesSheet.getLastRow() < ENTRY_CONFIG.MODES_START_ROW) {
    throw new Error('The Modes worksheet contains no procurement modes.');
  }
  const rawModes = modesSheet.getRange(
    ENTRY_CONFIG.MODES_START_ROW, ENTRY_CONFIG.MODES_COLUMN,
    modesSheet.getLastRow() - ENTRY_CONFIG.MODES_START_ROW + 1, 1
  ).getDisplayValues();
  const modes = [];
  const seenModes = new Set();
  rawModes.forEach(function(row) {
    const mode = String(row[0] || '').trim();
    if (!mode) return;
    const header = normalizeHeader_(mode);
    if (header === 'modeofprocurement' || header === 'modes' ||
        header === 'recommendedmodeofprocurement') return;
    const key = normalizeValue_(mode);
    if (!seenModes.has(key)) { seenModes.add(key); modes.push(mode); }
  });
  if (!modes.length) throw new Error('No procurement modes were found.');
  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(modes, true)
    .setAllowInvalid(false)
    .setHelpText('Select a procurement mode from the Modes worksheet.')
    .build();
  sheet.getRange(newRow, modeColumn + 1).setDataValidation(modeRule);

  spreadsheet.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange(newRow, descriptionColumn + 1));
  spreadsheet.toast(
    'Added entry row ' + newRow + ' with dropdowns and automatic autofill support.',
    'PPMP Form', 5
  );
  return newRow;
}


/**
 * Adds a checkbox-based Add Entry Row control below the last populated
 * PPMP line item. Checking it invokes addPPMPEntryRow via the existing
 * installable edit trigger; this avoids a custom menu command.
 */
function createAddEntryRowControl_(sheet) {
  if (!sheet) return;
  const firstEntryRow = 15;
  const lastDescriptionRow = findLastDescriptionEntryRow_(sheet, 1, firstEntryRow);
  const controlRow = Math.max(firstEntryRow, lastDescriptionRow + 1);

  // Ensure a blank row exists below the last item.
  if (controlRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), controlRow - sheet.getMaxRows());
  }

  const controlCell = sheet.getRange(controlRow, 1);
  const labelCell = sheet.getRange(controlRow, 2);
  controlCell.clearContent().clearDataValidations();
  controlCell.insertCheckboxes();
  controlCell.setValue(false);
  controlCell.setHorizontalAlignment('center');
  controlCell.setVerticalAlignment('middle');
  controlCell.setBackground('#d9ead3');
  controlCell.setNote('Check this box to add a new PPMP entry row.');
  labelCell.setValue('＋ ADD ENTRY ROW');
  labelCell.setFontWeight('bold');
  labelCell.setHorizontalAlignment('left');
  labelCell.setVerticalAlignment('middle');
  labelCell.setBackground('#d9ead3');
  labelCell.setNote('Check the box in column A to add a new entry row.');
  sheet.setRowHeight(controlRow, Math.max(28, sheet.getRowHeight(controlRow)));
}

/***************************************************************
 * STANDALONE REPAIR FOR PPMP FORM!B9
 *
 * Run this directly if the main setup function fails for an unrelated
 * reason. It reads Data List > End-User and applies validation to B9.
 ***************************************************************/
function refreshEndUserDropdownB9() {
  const destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!destinationSpreadsheet) {
    throw new Error('Open the destination spreadsheet containing PPMP Form, then run this function again.');
  }

  const sourceSpreadsheet = SpreadsheetApp.openById(
    ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
  );

  applyEndUserDropdownB9_(destinationSpreadsheet, sourceSpreadsheet);
  SpreadsheetApp.flush();

  destinationSpreadsheet.toast(
    'End-user dropdown refreshed in PPMP Form!B9.',
    'PPMP Form',
    5
  );
}


/***************************************************************
 * END-USER DROPDOWN IN PPMP FORM B9
 * Uses all nonblank values in column A, starting at row 2, from
 * the external Data List spreadsheet's End-User worksheet.
 ***************************************************************/
function applyEndUserDropdownB9_(destinationSpreadsheet, sourceSpreadsheet) {
  const sourceSheet = sourceSpreadsheet.getSheetByName('End-User');
  if (!sourceSheet) {
    throw new Error('Source worksheet "End-User" was not found in the Data List spreadsheet.');
  }

  const destinationSheet = destinationSpreadsheet.getSheetByName('PPMP Form');
  if (!destinationSheet) {
    throw new Error('The PPMP Form worksheet was not found.');
  }

  const lastRow = sourceSheet.getLastRow();
  const values = lastRow >= 2
    ? sourceSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues()
    : [];
  const options = [];
  const seen = new Set();
  values.forEach(function(row) {
    const value = String(row[0] || '').trim();
    if (!value) return;
    const key = normalizeValue_(value);
    if (seen.has(key)) return;
    seen.add(key);
    options.push(value);
  });

  if (!options.length) {
    throw new Error('No End-User items were found in column A from row 2 onward in the Data List End-User sheet.');
  }

  // Use an explicit dropdown list copied directly from the source values.
  // This avoids dependencies on the hidden helper sheet, which is also used
  // by the project-description dropdown and may be refreshed independently.
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .setHelpText('Select an End-User or Implementing Unit from Data List > End-User.')
    .build();

  // The End-user or Implementing Unit input is a merged range B9:D9.
  // Google Sheets stores data validation on the top-left cell of a merged
  // range, so apply the rule to B9 (the anchor cell) and the full merged range.
  const target = destinationSheet.getRange('B9:D9');
  target.setDataValidation(rule);
}

/***************************************************************
 * REFRESH DROPDOWN OPTIONS
 *
 * Rebuilds the hidden helper list and reapplies validations after
 * Entries source data changes.
 ***************************************************************/

function refreshEntryDropdownOptions_(spreadsheet, sourceData) {
  let helperSheet = spreadsheet.getSheetByName(
    ENTRY_CONFIG.HELPER_SHEET_NAME
  );

  if (!helperSheet) {
    helperSheet = spreadsheet.insertSheet(
      ENTRY_CONFIG.HELPER_SHEET_NAME
    );
  }

  // Only refresh the description list in column A. Column B stores the
  // separate End-User validation list for PPMP Form!B9 and must be preserved.
  helperSheet.getRange(1, 1, helperSheet.getMaxRows(), 1).clearContent();

  if (sourceData.options.length) {
    helperSheet.getRange(
      1, 1, sourceData.options.length, 1
    ).setValues(sourceData.options.map(function(value) {
      return [value];
    }));
  }

  if (!helperSheet.isSheetHidden()) helperSheet.hideSheet();

  let updated = 0;

  spreadsheet.getSheets().forEach(function(sheet) {
    if (sheet.getName() === ENTRY_CONFIG.HELPER_SHEET_NAME) return;

    const info = findDestinationHeaders_(sheet);
    if (!info) return;

    const descriptionColumn = info.headerMap[
      normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.description)
    ];
    if (descriptionColumn === undefined || sourceData.options.length === 0) return;

    // PPMP Form dropdown validation starts specifically at row 15.
    const validationStartRow = info.dataStartRow;

    if (sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET && sheet.getMaxRows() >= 14) {
      sheet.getRange(1, 1, 14, sheet.getMaxColumns()).clearDataValidations();
    }

    // On PPMP Form, validation must end at the last row that actually
    // contains a project description. Remove old rules below that row.
    const validationEndRow =
      sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET
        ? findLastDescriptionEntryRow_(sheet, descriptionColumn + 1, info.dataStartRow)
        : sheet.getMaxRows();

    const availableRows = validationEndRow - validationStartRow + 1;
    if (sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET) {
      const clearFromRow = Math.max(info.dataStartRow, validationEndRow + 1);
      if (clearFromRow <= sheet.getMaxRows()) {
        sheet.getRange(
          clearFromRow, descriptionColumn + 1,
          sheet.getMaxRows() - clearFromRow + 1, 1
        ).clearDataValidations();
      }
    }
    if (availableRows <= 0) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(helperSheet.getRange(1, 1, sourceData.options.length, 1), true)
      .setAllowInvalid(false)
      .setHelpText('Select a project description from the Entries source.')
      .build();

    sheet.getRange(
      validationStartRow, descriptionColumn + 1, availableRows, 1
    ).setDataValidation(rule);

    // Explicitly clear the reported cells after applying validation.
    if (sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET && sheet.getMaxRows() >= 4) {
      sheet.getRange('A4').clearDataValidations();
      sheet.getRange('F4').clearDataValidations();
    }

    updated++;
  });

  return updated;
}


/**
 * Returns the last row at or below startRow containing a nonblank
 * description. If there are no descriptions, returns startRow - 1.
 */
function findLastDescriptionEntryRow_(sheet, descriptionColumn, startRow) {
  const lastSheetRow = sheet.getLastRow();
  if (lastSheetRow < startRow) return startRow - 1;

  const values = sheet.getRange(
    startRow, descriptionColumn,
    lastSheetRow - startRow + 1, 1
  ).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim() !== '') {
      return startRow + i;
    }
  }

  return startRow - 1;
}


/***************************************************************
 * SOURCE ENTRIES EDIT TRIGGER
 *
 * Runs when the separate source spreadsheet's Entries sheet is
 * edited. Refreshes dropdowns, updates existing PPMP Form rows,
 * then rebuilds the category totals in PPMP Summary.
 ***************************************************************/

function handleEntriesSourceEdit(e) {
  if (!e || !e.range) return;

  const sourceSheet = e.range.getSheet();
  if (sourceSheet.getName() !== ENTRY_CONFIG.SOURCE_SHEET_NAME) return;

  const destinationId = PropertiesService.getScriptProperties()
    .getProperty('PPMP_DESTINATION_SPREADSHEET_ID');

  if (!destinationId) {
    throw new Error(
      'Destination spreadsheet is not configured. Run setupEntryDropdownAndAutofill() once.'
    );
  }

  const destinationSpreadsheet = SpreadsheetApp.openById(destinationId);
  const formSheet = destinationSpreadsheet.getSheetByName(
    ENTRY_CONFIG.PPMP_FORM_SHEET
  );

  if (!formSheet) throw new Error('The PPMP Form sheet was not found.');

  const currentSourceData = loadSourceLookup_(sourceSheet);
  refreshEntryDropdownOptions_(destinationSpreadsheet, currentSourceData);

  // Refresh the separate End-User list too, and restore B9 validation after
  // updating the shared helper sheet.
  const sourceSpreadsheet = SpreadsheetApp.openById(
    ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
  );
  applyEndUserDropdownB9_(destinationSpreadsheet, sourceSpreadsheet);

  const formInfo = findDestinationHeaders_(formSheet);
  if (!formInfo) throw new Error('Could not locate headers in PPMP Form.');

  const descriptionColumn = formInfo.headerMap[
    normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.description)
  ];

  const lastRow = formSheet.getLastRow();

  if (descriptionColumn !== undefined && lastRow >= formInfo.dataStartRow) {
    const range = formSheet.getRange(
      formInfo.dataStartRow,
      descriptionColumn + 1,
      lastRow - formInfo.dataStartRow + 1,
      1
    );

    const descriptions = range.getDisplayValues();
    const lookup = currentSourceData.lookup;
    const optionByBaseDescription = {};

    // Build a fallback index by the source item name before the
    // comma-space that separates it from technical specifications.
    currentSourceData.options.forEach(function(option) {
      const baseName = option.split(', ')[0];
      const key = normalizeValue_(baseName);
      if (!optionByBaseDescription[key]) optionByBaseDescription[key] = [];
      optionByBaseDescription[key].push(option);
    });

    descriptions.forEach(function(row, index) {
      const currentValue = String(row[0] || '').trim();
      if (!currentValue) return;

      let selectedOption = currentValue;
      let record = lookup[normalizeValue_(currentValue)];

      if (!record) {
        const baseName = currentValue.split(', ')[0];
        const matches = optionByBaseDescription[normalizeValue_(baseName)] || [];

        // Only auto-remap a stale selection when the item name matches
        // exactly one current source option.
        if (matches.length === 1) {
          selectedOption = matches[0];
          record = lookup[normalizeValue_(selectedOption)];
          range.getCell(index + 1, 1).setValue(selectedOption);
        }
      }

      if (record) {
        const targetRow = formInfo.dataStartRow + index;

        setIfPresent_(formSheet, targetRow, formInfo.headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.classification, record.classification);
        setIfPresent_(formSheet, targetRow, formInfo.headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.category, record.category);
        setIfPresent_(formSheet, targetRow, formInfo.headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.uom, record.uom);
        setIfPresent_(formSheet, targetRow, formInfo.headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.unitCost, record.unitCost);

        updateEstimatedBudget_(formSheet, targetRow, formInfo.headerMap);
      }
    });
  }

  SpreadsheetApp.flush();
  updatePPMPSummary_(destinationSpreadsheet);

  destinationSpreadsheet.toast(
    'Entries changed: PPMP Form and PPMP Summary have been refreshed.',
    'Procurement Automation',
    6
  );
}


/***************************************************************
 * LOAD SOURCE LOOKUP
 ***************************************************************/

function loadSourceLookup_(sourceSheet) {
  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    throw new Error(
      'The Entries worksheet contains no data.'
    );
  }

  const values = sourceSheet.getRange(
    1, 1, lastRow, lastColumn
  ).getDisplayValues();

  const info = findSourceHeaders_(values);

  if (!info) {
    throw new Error(
      'Could not locate the required source headers in Entries. ' +
      'Required: description, UOM, Unit Cost, Classification and Category.'
    );
  }

  const lookup = {};
  const options = [];
  const seen = new Set();

  for (let r = info.headerRow + 1; r < values.length; r++) {
    const row = values[r];

    const description =
      getArrayValue_(row, info.headerMap.description);

    const specifications =
      getArrayValue_(row, info.headerMap.specifications);

    if (!description) continue;

    // The PPMP Form description must combine the source name and
    // technical specifications using a comma and a space.
    const option = specifications
      ? description + ', ' + specifications
      : description;

    const key = normalizeValue_(option);

    if (seen.has(key)) continue;

    seen.add(key);
    options.push(option);

    lookup[key] = {
      classification:
        getArrayValue_(row, info.headerMap.classification),

      category:
        getArrayValue_(row, info.headerMap.category),

      // The source UOM value is retained for destination autofill.
      uom:
        getArrayValue_(row, info.headerMap.uom),

      unitCost:
        getArrayValue_(row, info.headerMap.unitCost)
    };
  }

  return {
    options: options,
    lookup: lookup
  };
}


/***************************************************************
 * FIND SOURCE HEADERS
 *
 * Supports alternative header names and arbitrary column order.
 * Technical Specifications is optional.
 ***************************************************************/

function findSourceHeaders_(values) {
  const aliases = {
    description: [
      'General Description/Generic Name',
      'General Description',
      'Description',
      'Item Description'
    ],

    specifications: [
      'Technical Specifications',
      'Specifications',
      'Technical Specification'
    ],

    // Prefer specific UOM headers. Do not match the generic
    // header "Unit", which can accidentally select an unrelated column.
    uom: [
      'Unit or Measurement/Size',
      'Unit or Measurement Size',
      'Unit/Measurement/Size',
      'Unit/Measure/Size',
      'Unit of Measurement',
      'Unit of Measure',
      'Measurement/Size',
      'UOM'
    ],

    unitCost: [
      'Unit Cost',
      'Cost per Unit',
      'Price per Unit'
    ],

    classification: [
      'Classification',
      'Type of Project',
      'Type of the Project to be Procured'
    ],

    category: [
      'Category',
      'Item Category'
    ]
  };

  const searchRows = Math.min(
    values.length,
    ENTRY_CONFIG.HEADER_SEARCH_ROWS
  );

  for (let r = 0; r < searchRows; r++) {
    const headerMap = {};

    for (let c = 0; c < values[r].length; c++) {
      const currentHeader =
        normalizeHeader_(values[r][c]);

      if (!currentHeader) continue;

      Object.keys(aliases).forEach(function(key) {
        const normalizedAliases = aliases[key].map(normalizeHeader_);
        const matches = normalizedAliases.indexOf(currentHeader) !== -1;

        if (matches && headerMap[key] === undefined) {
          headerMap[key] = c;
        }
      });
    }

    // Description, UOM, Unit Cost, Classification and Category
    // are required. Specifications is optional.
    const requiredKeys = [
      'description',
      'uom',
      'unitCost',
      'classification',
      'category'
    ];

    const allFound = requiredKeys.every(function(key) {
      return headerMap[key] !== undefined;
    });

    if (allFound) {
      return {
        headerRow: r,
        headerMap: headerMap
      };
    }
  }

  return null;
}


/***************************************************************
 * APPLY RECOMMENDED MODE DROPDOWN
 *
 * Reads Modes from the external source spreadsheet.
 ***************************************************************/

function applyRecommendedModeDropdown_(destinationSpreadsheet) {
  const sourceSpreadsheet =
    SpreadsheetApp.openById(
      ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
    );

  const modesSheet =
    sourceSpreadsheet.getSheetByName(
      ENTRY_CONFIG.MODES_SHEET_NAME
    );

  if (!modesSheet) {
    throw new Error(
      'The Modes worksheet was not found in the external source spreadsheet.'
    );
  }

  const lastRow = modesSheet.getLastRow();

  if (lastRow < ENTRY_CONFIG.MODES_START_ROW) {
    throw new Error(
      'The Modes worksheet contains no procurement modes.'
    );
  }

  const rawModes = modesSheet.getRange(
    ENTRY_CONFIG.MODES_START_ROW,
    ENTRY_CONFIG.MODES_COLUMN,
    lastRow - ENTRY_CONFIG.MODES_START_ROW + 1,
    1
  ).getDisplayValues();

  const modes = [];
  const seen = new Set();

  rawModes.forEach(function(row) {
    const mode = String(row[0] || '').trim();
    if (!mode) return;

    const header = normalizeHeader_(mode);

    if (
      header === 'modeofprocurement' ||
      header === 'modes' ||
      header === 'recommendedmodeofprocurement'
    ) {
      return;
    }

    const key = normalizeValue_(mode);

    if (!seen.has(key)) {
      seen.add(key);
      modes.push(mode);
    }
  });

  if (modes.length === 0) {
    throw new Error(
      'No procurement modes were found in the Modes worksheet.'
    );
  }

  const rule =
    SpreadsheetApp.newDataValidation()
      .requireValueInList(modes, true)
      .setAllowInvalid(false)
      .setHelpText(
        'Select a procurement mode from the Modes worksheet.'
      )
      .build();

  let updated = 0;

  destinationSpreadsheet.getSheets().forEach(function(sheet) {
    if (
      sheet.getName() ===
      ENTRY_CONFIG.HELPER_SHEET_NAME
    ) {
      return;
    }

    const info = findDestinationHeaders_(sheet);
    if (!info) return;

    const column =
      info.headerMap[
        normalizeHeader_(
          ENTRY_CONFIG.DESTINATION_HEADERS.recommendedMode
        )
      ];

    if (column === undefined) return;

    // PPMP Form procurement-mode dropdown validation starts at row 15.
    const validationStartRow = info.dataStartRow;

    if (sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET && sheet.getMaxRows() >= 14) {
      sheet.getRange(1, 1, 14, sheet.getMaxColumns()).clearDataValidations();
    }

    // Match mode validation to the last populated description row.
    const descriptionColumn = info.headerMap[
      normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.description)
    ];
    const validationEndRow =
      sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET &&
      descriptionColumn !== undefined
        ? findLastDescriptionEntryRow_(sheet, descriptionColumn + 1, info.dataStartRow)
        : sheet.getMaxRows();

    const availableRows =
      validationEndRow - validationStartRow + 1;

    if (sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET) {
      const clearFromRow = Math.max(info.dataStartRow, validationEndRow + 1);
      if (clearFromRow <= sheet.getMaxRows()) {
        sheet.getRange(
          clearFromRow, column + 1,
          sheet.getMaxRows() - clearFromRow + 1, 1
        ).clearDataValidations();
      }
    }

    if (availableRows <= 0) return;

    sheet.getRange(
      validationStartRow,
      column + 1,
      availableRows,
      1
    ).setDataValidation(rule);

    // Explicitly clear the reported cells after applying mode validation.
    if (sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET && sheet.getMaxRows() >= 4) {
      sheet.getRange('A4').clearDataValidations();
      sheet.getRange('F4').clearDataValidations();
    }

    updated++;
  });

  return updated;
}


/***************************************************************
 * HANDLE EDITS
 *
 * Autofills fields when a description is selected.
 * Recalculates estimated budget when relevant fields change.
 * Refreshes PPMP Summary after the edit has been processed.
 ***************************************************************/

function handleEntrySelection(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const spreadsheet = e.source;

  if (
    sheet.getName() ===
    ENTRY_CONFIG.HELPER_SHEET_NAME
  ) {
    return;
  }

  const isPPMPForm =
    sheet.getName() ===
    ENTRY_CONFIG.PPMP_FORM_SHEET;

  // The checkbox below the table acts as the visible Add Entry Row control.
  // addPPMPEntryRow clears that control row and converts it into a new item row.
  if (
    isPPMPForm &&
    range.getNumRows() === 1 &&
    range.getNumColumns() === 1 &&
    range.getColumn() === 1 &&
    range.getValue() === true &&
    range.getRow() >= 15
  ) {
    // Clear the checked control first so it is not mistaken for a project
    // description when addPPMPEntryRow locates the next blank line-item row.
    range.clearContent();
    addPPMPEntryRow();
    return;
  }

  const info = findDestinationHeaders_(sheet);

  if (!info) {
    if (isPPMPForm) {
      updatePPMPSummary_(spreadsheet);
    }
    return;
  }

  const headerMap = info.headerMap;

  const descriptionColumn =
    headerMap[
      normalizeHeader_(
        ENTRY_CONFIG.DESTINATION_HEADERS.description
      )
    ];

  // PPMP Form's entry rows begin at row 5, even when the sheet has
  // multi-row headings or a large formatted template.
  const dataStartRow = isPPMPForm ? 5 : info.dataStartRow;

  const unitCostColumn =
    headerMap[
      normalizeHeader_(
        ENTRY_CONFIG.DESTINATION_HEADERS.unitCost
      )
    ];

  const quantityColumn =
    headerMap[
      normalizeHeader_(
        ENTRY_CONFIG.DESTINATION_HEADERS.quantity
      )
    ];

  const firstColumn = range.getColumn();
  const lastColumn =
    firstColumn + range.getNumColumns() - 1;

  const descriptionEdited =
    descriptionColumn !== undefined &&
    descriptionColumn + 1 >= firstColumn &&
    descriptionColumn + 1 <= lastColumn;

  const unitCostEdited =
    unitCostColumn !== undefined &&
    unitCostColumn + 1 >= firstColumn &&
    unitCostColumn + 1 <= lastColumn;

  const quantityEdited =
    quantityColumn !== undefined &&
    quantityColumn + 1 >= firstColumn &&
    quantityColumn + 1 <= lastColumn;

  if (
    !descriptionEdited &&
    !unitCostEdited &&
    !quantityEdited
  ) {
    if (isPPMPForm) {
      updatePPMPSummary_(spreadsheet);
    }
    return;
  }

  let sourceLookup = null;

  if (descriptionEdited) {
    const sourceSpreadsheet =
      SpreadsheetApp.openById(
        ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
      );

    const sourceSheet =
      sourceSpreadsheet.getSheetByName(
        ENTRY_CONFIG.SOURCE_SHEET_NAME
      );

    if (!sourceSheet) {
      throw new Error(
        'The Entries source worksheet was not found.'
      );
    }

    sourceLookup =
      loadSourceLookup_(sourceSheet).lookup;
  }

  const firstRow = Math.max(
    range.getRow(),
    dataStartRow
  );

  const lastRow =
    range.getRow() + range.getNumRows() - 1;

  for (let row = firstRow; row <= lastRow; row++) {
    if (descriptionEdited) {
      const description = String(
        sheet.getRange(
          row,
          descriptionColumn + 1
        ).getDisplayValue() || ''
      ).trim();

      if (!description) {
        clearRelatedFields_(sheet, row, info);
        continue;
      }

      const record =
        sourceLookup[normalizeValue_(description)];

      if (record) {
        setIfPresent_(
          sheet,
          row,
          headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.classification,
          record.classification
        );

        setIfPresent_(
          sheet,
          row,
          headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.category,
          record.category
        );

        // The source Entries column UOM fills the PPMP Form's
        // "Unit or Measurement/Size" field.
        setIfPresent_(
          sheet,
          row,
          headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.uom,
          record.uom
        );

        setIfPresent_(
          sheet,
          row,
          headerMap,
          ENTRY_CONFIG.DESTINATION_HEADERS.unitCost,
          record.unitCost
        );
      } else {
        // If the selected description no longer exists in Entries,
        // remove stale autofilled values instead of keeping old data.
        clearRelatedFields_(sheet, row, info);
      }
    }

    if (
      descriptionEdited ||
      unitCostEdited ||
      quantityEdited
    ) {
      updateEstimatedBudget_(
        sheet,
        row,
        headerMap
      );
    }
  }

  // Refresh the summary AFTER autofill and calculations.
  if (isPPMPForm) {
    SpreadsheetApp.flush();
    updatePPMPSummary_(spreadsheet);
    createAddEntryRowControl_(sheet);
  }
}


/***************************************************************
 * CLEAR RELATED FIELDS
 ***************************************************************/

function clearRelatedFields_(sheet, row, info) {
  const headers = [
    ENTRY_CONFIG.DESTINATION_HEADERS.classification,
    ENTRY_CONFIG.DESTINATION_HEADERS.category,
    ENTRY_CONFIG.DESTINATION_HEADERS.uom,
    ENTRY_CONFIG.DESTINATION_HEADERS.unitCost,
    ENTRY_CONFIG.DESTINATION_HEADERS.quantity,
    ENTRY_CONFIG.DESTINATION_HEADERS.estimatedBudget,
    ENTRY_CONFIG.DESTINATION_HEADERS.recommendedMode
  ];

  headers.forEach(function(headerName) {
    const column =
      info.headerMap[normalizeHeader_(headerName)];

    if (column !== undefined) {
      sheet.getRange(
        row, column + 1
      ).clearContent();
    }
  });
}


/***************************************************************
 * UPDATE ESTIMATED BUDGET
 *
 * Estimated Budget = Unit Cost x Quantity
 ***************************************************************/

function updateEstimatedBudget_(sheet, row, headerMap) {
  const unitCostColumn =
    headerMap[
      normalizeHeader_(
        ENTRY_CONFIG.DESTINATION_HEADERS.unitCost
      )
    ];

  const quantityColumn =
    headerMap[
      normalizeHeader_(
        ENTRY_CONFIG.DESTINATION_HEADERS.quantity
      )
    ];

  const budgetColumn =
    headerMap[
      normalizeHeader_(
        ENTRY_CONFIG.DESTINATION_HEADERS.estimatedBudget
      )
    ];

  if (
    unitCostColumn === undefined ||
    quantityColumn === undefined ||
    budgetColumn === undefined
  ) {
    return;
  }

  const unitCost =
    sheet.getRange(
      row, unitCostColumn + 1
    ).getValue();

  const quantity =
    sheet.getRange(
      row, quantityColumn + 1
    ).getValue();

  const budgetCell =
    sheet.getRange(
      row, budgetColumn + 1
    );

  if (
    unitCost === '' ||
    unitCost === null ||
    quantity === '' ||
    quantity === null
  ) {
    budgetCell.clearContent();
    return;
  }

  const cost = parseNumericValue_(unitCost);
  const qty = parseNumericValue_(quantity);

  if (cost === null || qty === null) {
    budgetCell.clearContent();
    return;
  }

  budgetCell.setValue(cost * qty);
}


/***************************************************************
 * AUTOMATIC PPMP SUMMARY
 *
 * Groups repeated categories and totals their amounts.
 ***************************************************************/

function updatePPMPSummary_(spreadsheet) {
  spreadsheet =
    spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

  const formSheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  const summarySheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_SUMMARY_SHEET);

  if (!formSheet) throw new Error('The PPMP Form sheet was not found.');
  if (!summarySheet) throw new Error('The PPMP Summary sheet was not found.');

  // Keep Fiscal Year and End-User synchronized with the corresponding
  // values in PPMP Form, using the labels on each sheet to locate the cells.
  syncPPMPSummaryMetadata_(formSheet, summarySheet);

  const isPPMPForm = formSheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET;
  const formInfo = isPPMPForm
    ? { headerRow: 13, categoryColumn: 3, amountColumn: 13 }
    : (findCategoryAmountHeaders_(formSheet) ||
       findCategoryAmountHeadersFallback_(formSheet));

  if (!formInfo) {
    throw new Error('Could not identify the Category and Estimated Budget columns in PPMP Form.');
  }

  const summaryInfo =
    findCategoryAmountHeaders_(summarySheet) ||
    findCategoryAmountHeadersFallback_(summarySheet);

  if (!summaryInfo) {
    throw new Error(
      'PPMP Summary must contain a Category header and a Total Amount ' +
      'header (or Total ABC/Amount) within its first 10 rows.'
    );
  }

  const totals = new Map();
  // The PPMP Form line-item table starts at row 15. Require a description
  // in column A so stray template text such as "Category" or "Column 3"
  // is never treated as an item/category record.
  const firstDataRow = isPPMPForm ? 15 : 5;
  const lastDataRow = formSheet.getLastRow();

  if (lastDataRow >= firstDataRow) {
    const sourceStartColumn = isPPMPForm ? 1 : formInfo.categoryColumn;
    const sourceEndColumn = isPPMPForm ? 13 : formInfo.amountColumn;
    const data = formSheet.getRange(
      firstDataRow,
      sourceStartColumn,
      lastDataRow - firstDataRow + 1,
      sourceEndColumn - sourceStartColumn + 1
    ).getValues();

    data.forEach(function(row) {
      const descriptionIndex = 0;
      const categoryIndex = isPPMPForm ? 2 : formInfo.categoryColumn - sourceStartColumn;
      const amountIndex = isPPMPForm ? 12 : formInfo.amountColumn - sourceStartColumn;
      const description = String(row[descriptionIndex] || '').trim();
      const category = String(row[categoryIndex] || '').trim();

      if (!description || !category) return;

      const normalizedCategory = normalizeHeader_(category);
      if (normalizedCategory === 'category' || normalizedCategory === 'column3') return;

      const amount = parseNumericValue_(row[amountIndex]);
      const key = normalizeValue_(category);

      if (!totals.has(key)) {
        totals.set(key, { category: category, total: 0 });
      }
      if (amount !== null) totals.get(key).total += amount;
    });
  }

  const results = Array.from(totals.values()).sort(function(a, b) {
    return a.category.localeCompare(b.category);
  });

  const summaryHeaderRow = Math.max(
    summaryInfo.headerRow || 0,
    summaryInfo.amountHeaderRow || 0
  );
  const outputStartRow = summaryHeaderRow + 1;
  const availableRows = summarySheet.getMaxRows() - summaryHeaderRow;

  if (availableRows > 0) {
    // Clear stale summary output, while leaving metadata above the table intact.
    summarySheet.getRange(
      outputStartRow, 1, availableRows, summarySheet.getMaxColumns()
    ).clearContent();
  }

  if (results.length === 0) return;

  // Reserve one additional row after the category breakdown for GRAND TOTAL.
  const grandTotal = results.reduce(function(sum, item) {
    return sum + item.total;
  }, 0);
  const grandTotalRow = outputStartRow + results.length;
  const requiredLastRow = grandTotalRow;
  if (requiredLastRow > summarySheet.getMaxRows()) {
    summarySheet.insertRowsAfter(
      summarySheet.getMaxRows(),
      requiredLastRow - summarySheet.getMaxRows()
    );
  }

  if (results.length > 0) {
    const categoryOutputRange = summarySheet.getRange(
      outputStartRow, summaryInfo.categoryColumn, results.length, 1
    );
    categoryOutputRange
      .setValues(results.map(function(item) { return [item.category]; }))
      .setFontSize(11)
      .setFontWeight('normal');

    const amountOutputRange = summarySheet.getRange(
      outputStartRow, summaryInfo.amountColumn, results.length, 1
    );
    amountOutputRange
      .setValues(results.map(function(item) { return [item.total]; }))
      .setNumberFormat('#,##0.00')
      .setFontSize(11)
      .setFontWeight('normal');
  }

  // Put GRAND TOTAL immediately below the last category and its sum in the
  // adjacent Total Amount column. Clear-and-rebuild above prevents stale totals.
  summarySheet.getRange(grandTotalRow, summaryInfo.categoryColumn)
    .setValue('GRAND TOTAL')
    .setFontSize(14)
    .setFontWeight('bold');
  summarySheet.getRange(grandTotalRow, summaryInfo.amountColumn)
    .setValue(grandTotal)
    .setNumberFormat('#,##0.00')
    .setFontSize(14)
    .setFontWeight('bold');
}


/**
 * Copies Fiscal Year and End-User values from PPMP Form to the value
 * cells beside the matching labels in PPMP Summary. The labels may be
 * positioned differently on the two sheets.
 */
function syncPPMPSummaryMetadata_(formSheet, summarySheet) {
  // Source fields in PPMP Form: Fiscal Year = B8, End-User = B9.
  // Destination fields in PPMP Summary: Fiscal Year = B7, End-User = B8.
  // Map each source cell to its required destination cell explicitly.
  [
    { source: 'B8', destination: 'B7' },
    { source: 'B9', destination: 'B8' }
  ].forEach(function(field) {
    summarySheet.getRange(field.destination).setValue(
      formSheet.getRange(field.source).getValue()
    );
  });
}

/**
 * Fallback header finder for templates with merged or multi-line headings.
 * Returns 1-based column indexes, matching updatePPMPSummary_.
 */
function findCategoryAmountHeadersFallback_(sheet) {
  const rowsToSearch = Math.min(sheet.getMaxRows(), 15);
  const columnsToSearch = sheet.getLastColumn();
  if (!rowsToSearch || !columnsToSearch) return null;

  const values = sheet.getRange(
    1, 1, rowsToSearch, columnsToSearch
  ).getDisplayValues();

  const categoryAliases = ['category', 'item category'];
  const amountAliases = [
    normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.estimatedBudget),
    'estimatedbudget',
    'estimatedbudgetph',
    'totalamount',
    'totalabc',
    'amount',
    'abc'
  ];

  let categoryColumn = null;
  let categoryHeaderRow = null;
  let amountColumn = null;
  let amountHeaderRow = null;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const header = normalizeHeader_(values[r][c]);
      if (!header) continue;

      if (categoryColumn === null && categoryAliases.indexOf(header) !== -1) {
        categoryColumn = c + 1;
        categoryHeaderRow = r + 1;
      }

      if (amountColumn === null && amountAliases.indexOf(header) !== -1) {
        amountColumn = c + 1;
        amountHeaderRow = r + 1;
      }
    }
  }

  if (categoryColumn === null || amountColumn === null) return null;

  return {
    headerRow: categoryHeaderRow,
    categoryColumn: categoryColumn,
    amountColumn: amountColumn,
    amountHeaderRow: amountHeaderRow
  };
}


/***************************************************************
 * FIND CATEGORY AND AMOUNT HEADERS
 ***************************************************************/

function findCategoryAmountHeaders_(sheet) {
  const lastRow = Math.min(
    sheet.getLastRow(),
    ENTRY_CONFIG.HEADER_SEARCH_ROWS
  );

  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    return null;
  }

  const values = sheet.getRange(
    1, 1, lastRow, lastColumn
  ).getDisplayValues();

  const categoryHeader = normalizeHeader_('Category');

  // Prefer the actual line-item amount column on PPMP Form.
  // Summary sheets can use Total Amount or similar headings.
  // Use the PPMP Form's line-item estimated budget as the source
  // amount. In PPMP Summary, write totals specifically to Total Amount
  // when that header exists, rather than accidentally choosing an
  // unrelated Estimated Budget column.
  const isSummarySheet =
    sheet.getName() === ENTRY_CONFIG.PPMP_SUMMARY_SHEET;

  const preferredAmountHeaders = isSummarySheet
    ? [
        normalizeHeader_('Total Amount'),
        normalizeHeader_('Total ABC'),
        normalizeHeader_('Amount'),
        normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.estimatedBudget),
        normalizeHeader_('Estimated Budget')
      ]
    : [
        normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.estimatedBudget),
        normalizeHeader_('Estimated Budget'),
        normalizeHeader_('Total Amount'),
        normalizeHeader_('Total ABC'),
        normalizeHeader_('Amount')
      ];

  let categoryColumn = null;
  let categoryHeaderRow = null;
  const amountColumnsByHeader = {};

  // Scan the whole header area instead of requiring Category and
  // the amount header to be on the same row. PPMP Form often uses
  // multi-row and merged headers.
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const header = normalizeHeader_(values[r][c]);
      if (!header) continue;

      if (header === categoryHeader && categoryColumn === null) {
        categoryColumn = c + 1;
        categoryHeaderRow = r + 1;
      }

      if (
        preferredAmountHeaders.indexOf(header) !== -1 &&
        amountColumnsByHeader[header] === undefined
      ) {
        amountColumnsByHeader[header] = {
          column: c + 1,
          row: r + 1
        };
      }
    }
  }

  if (!categoryColumn) return null;

  let amountInfo = null;
  for (let i = 0; i < preferredAmountHeaders.length; i++) {
    const candidate = amountColumnsByHeader[preferredAmountHeaders[i]];
    if (candidate) {
      amountInfo = candidate;
      break;
    }
  }

  if (!amountInfo) return null;

  return {
    // Keep the category header row as the output anchor for PPMP Summary.
    headerRow: categoryHeaderRow,
    categoryColumn: categoryColumn,
    amountColumn: amountInfo.column,
    amountHeaderRow: amountInfo.row
  };
}

/***************************************************************
 * FIND DESTINATION HEADERS
 ***************************************************************/

function findDestinationHeaders_(sheet) {
  // The ZCMC PPMP Form template has a fixed 15-column layout and its
  // first data row is 15 (the field labels occupy rows above it).
  // Use this explicit mapping so merged headings and repeated "Column 1"
  // labels cannot confuse autofill column detection.
  if (sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET) {
    return {
      headerRow: 13,
      dataStartRow: 15,
      headerMap: {
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.description)]: 0,       // A
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.classification)]: 1,   // B
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.category)]: 2,         // C
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.quantity)]: 3,         // D
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.uom)]: 4,              // E
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.recommendedMode)]: 5,  // F
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.unitCost)]: 11,        // L
        [normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.estimatedBudget)]: 12  // M
      }
    };
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    return null;
  }

  const rowsToSearch = Math.min(
    lastRow,
    ENTRY_CONFIG.HEADER_SEARCH_ROWS
  );

  const values = sheet.getRange(
    1, 1, rowsToSearch, lastColumn
  ).getDisplayValues();

  const descriptionHeader = normalizeHeader_(
    ENTRY_CONFIG.DESTINATION_HEADERS.description
  );

  const uomHeader = normalizeHeader_(
    ENTRY_CONFIG.DESTINATION_HEADERS.uom
  );

  const uomAliases = [
    'uom',
    'unit',
    'unit of measurement',
    'unit of measure',
    'unit/measurement/size',
    'measurement/size',
    'unit or measurement/size',
    'unit or measurement size',
    'unit/measure/size',
    'measurement',
    'size'
  ].map(normalizeHeader_);

  // PPMP Form can use a two-row header: the main description and
  // other headings may be on one row, while "Unit or Measurement/Size"
  // appears on the next row beneath the merged "Quantity and Size..."
  // parent heading. Collect headers across the searched rows, preserving
  // the first match for each field so the subheader is not missed.
  const headerMap = {};
  let headerRow = null;
  let lastHeaderRow = 0;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const normalized = normalizeHeader_(values[r][c]);
      if (!normalized) continue;

      if (
        normalized === descriptionHeader &&
        headerRow === null
      ) {
        headerRow = r + 1;
      }

      Object.keys(
        ENTRY_CONFIG.DESTINATION_HEADERS
      ).forEach(function(key) {
        const expected = normalizeHeader_(
          ENTRY_CONFIG.DESTINATION_HEADERS[key]
        );

        if (normalized === expected) {
          lastHeaderRow = Math.max(lastHeaderRow, r + 1);
          if (headerMap[expected] === undefined) {
            headerMap[expected] = c;
          }
        }
      });

      const looksLikeUomHeader =
        uomAliases.indexOf(normalized) !== -1 ||
        (
          normalized.indexOf('unit') !== -1 &&
          (
            normalized.indexOf('measurement') !== -1 ||
            normalized.indexOf('measure') !== -1 ||
            normalized.indexOf('size') !== -1
          )
        );

      if (looksLikeUomHeader) {
        lastHeaderRow = Math.max(lastHeaderRow, r + 1);
        if (headerMap[uomHeader] === undefined) {
          headerMap[uomHeader] = c;
        }
      }
    }
  }

  if (headerRow === null) {
    return null;
  }

  // PPMP Form data-entry rows begin at row 5. Keep validation and
  // edit handling aligned to that fixed first entry row, regardless
  // of whether the sheet uses one-row or multi-row headers.
  const dataStartRow =
    sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET
      ? 5
      : Math.max(headerRow, lastHeaderRow) + 1;

  return {
    headerRow: headerRow,
    dataStartRow: dataStartRow,
    headerMap: headerMap
  };
}

/***************************************************************
 * SET DESTINATION CELL IF HEADER EXISTS
 ***************************************************************/

function setIfPresent_(
  sheet,
  row,
  headerMap,
  headerName,
  value
) {
  const column =
    headerMap[normalizeHeader_(headerName)];

  if (column === undefined) return;

  const cell = sheet.getRange(
    row, column + 1
  );

  if (value === null || value === undefined) {
    cell.clearContent();
  } else {
    cell.setValue(value);
  }
}


/***************************************************************
 * GET ARRAY VALUE
 ***************************************************************/

function getArrayValue_(row, columnIndex) {
  if (
    columnIndex === undefined ||
    columnIndex < 0
  ) {
    return '';
  }

  return String(
    row[columnIndex] === null ||
    row[columnIndex] === undefined
      ? ''
      : row[columnIndex]
  ).trim();
}


/***************************************************************
 * NORMALIZE HEADER
 *
 * Ignores capitalization, spaces, punctuation and slashes.
 ***************************************************************/

function normalizeHeader_(value) {
  return String(
    value === null || value === undefined
      ? ''
      : value
  )
    .toLowerCase()
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[^a-z0-9]/g, '');
}


/***************************************************************
 * NORMALIZE LOOKUP VALUE
 ***************************************************************/

function normalizeValue_(value) {
  return String(
    value === null || value === undefined
      ? ''
      : value
  )
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}


/***************************************************************
 * PARSE NUMERIC VALUE
 ***************************************************************/

function parseNumericValue_(value) {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return null;
  }

  if (typeof value === 'number') {
    return isFinite(value) ? value : null;
  }

  const cleaned = String(value)
    .trim()
    .replace(/,/g, '')
    .replace(/[₱$]/g, '')
    .replace(/\s/g, '');

  if (!cleaned) return null;

  const number = Number(cleaned);

  return isFinite(number) ? number : null;
}