
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
  // Intentionally do not add a custom menu to the spreadsheet UI.
  // Installable edit triggers handle user edits automatically.
}

/**
 * Recreate the installable edit trigger for the currently active destination
 * spreadsheet. Run from the bound Apps Script project attached to PPMP Form.
 */
function setupPPMPAutofillTriggerOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'Open the destination spreadsheet that contains the PPMP Form tab, then run this command again.'
    );
  }

  const targetSheet = ss.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (!targetSheet) {
    throw new Error(
      'This spreadsheet does not contain a tab named "' +
      ENTRY_CONFIG.PPMP_FORM_SHEET +
      '". Open the correct destination spreadsheet before enabling automatic autofill.'
    );
  }

  // Remove older edit handlers as well as duplicates so only one
  // destination-sheet edit trigger processes each selection.
  const legacyHandlers = [
    'handlePPMPFormEdit',
    'autoSuggestEntriesOnEdit',
    'handleEntrySelection'
  ];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (legacyHandlers.indexOf(trigger.getHandlerFunction()) !== -1 &&
        trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // This installable trigger is essential: autofill reads the separate
  // Entries spreadsheet, which a simple onEdit trigger cannot reliably access.
  ScriptApp.newTrigger('handlePPMPFormEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  const installed = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === 'handlePPMPFormEdit' &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
      trigger.getTriggerSourceId() === ss.getId();
  });

  if (!installed) {
    throw new Error(
      'Google Apps Script did not confirm the edit trigger. Run this command again and approve the requested permissions.'
    );
  }

  ss.toast(
    'Automatic PPMP autofill is enabled for this spreadsheet. Choose a General Description in PPMP Form, column A, row 15 or later.',
    'PPMP Form',
    10
  );
}

/**
 * Manual diagnostic/fallback: select any cell on the PPMP Form row to test,
 * then use PPMP Form > Test / Repair Selected PPMP Row.
 */
function testRepairSelectedPPMPRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the destination spreadsheet first.');

  // Always resolve the intended worksheet by its configured name. Do not
  // reject the test just because Apps Script reports another active tab.
  const sheet = ss.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (!sheet) {
    const names = ss.getSheets().map(function (s) {
      return '"' + s.getName() + '"';
    }).join(', ');
    throw new Error(
      'Could not find the worksheet named "' +
      ENTRY_CONFIG.PPMP_FORM_SHEET + '". Available sheets: ' + names
    );
  }

  let activeRange = null;
  const activeSheet = ss.getActiveSheet();
  if (activeSheet && activeSheet.getSheetId() === sheet.getSheetId()) {
    activeRange = sheet.getActiveRange();
  } else {
    // Some menu/editor execution contexts do not report the expected tab as
    // active. Ask the user to select the intended row, then rerun.
    ss.setActiveSheet(sheet);
    activeRange = sheet.getActiveRange();
  }

  if (!activeRange) {
    throw new Error(
      'The PPMP Form sheet was found. Select a cell in the desired line-item row (row 15 or later), then run the test again.'
    );
  }

  const row = activeRange.getRow();
  if (row < 15) {
    throw new Error('Select a line-item row starting at row 15.');
  }

  const selected = String(
    sheet.getRange(row, 1).getDisplayValue() || ''
  ).trim();

  if (!selected) {
    const message = 'PPMP Form!A' + row +
      ' is blank. Select a General Description first.';
    ss.toast(message, 'PPMP Autofill Diagnostic', 10);
    throw new Error(message);
  }

  const sourceSS = SpreadsheetApp.openById(
    ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
  );
  const sourceSheet = sourceSS.getSheetByName(
    ENTRY_CONFIG.SOURCE_SHEET_NAME
  );

  if (!sourceSheet) {
    throw new Error('The source Entries worksheet was not found.');
  }

  const sourceData = loadSourceLookup_(sourceSheet);
  const record = findMatchingEntryRecord_(
    selected,
    sourceData.lookup
  );

  if (!record) {
    const message = 'No matching source record for A' + row +
      '. Selected value: "' + selected +
      '". Source options loaded: ' + sourceData.options.length +
      '. Examples: ' + sourceData.options.slice(0, 3).join(' | ');
    ss.toast(message, 'PPMP Autofill Diagnostic', 12);
    throw new Error(message);
  }

  sheet.getRange(row, 2).setValue(record.classification || '');
  sheet.getRange(row, 3).setValue(record.category || '');
  sheet.getRange(row, 5).setValue(record.uom || '');

  const cost = parseNumericValue_(record.unitCost);
  if (cost === null) {
    sheet.getRange(row, 12).clearContent();
  } else {
    sheet.getRange(row, 12).setValue(cost);
  }

  const qty = parseNumericValue_(
    sheet.getRange(row, 4).getValue()
  );

  if (cost === null || qty === null) {
    sheet.getRange(row, 13).clearContent();
  } else {
    sheet.getRange(row, 13).setValue(cost * qty);
  }

  SpreadsheetApp.flush();

  ss.toast(
    'Source match found for row ' + row +
    '. Check B, C, E, L and M.',
    'PPMP Autofill Diagnostic',
    8
  );
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

  const ppmpFormSheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (!ppmpFormSheet) throw new Error('The PPMP Form sheet was not found.');
  initializePPMPEntryTable_(ppmpFormSheet);
  removeLegacyAddEntryRowLabel_(ppmpFormSheet);

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
      handler === 'autoSuggestEntriesOnEdit' ||
      handler === 'handlePPMPFormEdit' ||
      handler === 'handleEntriesSourceEdit' ||
      handler === 'handlePPMPRowChange'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Recreate the installable PPMP Form edit trigger. It is required because
  // autofill reads the separate source spreadsheet with SpreadsheetApp.openById().
  ScriptApp.newTrigger('handlePPMPFormEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  ScriptApp.newTrigger('handleEntriesSourceEdit')
    .forSpreadsheet(sourceSpreadsheet)
    .onEdit()
    .create();

  // Watch manually inserted/deleted sheet rows and rebuild the PPMP table grid.
  ScriptApp.newTrigger('handlePPMPRowChange')
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();

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

  // Extend the visible table grid after formatting and clearing the new row.
  // Header-map indexes are zero-based, so convert the final table column to
  // a Sheets column count by adding one.
  const tableLastColumn = Math.max(
    descriptionColumn,
    modeColumn,
    info.headerMap[normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.classification)] ?? 0,
    info.headerMap[normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.category)] ?? 0,
    info.headerMap[normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.quantity)] ?? 0,
    info.headerMap[normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.uom)] ?? 0,
    info.headerMap[normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.unitCost)] ?? 0,
    info.headerMap[normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.estimatedBudget)] ?? 0
  ) + 1;

  // Match the borders of the previous item row first, then force a complete
  // solid grid on the new row so the outer and inner lines remain visible.
  if (newRow > firstEntryRow) {
    sheet.getRange(newRow - 1, 1, 1, tableLastColumn)
      .copyTo(
        sheet.getRange(newRow, 1, 1, tableLastColumn),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
  }
  sheet.getRange(newRow, 1, 1, tableLastColumn).setBorder(
    true, true, true, true, true, true,
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID
  );

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

  // FINAL BORDER PASS: run only after formatting, clearing, and dropdown setup.
  // Always border the complete A:O table, including the 20 initial rows and
  // every row through the newly prepared entry. Then explicitly re-border the
  // new row itself so its four outer edges and all cell separators are visible.
  const borderThroughRow = Math.max(34, newRow);
  applyPPMPEntryTableBorders_(sheet, borderThroughRow);
  sheet.getRange(newRow, 1, 1, 15).setBorder(
    true, true, true, true, true, true,
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID
  );
  SpreadsheetApp.flush();

  spreadsheet.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange(newRow, descriptionColumn + 1));
  spreadsheet.toast(
    'Added entry row ' + newRow + ' with dropdowns and automatic autofill support.',
    'PPMP Form', 5
  );
  return newRow;
}


/**
 * Applies a solid black cell grid across the PPMP Form table columns A:O.
 * The whole-row operation is safe for merged cells and sets horizontal and
 * vertical borders after the row's content, formatting, and validations settle.
 */
function applyPPMPEntryRowBorders_(sheet, row) {
  if (!sheet || row < 15) return;
  const tableRange = sheet.getRange(row, 1, 1, 15); // A:O
  tableRange.setBorder(
    true, true, true, true, true, true,
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID
  );

  // Reapply the row's bottom border as a clear horizontal separator.
  tableRange.setBorder(
    null, null, true, null, null, null,
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID
  );
}


/**
 * Reapply the complete PPMP Form table grid from row 15 through lastRow.
 * Applying borders to a multi-row A:O range creates both internal vertical
 * separators and horizontal row lines; then each row gets an explicit edge pass.
 */
/**
 * Ensures the PPMP Form starts with 20 visible entry rows (15–34).
 * Copies only formatting and validation from row 15 to the remaining
 * initial rows, so existing entries are never overwritten.
 */
function initializePPMPEntryTable_(sheet) {
  if (!sheet) throw new Error('PPMP Form sheet is required.');
  const firstRow = 15;
  const initialLastRow = 34;
  const tableColumns = 15; // A:O

  if (sheet.getMaxRows() < initialLastRow) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      initialLastRow - sheet.getMaxRows()
    );
  }

  const template = sheet.getRange(firstRow, 1, 1, tableColumns);
  for (let row = firstRow + 1; row <= initialLastRow; row++) {
    const target = sheet.getRange(row, 1, 1, tableColumns);
    template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }

  applyPPMPEntryTableBorders_(sheet, initialLastRow);
}

/**
 * Reapply the complete PPMP Form table grid from row 15 through at least row 34.
 * Applying borders to a multi-row A:O range creates both internal vertical
 * separators and horizontal row lines; then each row gets an explicit edge pass.
 */
function applyPPMPEntryTableBorders_(sheet, lastRow) {
  if (!sheet) throw new Error('PPMP Form sheet is required to apply table borders.');
  const firstRow = 15;
  const initialLastRow = 34; // 20 initial entry rows: 15 through 34
  const finalRow = Math.max(initialLastRow, Number(lastRow) || firstRow);
  if (finalRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), finalRow - sheet.getMaxRows());
  }

  // First set the grid on the entire table range so every cell has
  // solid vertical separators and every row has horizontal separators.
  sheet.getRange(firstRow, 1, finalRow - firstRow + 1, 15).setBorder(
    true, true, true, true, true, true,
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID
  );

  // Then force borders row-by-row. This is important for newly created rows
  // whose formatting was copied from an unbordered or partially bordered row.
  for (let row = firstRow; row <= finalRow; row++) {
    sheet.getRange(row, 1, 1, 15).setBorder(
      true, true, true, true, true, true,
      '#000000',
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
  SpreadsheetApp.flush();
}

/**
 * Removes the old ＋ ADD ENTRY ROW label that may already exist in the
 * spreadsheet from an earlier version of this script. Removing the code
 * that created it does not erase the cell value that was previously written.
 */
function removeLegacyAddEntryRowLabel_(sheet) {
  if (!sheet) return 0;
  const usedRange = sheet.getDataRange();
  const values = usedRange.getDisplayValues();
  let removed = 0;

  for (let r = 0; r < values.length; r++) {
    for (let col = 0; col < values[r].length; col++) {
      const value = String(values[r][col] || '').trim();
      if (/^＋?\s*ADD ENTRY ROW$/i.test(value)) {
        const cell = sheet.getRange(r + 1, col + 1);
        cell.clearContent();
        cell.setBackground(null);
        cell.setFontWeight('normal');
        cell.clearNote();
        removed++;
      }
    }
  }
  SpreadsheetApp.flush();
  return removed;
}

/**
 * Run this function once to remove the old label immediately.
 */
function removeAddEntryRowLabel() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Open the destination spreadsheet first.');
  const sheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (!sheet) throw new Error('The PPMP Form sheet was not found.');
  const removed = removeLegacyAddEntryRowLabel_(sheet);
  spreadsheet.toast('Removed ' + removed + ' old Add Entry Row label(s).', 'PPMP Form', 5);
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
        ? Math.max(34, findLastDescriptionEntryRow_(sheet, descriptionColumn + 1, info.dataStartRow))
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
      // Retain source identity fields so selections can still be matched
      // when Sheets returns only the item name instead of the full dropdown label.
      description: description,
      specifications: specifications,
      classification:
        getArrayValue_(row, info.headerMap.classification),

      category:
        getArrayValue_(row, info.headerMap.category),

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
    // Prefer the source field explicitly requested for the PPMP description.
    // If both Item Description and a pre-combined General Description exist,
    // build the dropdown from Item Description + Technical Specifications.
    description: [
      'Item Description', 'Item Name', 'Generic Name',
      'General Description/Generic Name', 'General Description',
      'General Description Generic Name', 'Description', 'Project Description'
    ],
    specifications: [
      'Technical Specifications', 'Technical Specification',
      'Specifications', 'Specification', 'Technical Description'
    ],
    // Use specific UOM headers only. A generic header such as "Unit"
    // can accidentally match an unrelated column before the actual UOM column.
    uom: [
      'UOM',
      'Unit of Measurement',
      'Unit of Measure',
      'Unit or Measurement/Size',
      'Unit or Measurement Size',
      'Unit/Measurement/Size',
      'Unit/Measure/Size',
      'Measurement/Size'
    ],
    unitCost: [
      'Unit Cost', 'Cost per Unit', 'Price per Unit',
      'Unit Price', 'Estimated Unit Cost', 'Cost'
    ],
    classification: [
      'Classification', 'Type of Project',
      'Type of the Project to be Procured', 'Project Type',
      'Type', 'Item Classification'
    ],
    category: ['Category', 'Item Category', 'Item Classification Category']
  };

  const requiredKeys = ['description', 'uom', 'unitCost', 'classification', 'category'];
  const searchRows = Math.min(values.length, ENTRY_CONFIG.HEADER_SEARCH_ROWS);
  let best = null;

  // Score every candidate header row. A stray title or partial header row
  // must not win over a row containing all the required fields.
  for (let r = 0; r < searchRows; r++) {
    const headerMap = {};
    for (let col = 0; col < values[r].length; col++) {
      const current = normalizeHeader_(values[r][col]);
      if (!current) continue;

      Object.keys(aliases).forEach(function(key) {
        if (headerMap[key] !== undefined) return;
        const aliasKeys = aliases[key].map(normalizeHeader_);
        if (aliasKeys.indexOf(current) !== -1) headerMap[key] = col;
      });
    }

    const foundCount = requiredKeys.filter(function(key) {
      return headerMap[key] !== undefined;
    }).length;

    if (!best || foundCount > best.foundCount) {
      best = { headerRow: r, headerMap: headerMap, foundCount: foundCount };
    }
    if (foundCount === requiredKeys.length) {
      return { headerRow: r, headerMap: headerMap };
    }
  }

  // Some versions of the source sheet have a split/merged heading area.
  // Collect each field's column from the strongest rows as a fallback.
  if (best && best.foundCount >= 3) {
    const combined = Object.assign({}, best.headerMap);
    for (let r = 0; r < searchRows; r++) {
      for (let col = 0; col < values[r].length; col++) {
        const current = normalizeHeader_(values[r][col]);
        if (!current) continue;
        Object.keys(aliases).forEach(function(key) {
          if (combined[key] !== undefined) return;
          if (aliases[key].some(function(alias) {
            return normalizeHeader_(alias) === current;
          })) combined[key] = col;
        });
      }
    }
    if (requiredKeys.every(function(key) { return combined[key] !== undefined; })) {
      return { headerRow: best.headerRow, headerMap: combined };
    }
  }

  return null;
}

/**
 * Match the selected dropdown label to the source Entries record.
 * Uses exact full label first, then normalized label/name comparisons.
 * Ambiguous base-name matches are rejected rather than filling the wrong item.
 */
function findMatchingEntryRecord_(selectedValue, lookup) {
  const selected = normalizeValue_(selectedValue);
  if (!selected) return null;

  // Exact lookup is fastest; normalizeValue_ also handles case/spacing.
  if (lookup[selected]) return lookup[selected];

  const keys = Object.keys(lookup);
  const exactMatches = [];
  const prefixMatches = [];
  const nameMatches = [];

  keys.forEach(function(key) {
    const record = lookup[key];
    const name = normalizeValue_(record.description);
    const specs = normalizeValue_(record.specifications);
    const fullComma = normalizeValue_(record.specifications
      ? record.description + ', ' + record.specifications
      : record.description);
    const fullSpace = normalizeValue_(record.specifications
      ? record.description + ' ' + record.specifications
      : record.description);

    if (
      selected === fullComma ||
      selected === fullSpace ||
      selected === name ||
      normalizeHeader_(selected) === normalizeHeader_(fullComma) ||
      normalizeHeader_(selected) === normalizeHeader_(fullSpace) ||
      normalizeHeader_(selected) === normalizeHeader_(name)
    ) {
      exactMatches.push(record);
      return;
    }

    // Dropdown labels may be displayed with different punctuation or
    // whitespace; accept a name prefix only when it identifies one record.
    if (name && (selected.indexOf(name + ',') === 0 ||
                 selected.indexOf(name + ' ') === 0 ||
                 selected.indexOf(name + '-') === 0)) {
      prefixMatches.push(record);
      return;
    }

    if (specs && selected.indexOf(specs) !== -1 &&
        selected.indexOf(name) === 0) {
      exactMatches.push(record);
      return;
    }

    if (name && (name.indexOf(selected) === 0 || selected.indexOf(name) === 0)) {
      nameMatches.push(record);
    }
  });

  if (exactMatches.length === 1) return exactMatches[0];
  if (prefixMatches.length === 1) return prefixMatches[0];
  if (nameMatches.length === 1) return nameMatches[0];
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
        ? Math.max(34, findLastDescriptionEntryRow_(sheet, descriptionColumn + 1, info.dataStartRow))
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

  // The PPMP Form line-item table starts at row 15. Ignore edits to
  // title/metadata rows so they cannot be mistaken for line-item edits.
  const dataStartRow = isPPMPForm ? 15 : info.dataStartRow;

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

  const firstRow = Math.max(range.getRow(), dataStartRow);
  const lastRow = range.getRow() + range.getNumRows() - 1;

  // A paste/edit wholly above the line-item table must never autofill or
  // calculate against metadata cells. Summary metadata is still refreshed.
  if (firstRow > lastRow) {
    if (isPPMPForm) {
      SpreadsheetApp.flush();
      updatePPMPSummary_(spreadsheet);
    }
    return;
  }

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

      let record = sourceLookup[normalizeValue_(description)];

      // Older dropdown values may contain a slightly different spacing or
      // punctuation version. If the exact combined label is not found, only
      // use a unique match on the item name before the specifications.
      if (!record && description) {
        const baseName = description.split(',')[0].trim();
        const matches = Object.keys(sourceLookup).filter(function(key) {
          return key.split(',')[0].trim() === normalizeValue_(baseName) ||
            normalizeValue_(key.split(',')[0].trim()) === normalizeValue_(baseName);
        });
        if (matches.length === 1) record = sourceLookup[matches[0]];
      }

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
  }
}


/**
 * Installable edit handler for the PPMP Form sheet.
 * Always refreshes the summary after normal autofill/calculation handling,
 * including edits to category, quantity, cost, fiscal year, end-user, clearing
 * cells, and multi-cell pastes. No page reload is required.
 */
/**
 * Backward-compatible handler for older installable triggers.
 * Some existing spreadsheets still reference this function name.
 * Keep this alias so those triggers do not fail with
 * "Script function not found: autoSuggestEntriesOnEdit".
 */
function autoSuggestEntriesOnEdit(e) {
  return handlePPMPFormEdit(e);
}


function handlePPMPFormEdit(e) {
  if (!e || !e.range || !e.source) {
    console.error(
      'PPMP AUTOFILL DIAGNOSTIC: Missing edit event, range, or source. ' +
      'Do not run handlePPMPFormEdit manually; test by editing the sheet.'
    );
    return;
  }

  const editedRange = e.range;
  const sheet = editedRange.getSheet();
  const spreadsheet = e.source;

  // Record enough context to diagnose wrong-project, wrong-sheet, and
  // wrong-cell issues in Apps Script > Executions > Logs.
  console.log('PPMP AUTOFILL EVENT: ' + JSON.stringify({
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    sheetName: sheet.getName(),
    editedRange: editedRange.getA1Notation(),
    firstRow: editedRange.getRow(),
    lastRow: editedRange.getLastRow(),
    firstColumn: editedRange.getColumn(),
    lastColumn: editedRange.getLastColumn(),
    displayedValue: editedRange.getDisplayValue(),
    triggerUid: e.triggerUid || 'Not provided'
  }));

  if (sheet.getName() !== ENTRY_CONFIG.PPMP_FORM_SHEET) {
    console.log(
      'PPMP AUTOFILL STOP: Expected sheet "' +
      ENTRY_CONFIG.PPMP_FORM_SHEET + '", received "' + sheet.getName() + '".'
    );
    return;
  }

  // Process edits that intersect the General Description column A.
  // This supports A1 as requested as well as the PPMP line-item rows.
  // The processor only writes fields when the cell value matches Entries.
  if (editedRange.getColumn() > 1 || editedRange.getLastColumn() < 1) {
    console.log(
      'PPMP AUTOFILL STOP: Edit does not intersect General Description column A.'
    );
    return;
  }

  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lock.waitLock(30000);
    lockAcquired = true;

    console.log(
      'PPMP AUTOFILL PASS: Event reached the expected sheet and line-item range.'
    );

    // Process the item fields first. Any source access, lookup, or write error
    // is logged and shown in a toast rather than silently leaving blanks.
    processPPMPFormEditDirect_(e, spreadsheet);
    SpreadsheetApp.flush();

    console.log('PPMP AUTOFILL PASS: Direct edit processor returned.');

    try {
      updatePPMPSummary_(spreadsheet);
      SpreadsheetApp.flush();
      console.log('PPMP AUTOFILL PASS: PPMP Summary refreshed.');
    } catch (summaryError) {
      console.error(
        'PPMP SUMMARY DIAGNOSTIC ERROR: ' +
        (summaryError.stack || summaryError.message || summaryError)
      );
      spreadsheet.toast(
        'PPMP Summary refresh failed: ' + summaryError.message,
        'PPMP Summary',
        10
      );
    }
  } catch (autofillError) {
    console.error(
      'PPMP AUTOFILL DIAGNOSTIC ERROR: ' +
      (autofillError.stack || autofillError.message || autofillError)
    );
    spreadsheet.toast(
      'Autofill failed: ' + (autofillError.message || autofillError),
      'PPMP Form',
      12
    );
    throw autofillError;
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * Deterministic PPMP Form edit processor for the known template layout:
 * A Description, B Classification, C Category, D Quantity, E UOM,
 * L Unit Cost, M Estimated Budget. Line items begin at row 15.
 *
 * This deliberately avoids dynamic header detection because the template uses
 * merged and multi-row headers that can be ambiguous.
 */
function processPPMPFormEditDirect_(e, spreadsheet) {
  const editedRange = e.range;
  const sheet = editedRange.getSheet();
  if (!sheet || sheet.getName() !== ENTRY_CONFIG.PPMP_FORM_SHEET) return;

  const firstRow = editedRange.getRow();
  const lastRow = editedRange.getLastRow();
  const firstCol = editedRange.getColumn();
  const lastCol = editedRange.getLastColumn();
  if (lastRow < 1) return;

  // The line-item table starts at row 15. Ignore metadata/header edits.
  // Keep this guard in the processor as well as the event handler so direct
  // diagnostic calls cannot accidentally write into the form heading.
  if (lastRow < 15) return;

  const rowStart = Math.max(15, firstRow);
  const descriptionWasEdited = firstCol <= 1 && lastCol >= 1;
  const quantityWasEdited = firstCol <= 4 && lastCol >= 4;
  const costWasEdited = firstCol <= 12 && lastCol >= 12;

  // Ignore edits unrelated to General Description, Quantity, or Unit Cost.
  if (!descriptionWasEdited && !quantityWasEdited && !costWasEdited) return;

  const sourceSpreadsheet = SpreadsheetApp.openById(
    ENTRY_CONFIG.SOURCE_SPREADSHEET_ID
  );
  const sourceSheet = sourceSpreadsheet.getSheetByName(
    ENTRY_CONFIG.SOURCE_SHEET_NAME
  );
  if (!sourceSheet) throw new Error('The source Entries worksheet was not found.');

  const sourceData = loadSourceLookup_(sourceSheet);
  const lookup = sourceData.lookup;

  for (let row = rowStart; row <= lastRow; row++) {
    const descriptionCell = sheet.getRange(row, 1);
    const selectedDescription = String(
      descriptionCell.getDisplayValue() || ''
    ).trim();

    // Only populate source fields when the General Description cell was edited,
    // or when an existing row's A-cell can be resolved during a quantity/cost edit.
    if (descriptionWasEdited || quantityWasEdited || costWasEdited) {
      if (!selectedDescription) {
        if (descriptionWasEdited) {
          sheet.getRange(row, 2).clearContent();  // Type / Classification
          sheet.getRange(row, 3).clearContent();  // Category
          sheet.getRange(row, 5).clearContent();  // Unit or Measurement
          sheet.getRange(row, 12).clearContent(); // Unit Cost
          sheet.getRange(row, 13).clearContent(); // Estimated Budget
        }
        continue;
      }

      const record = findMatchingEntryRecord_(selectedDescription, lookup);
      if (record) {
        // Fixed PPMP Form layout: A Description, B Type, C Category,
        // D Quantity, E Unit/Measurement, L Unit Cost, M Estimated Budget.
        sheet.getRange(row, 2).setValue(record.classification || '');
        sheet.getRange(row, 3).setValue(record.category || '');
        sheet.getRange(row, 5).setValue(record.uom || '');

        const unitCost = parseNumericValue_(record.unitCost);
        if (unitCost === null) {
          sheet.getRange(row, 12).clearContent();
        } else {
          sheet.getRange(row, 12).setValue(unitCost);
        }
      } else if (descriptionWasEdited) {
        throw new Error(
          'Could not match PPMP Form!A' + row + ' to a record in Entries. ' +
          'Value entered: "' + selectedDescription + '". ' +
          'Please select the full General Description from the dropdown, ' +
          'or make sure the item name and technical specifications match Entries.'
        );
      }
    }

    // Always recalculate Estimated Budget when the description, quantity,
    // or unit cost is edited, using the current values in D and L.
    if (descriptionWasEdited || quantityWasEdited || costWasEdited) {
      const unitCost = parseNumericValue_(sheet.getRange(row, 12).getValue());
      const quantity = parseNumericValue_(sheet.getRange(row, 4).getValue());
      const budgetCell = sheet.getRange(row, 13);

      if (unitCost === null || quantity === null) {
        budgetCell.clearContent();
      } else {
        budgetCell.setValue(unitCost * quantity);
      }
    }
  }

  SpreadsheetApp.flush();
  updatePPMPSummary_(spreadsheet);
  SpreadsheetApp.flush();
}

/**
 * Installable on-change handler: when rows are inserted into the destination
 * spreadsheet, rebuild the PPMP Form borders across A:O. This complements
 * addPPMPEntryRow(), which also applies borders after preparing its new row.
 */
function handlePPMPRowChange(e) {
  if (!e || !e.source || e.changeType !== 'INSERT_ROW') return;

  const spreadsheet = e.source;
  const sheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (!sheet) return;

  // Google Sheets' onChange event does not provide the inserted row number.
  // Therefore, border the full initial table plus a buffer beyond the last
  // used row. This catches blank rows inserted directly below the current
  // data, including when multiple rows are inserted at once.
  const info = findDestinationHeaders_(sheet);
  const descriptionIndex = info
    ? info.headerMap[normalizeHeader_(ENTRY_CONFIG.DESTINATION_HEADERS.description)]
    : undefined;
  const descriptionColumn = Number.isInteger(descriptionIndex)
    ? descriptionIndex + 1
    : 1;
  const lastEntryRow = findLastDescriptionEntryRow_(sheet, descriptionColumn, 15);
  const lastUsedRow = Math.max(sheet.getLastRow(), lastEntryRow, 34);
  const borderThroughRow = Math.min(
    sheet.getMaxRows(),
    Math.max(34, lastUsedRow + 25)
  );

  applyPPMPEntryTableBorders_(sheet, borderThroughRow);
  SpreadsheetApp.flush();
}

/**
 * Manually repair the complete PPMP Form table border grid (A:O).
 */
function repairPPMPEntryTableBorders() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Open the destination spreadsheet first.');
  const sheet = spreadsheet.getSheetByName(ENTRY_CONFIG.PPMP_FORM_SHEET);
  if (!sheet) throw new Error('The PPMP Form sheet was not found.');
  const lastRow = Math.max(34, sheet.getLastRow());
  applyPPMPEntryTableBorders_(sheet, lastRow);
  spreadsheet.toast('PPMP Form table borders refreshed for rows 15–' + lastRow + '.', 'PPMP Form', 5);
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

  // Commit cleared output and synchronized Fiscal Year/End-User immediately,
  // including when the PPMP Form currently has no completed entries.
  if (results.length === 0) {
    SpreadsheetApp.flush();
    return;
  }

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

  // Make summary writes visible immediately without requiring a browser reload.
  SpreadsheetApp.flush();
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

  // The PPMP Form has title/metadata above the line-item table.
  // All description dropdowns and line-item edits must start at row 15.
  const dataStartRow =
    sheet.getName() === ENTRY_CONFIG.PPMP_FORM_SHEET
      ? 15
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