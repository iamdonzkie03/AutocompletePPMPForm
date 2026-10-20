
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
  HEADER_SEARCH_ROWS: 10,

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

  helperSheet.clearContents();

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

    const availableRows = sheet.getMaxRows() - info.dataStartRow + 1;
    if (availableRows <= 0) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(helperSheet.getRange(1, 1, sourceData.options.length, 1), true)
      .setAllowInvalid(false)
      .setHelpText('Select a project description from the Entries source.')
      .build();

    sheet.getRange(
      info.dataStartRow, descriptionColumn + 1, availableRows, 1
    ).setDataValidation(rule);

    updated++;
  });

  return updated;
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

    const availableRows =
      sheet.getMaxRows() - info.dataStartRow + 1;

    if (availableRows <= 0) return;

    sheet.getRange(
      info.dataStartRow,
      column + 1,
      availableRows,
      1
    ).setDataValidation(rule);

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
    info.dataStartRow
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

  const formSheet =
    spreadsheet.getSheetByName(
      ENTRY_CONFIG.PPMP_FORM_SHEET
    );

  const summarySheet =
    spreadsheet.getSheetByName(
      ENTRY_CONFIG.PPMP_SUMMARY_SHEET
    );

  if (!formSheet) {
    throw new Error(
      'The PPMP Form sheet was not found.'
    );
  }

  if (!summarySheet) {
    throw new Error(
      'The PPMP Summary sheet was not found.'
    );
  }

  const formInfo =
    findCategoryAmountHeaders_(formSheet);

  if (!formInfo) {
    throw new Error(
      'Could not locate Category and an amount column in PPMP Form.'
    );
  }

  const summaryInfo =
    findCategoryAmountHeaders_(summarySheet);

  if (!summaryInfo) {
    throw new Error(
      'PPMP Summary must contain Category and an amount header ' +
      'on the same row within its first 10 rows.'
    );
  }

  const totals = new Map();
  // PPMP Form entries begin at row 5; do not summarize header rows.
  const firstDataRow = ENTRY_CONFIG.PPMP_FORM_SHEET === formSheet.getName()
    ? 5
    : formInfo.headerRow + 1;
  const lastDataRow = formSheet.getLastRow();
  const lastColumn = formSheet.getLastColumn();

  if (lastDataRow >= firstDataRow) {
    const data = formSheet.getRange(
      firstDataRow,
      1,
      lastDataRow - firstDataRow + 1,
      lastColumn
    ).getValues();

    data.forEach(function(row) {
      const category = String(
        row[formInfo.categoryColumn - 1] || ''
      ).trim();

      if (!category) return;

      const amount = parseNumericValue_(
        row[formInfo.amountColumn - 1]
      );

      const key = normalizeValue_(category);

      if (!totals.has(key)) {
        totals.set(key, {
          category: category,
          total: 0
        });
      }

      if (amount !== null) {
        totals.get(key).total += amount;
      }
    });
  }

  const results =
    Array.from(totals.values()).sort(
      function(a, b) {
        return a.category.localeCompare(b.category);
      }
    );

  const outputStartRow =
    summaryInfo.headerRow + 1;

  const availableRows =
    summarySheet.getMaxRows() -
    summaryInfo.headerRow;

  if (availableRows > 0) {
    summarySheet.getRange(
      outputStartRow,
      summaryInfo.categoryColumn,
      availableRows,
      1
    ).clearContent();

    summarySheet.getRange(
      outputStartRow,
      summaryInfo.amountColumn,
      availableRows,
      1
    ).clearContent();
  }

  if (results.length === 0) return;

  const requiredLastRow =
    outputStartRow + results.length - 1;

  if (
    requiredLastRow >
    summarySheet.getMaxRows()
  ) {
    summarySheet.insertRowsAfter(
      summarySheet.getMaxRows(),
      requiredLastRow - summarySheet.getMaxRows()
    );
  }

  summarySheet.getRange(
    outputStartRow,
    summaryInfo.categoryColumn,
    results.length,
    1
  ).setValues(
    results.map(function(item) {
      return [item.category];
    })
  );

  summarySheet.getRange(
    outputStartRow,
    summaryInfo.amountColumn,
    results.length,
    1
  ).setValues(
    results.map(function(item) {
      return [item.total];
    })
  );

  summarySheet.getRange(
    outputStartRow,
    summaryInfo.amountColumn,
    results.length,
    1
  ).setNumberFormat('#,##0.00');
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
  const preferredAmountHeaders = [
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