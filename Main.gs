/**
 * MAIN SCRIPT
 * Central Hub for Triggers, Menus, and Sync Logic.
 * UPDATED: Phase 5 (Release Metadata Management)
 */

// =========================================
// 1. GLOBAL TRIGGER ROUTER (Installable Trigger)
// =========================================
/**
 * IMPORTANT: This function must be set up as an INSTALLABLE TRIGGER.
 * Go to Triggers -> Add Trigger -> processGlobalEdits -> From Spreadsheet -> On Edit.
 * DO NOT name this "onEdit", or simple trigger limitations will block the password popup.
 */
function processGlobalEdits(e) {
  if (!e) return;

  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  if (sheetName !== "ORDERING LIST") return;

  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  var newVal = e.value;
  var oldVal = e.oldValue;

  // ROUTE 1: Config Section Edits (Column 4 - Part ID)
  if (col === 4) {
    // --- LOCATE SECTIONS ---
    var configFinder = sheet.getRange("A:A").createTextFinder("CONFIG").matchEntireCell(true).findNext();
    var moduleFinder = sheet.getRange("A:A").createTextFinder("MODULE").matchEntireCell(true).findNext();

    // Define Boundaries for Config Section
    var configStart = configFinder ? configFinder.getRow() + 1 : 0;
    var configEnd = moduleFinder ? moduleFinder.getRow() - 1 : 0;

    // Execute Config Section Logic (Shopping Lists)
    if (row >= configStart && row <= configEnd && configStart > 0) {
      handleConfigSection(sheet, row, newVal, oldVal);
    }
  }

  // ROUTE 2: Release Checkbox Edits (Column 7 - RELEASE)
  if (col === 7) {
    handleCheckboxEdit(e); // Located in OrderingListHandlers.gs
  }
}

// =========================================
// 2. MENU CREATION & INITIALIZATION
// =========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // 1. Auto-Fix Locale (Phase 5 Requirement)
  // Ensures Date objects are written as dd/MM/yyyy compatible
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale('en_MY');

  // Menu 1: Refresh (Database Sync)
  ui.createMenu('Refresh')
    .addItem('Sync REF_DATA (DB Only)', 'runMasterSync')
    // New Tool for Option B (Retroactive Fix)
    //.addItem('Initialize Release Columns (One-Time)', 'initializeReleaseColumns') 
    .addToUi();

  // Menu 2: Configurator (The UI)
  ui.createMenu('Configurator')
    .addItem('Open Configurator Window', 'openSidebar')
    .addToUi();

  // Menu 3: Sync to Order List (Production Sync)
  ui.createMenu('Sync to Order List')
    .addItem('Run Synchronization', 'runProductionSync')
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ModuleConfigurator')
    .setWidth(450)
    .setHeight(650);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Module Configurator');
}

// =========================================
// 3. MASTER SYNC LOGIC (REF_DATA ONLY)
// =========================================
function runMasterSync() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sourceSpreadsheetId = "1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA";
    var sourceTabName = "BOM Structure Tree Diagram";

    var sourceSS;
    try {
      sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
    } catch (e) {
      throw new Error("Could not open Source Spreadsheet. Check ID. " + e.message);
    }

    var sourceSheet = sourceSS.getSheetByName(sourceTabName);
    if (!sourceSheet) throw new Error("Source tab '" + sourceTabName + "' not found.");

    // We only touch REF_DATA
    var destSS = SpreadsheetApp.getActiveSpreadsheet();
    var refSheet = destSS.getSheetByName("REF_DATA");

    // A. Update Reference Data (Cols C:AD)
    updateReferenceData(sourceSS, sourceSheet);

    // B. Update Shopping Lists (Cols I:L)
    updateShoppingLists(sourceSheet);

    // C. PROCESS TOOLING OPTIONS (Cols P:V)
    processToolingOptions(sourceSS, refSheet);

    // D. PROCESS VISION DATA REMOVED (Cols AF:AK No longer used)

    ui.alert("Sync Complete", "REF_DATA has been updated.", ui.ButtonSet.OK);
  } catch (e) {
    console.error(e);
    ui.alert("Error during Sync", e.message, ui.ButtonSet.OK);
  }
}

// =========================================
// 4. CONFIG SECTION HANDLERS
// =========================================
function handleConfigSection(sheet, row, newVal, oldVal) {
  var BASIC_TOOL_TRIGGER = "430001-A378";
  var PNEUMATIC_TRIGGER = "430001-A714";

  // Delete Logic
  if (oldVal === BASIC_TOOL_TRIGGER) {
    if (row + 10 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 10);
  }
  if (oldVal === PNEUMATIC_TRIGGER) {
    if (row + 3 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 3);
  }

  // Insert Logic
  if (newVal === BASIC_TOOL_TRIGGER) {
    insertShoppingList(sheet, row, 10, "REF_DATA!I:I", "REF_DATA!I:J");
  }
  if (newVal === PNEUMATIC_TRIGGER) {
    insertShoppingList(sheet, row, 3, "REF_DATA!K:K", "REF_DATA!K:L");
  }
}

function insertShoppingList(sheet, row, count, dropdownRef, vlookupRef) {
  sheet.insertRowsAfter(row, count);
  var startInsertRow = row + 1;

  // Validation Rule
  var dropDownRange = sheet.getRange(startInsertRow, 4, count, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRef), true)
    .setAllowInvalid(true).build();
  dropDownRange.setDataValidation(rule);

  // Description Formulas
  var descRange = sheet.getRange(startInsertRow, 5, count, 1);
  var formulas = [];
  for (var i = 0; i < count; i++) {
    formulas.push(['=IFERROR(VLOOKUP(D' + (startInsertRow + i) + ', ' + vlookupRef + ', 2, FALSE), "")']);
  }
  descRange.setFormulas(formulas);

  // Checkboxes & Formatting
  sheet.getRange(startInsertRow, 7, count, 1).insertCheckboxes();
  sheet.getRange(startInsertRow, 9, count, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
  sheet.getRange(startInsertRow, 6, count, 1).clearContent();
}

// =========================================
// 5. PRODUCTION SYNC (FINAL PHASE 4.3)
// =========================================
function runProductionSync() {
  var ui = SpreadsheetApp.getUi();

  // Confirmation Dialog
  var result = ui.alert(
    'Confirm Synchronization',
    'This will append all new (unsynced) configurations from the Trial Layout to the Ordering List.\n\nAre you sure you want to proceed?',
    ui.ButtonSet.YES_NO);

  if (result == ui.Button.NO) {
    return;
  }

  try {
    // 1. Extract Data
    var payload = extractProductionData();

    // 2. Validation (Gatekeeper Fix Phase 7c)
    // Runs if ANY relevant payload section has data
    var hasData = (
      payload.rowsToMarkSynced.length > 0 ||
      payload.PC.length > 0 ||
      payload.CONFIG.length > 0 ||
      payload.TOOLING.length > 0 ||
      payload.CORE.length > 0
    );

    if (!hasData) {
      ui.alert("No New Data", "No new configurations or machine setup data found to sync.", ui.ButtonSet.OK);
      return;
    }

    // 3. Inject Data (Write to Production)
    var itemsAdded = injectProductionData(payload);

    // 4. Update Status (Write to Staging)
    markRowsAsSynced(payload.rowsToMarkSynced);

    // 5. Success Message
    var msg = "Synchronization Successful!\n\n";
    msg += "Rows Processed: " + payload.rowsToMarkSynced.length + "\n";
    msg += "Items Added to Order List: " + itemsAdded;

    ui.alert("Success", msg, ui.ButtonSet.OK);

  } catch (e) {
    console.error(e);
    ui.alert("Sync Error", e.message, ui.ButtonSet.OK);
  }
}

// =========================================
// 6. SYNC UTILITIES
// =========================================

function updateReferenceData(sourceSS, sourceSheet) {
  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var refSheetName = "REF_DATA";
  var refSheet = destSS.getSheetByName(refSheetName);

  if (!refSheet) {
    refSheet = destSS.insertSheet(refSheetName);
    refSheet.hideSheet();
  }

  // 1. Backup Existing Mappings (Preserve W:AF)
  var lastRefRow = refSheet.getLastRow();
  var existingMappings = {};
  if (lastRefRow > 0) {
    var currentData = refSheet.getRange(1, 3, lastRefRow, 30).getValues(); // C:AF
    for (var i = 0; i < currentData.length; i++) {
      var pId = currentData[i][0].toString().trim();
      if (pId !== "") {
        existingMappings[pId] = {
          eId: currentData[i][20], eDesc: currentData[i][21],
          tId: currentData[i][22], tDesc: currentData[i][23],
          jId: currentData[i][24], jDesc: currentData[i][25],
          vId: currentData[i][26], vDesc: currentData[i][27],
          spId: currentData[i][28], spDesc: currentData[i][29]
        };
      }
    }
  }

  // 2. Clear Target Areas
  refSheet.getRange("A:D").clear();
  refSheet.getRange("W:AF").clear();

  // 3. Fetch New Data
  var configItems = fetchRawItems(sourceSheet, "OPTIONAL MODULE: 430001-A712", 6, 7, ["CONFIGURABLE MODULE"]);
  var moduleItems = fetchRawItems(sourceSheet, "CONFIGURABLE MODULE: 430001-A713", 6, 7, ["CONFIGURABLE VISION MODULE"]);

  // 4. Write Data
  if (configItems.length > 0) refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems);
  if (moduleItems.length > 0) {
    var moduleOutput = [];
    var mappingOutput = [];

    for (var m = 0; m < moduleItems.length; m++) {
      var mId = moduleItems[m][0].toString().trim();
      var mDesc = moduleItems[m][1];

      var eId = "", eDesc = "", tId = "", tDesc = "", jId = "", jDesc = "", vId = "", vDesc = "";
      var spId = "", spDesc = "";

      if (existingMappings[mId]) {
        eId = existingMappings[mId].eId; eDesc = existingMappings[mId].eDesc;
        tId = existingMappings[mId].tId; tDesc = existingMappings[mId].tDesc;
        jId = existingMappings[mId].jId; jDesc = existingMappings[mId].jDesc;
        vId = existingMappings[mId].vId; vDesc = existingMappings[mId].vDesc;
        spId = existingMappings[mId].spId; spDesc = existingMappings[mId].spDesc;
      }
      moduleOutput.push([mId, mDesc]);
      mappingOutput.push([eId, eDesc, tId, tDesc, jId, jDesc, vId, vDesc, spId, spDesc]);
    }
    refSheet.getRange(1, 3, moduleOutput.length, 2).setValues(moduleOutput);
    refSheet.getRange(1, 23, mappingOutput.length, 10).setValues(mappingOutput);
  }
}

function updateShoppingLists(sourceSheet) {
  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = destSS.getSheetByName("REF_DATA");
  refSheet.getRange("I:L").clear();

  var basicItems = fetchShoppingListItems(sourceSheet, "List-Optional Basic Tool Module: 430001-A378", 12, 13, "STRICT");
  if (basicItems.length > 0) refSheet.getRange(1, 9, basicItems.length, 2).setValues(basicItems);

  var pneumaticItems = fetchShoppingListItems(sourceSheet, "List-Optional Pneumatic Module : 430001-A714", 12, 13, "SKIP_EMPTY");
  if (pneumaticItems.length > 0) refSheet.getRange(1, 11, pneumaticItems.length, 2).setValues(pneumaticItems);
}

function processToolingOptions(sourceSS, refSheet) {
  var toolingSheet = sourceSS.getSheetByName("Tooling Illustration");
  if (!toolingSheet) return;

  // --- STEP 1: FETCH PARENT DESCRIPTIONS FROM BOM TREE (THE JOIN) ---
  // Source: BOM Structure Tree Diagram (Cols O & P)
  var bomSheet = sourceSS.getSheetByName("BOM Structure Tree Diagram");
  var parentDescMap = {};

  if (bomSheet) {
    var lastBomRow = bomSheet.getLastRow();
    // Read O (15) and P (16)
    var bomData = bomSheet.getRange(1, 15, lastBomRow, 2).getValues();
    for (var b = 0; b < bomData.length; b++) {
      var pId = String(bomData[b][0]).trim();
      var pDesc = String(bomData[b][1]).trim();
      if (pId && pId !== "Part ID") {
        parentDescMap[pId] = pDesc;
      }
    }
  }

  // --- STEP 2: PROCESS TOOLING ILLUSTRATION (THE STRUCTURE) ---
  var lastRow = toolingSheet.getLastRow();
  var rawData = toolingSheet.getRange(1, 1, lastRow, 8).getValues();
  var databaseOutput = [];
  var currentParentID = null;
  var currentCategory = null;

  for (var i = 0; i < rawData.length; i++) {
    var colA = String(rawData[i][0]).trim();
    var colB = String(rawData[i][1]).trim();
    var colF = String(rawData[i][5]).trim(); // Child ID
    var colH = String(rawData[i][7]).trim(); // Description
    var match = colA.match(/\[(.*?)\]/);

    // 1. Parent Detection (Start Block)
    if (match && match[1]) {
      currentParentID = match[1];
      currentCategory = null;
    }
    // 2. Stop Signal: Non-Bracketed Text in Col A (Visual Break)
    else if (colA !== "" && !match) {
      currentParentID = null;
    }

    if (colB !== "") { currentCategory = colB; }

    // 4. Processing (Only if valid parent exists)
    if (currentParentID && colF !== "" && colF !== "Part ID") {
      databaseOutput.push([currentParentID, colF, (currentCategory || ""), colH]);
    }
  }

  refSheet.getRange("P:S").clearContent();
  if (databaseOutput.length > 0) refSheet.getRange(1, 16, databaseOutput.length, 4).setValues(databaseOutput);

  // --- STEP 3: GENERATE MENU WITH MERGED DESCRIPTIONS ---
  var menuOutput = [];
  if (databaseOutput.length > 0) {
    var grouped = {};
    var orderParents = [];
    for (var k = 0; k < databaseOutput.length; k++) {
      var pID = databaseOutput[k][0];
      if (!grouped[pID]) { grouped[pID] = []; orderParents.push(pID); }
      grouped[pID].push({
        partId: databaseOutput[k][1],
        cat: databaseOutput[k][2],
        desc: databaseOutput[k][3]
      });
    }

    for (var p = 0; p < orderParents.length; p++) {
      var parent = orderParents[p];
      var items = grouped[parent];

      // LOOKUP DESCRIPTION
      var parentDescription = parentDescMap[parent] || "";
      var combinedParentLabel = parent + (parentDescription ? " | " + parentDescription : "");

      var lastCat = null;
      for (var m = 0; m < items.length; m++) {
        var itm = items[m];
        var thisCat = itm.cat;
        if (thisCat !== "" && thisCat !== lastCat) {
          menuOutput.push([combinedParentLabel, "--- " + thisCat + " ---"]);
          lastCat = thisCat;
        }

        var packedValue = itm.partId + " :: " + (itm.desc || "");
        menuOutput.push([combinedParentLabel, packedValue]);
      }
    }
  }
  refSheet.getRange("U:V").clearContent();
  if (menuOutput.length > 0) refSheet.getRange(1, 21, menuOutput.length, 2).setValues(menuOutput);
}

function fetchShoppingListItems(sourceSheet, triggerPhrase, colID, colDesc, stopMode) {
  var lastRow = sourceSheet.getLastRow();
  var idColumnVals = sourceSheet.getRange(1, colID, lastRow, 1).getValues();
  var startRowIndex = -1;

  for (var i = 0; i < idColumnVals.length; i++) {
    if (idColumnVals[i][0].toString().trim().indexOf(triggerPhrase) > -1) { startRowIndex = i + 1; break; }
  }

  if (startRowIndex === -1) return [];
  var rowsRemaining = lastRow - startRowIndex;
  if (rowsRemaining < 1) return [];

  var idData = sourceSheet.getRange(startRowIndex + 1, colID, rowsRemaining, 1).getValues();
  var descData = sourceSheet.getRange(startRowIndex + 1, colDesc, rowsRemaining, 1).getValues();
  var items = [];

  for (var k = 0; k < idData.length; k++) {
    var pID = idData[k][0].toString().trim();
    var desc = descData[k][0].toString().trim();
    if (pID.toUpperCase().indexOf("LIST-") === 0) break;
    if (stopMode === "STRICT") { if (pID === "" || pID === "---") break; }
    if (stopMode === "SKIP_EMPTY") { if (pID === "" || pID === "---") continue; if (!/^\d/.test(pID)) break; }
    items.push([pID, desc]);
  }
  return items;
}

function fetchRawItems(sourceSheet, triggerPhrase, colID, colDesc, stopPhrases) {
  var lastRow = sourceSheet.getLastRow();
  var rangeValues = sourceSheet.getRange(1, colID, lastRow, 1).getValues();
  var startRowIndex = -1;

  for (var i = 0; i < rangeValues.length; i++) {
    if (rangeValues[i][0].toString().trim().indexOf(triggerPhrase) > -1) { startRowIndex = i + 1; break; }
  }

  if (startRowIndex === -1) return [];
  var rowsToGrab = lastRow - startRowIndex;
  var idData = sourceSheet.getRange(startRowIndex + 1, colID, rowsToGrab, 1).getValues();
  var descData = sourceSheet.getRange(startRowIndex + 1, colDesc, rowsToGrab, 1).getValues();
  var items = [];

  for (var k = 0; k < idData.length; k++) {
    var pID = idData[k][0].toString().trim();
    var desc = descData[k][0].toString().trim();
    if (stopPhrases && stopPhrases.some(function (s) { return pID.indexOf(s) > -1; })) break;
    if (pID.indexOf(":") > -1) break;
    if (pID !== "" && pID !== "---") items.push([pID, desc]);
  }
  return items;
}
