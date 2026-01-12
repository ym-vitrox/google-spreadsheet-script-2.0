/**
 * MAIN SCRIPT
 * Central Hub for Triggers, Menus, and Sync Logic.
 */

// =========================================
// 1. LIVE TRIGGER (Config Section Only)
// =========================================
function onEdit(e) {
  if (!e) return;
  
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "ORDERING LIST") return;

  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  // We strictly look for edits in Column D (Part ID)
  if (col !== 4) return;
  
  var newVal = e.value; 
  var oldVal = e.oldValue; 
  
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

// =========================================
// 2. MENU CREATION
// =========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Refresh') 
    .addItem('Sync REF_DATA (DB Only)', 'runMasterSync') 
    .addToUi();

  ui.createMenu('Configurator')
    .addItem('Open Configurator Window', 'openSidebar')
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
// 5. SYNC UTILITIES
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
  // CHANGE: Now backing up 30 columns (C:AF) to include Spare Parts (AE/AF)
  var lastRefRow = refSheet.getLastRow();
  var existingMappings = {}; 
  if (lastRefRow > 0) {
    var currentData = refSheet.getRange(1, 3, lastRefRow, 30).getValues(); // Was 28 (C:AD), now 30 (C:AF)
    for (var i = 0; i < currentData.length; i++) {
      var pId = currentData[i][0].toString().trim();
      if (pId !== "") {
        existingMappings[pId] = {
          eId: currentData[i][20], eDesc: currentData[i][21],
          tId: currentData[i][22], tDesc: currentData[i][23],
          jId: currentData[i][24], jDesc: currentData[i][25],
          vId: currentData[i][26], vDesc: currentData[i][27],
          // New Spare Part Columns
          spId: currentData[i][28], spDesc: currentData[i][29]
        };
      }
    }
  }
  
  // 2. Clear Target Areas
  refSheet.getRange("A:D").clear(); 
  refSheet.getRange("W:AF").clear(); // CHANGE: Clear W:AF instead of W:AD
  
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
      var spId = "", spDesc = ""; // New vars
      
      if (existingMappings[mId]) {
        eId = existingMappings[mId].eId; eDesc = existingMappings[mId].eDesc;
        tId = existingMappings[mId].tId; tDesc = existingMappings[mId].tDesc;
        jId = existingMappings[mId].jId; jDesc = existingMappings[mId].jDesc;
        vId = existingMappings[mId].vId; vDesc = existingMappings[mId].vDesc;
        spId = existingMappings[mId].spId; spDesc = existingMappings[mId].spDesc; // Restore new data
      }
      moduleOutput.push([mId, mDesc]);
      // Push 10 columns now
      mappingOutput.push([eId, eDesc, tId, tDesc, jId, jDesc, vId, vDesc, spId, spDesc]);
    }
    refSheet.getRange(1, 3, moduleOutput.length, 2).setValues(moduleOutput);
    refSheet.getRange(1, 23, mappingOutput.length, 10).setValues(mappingOutput); // CHANGE: Width is 10
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
    
    if (match && match[1]) { currentParentID = match[1]; currentCategory = null; }
    if (colB !== "") { currentCategory = colB; }
    
    if (currentParentID && colF !== "" && colF !== "Part ID") {
      databaseOutput.push([currentParentID, colF, (currentCategory || ""), colH]);
    }
  }
  
  refSheet.getRange("P:S").clearContent();
  if (databaseOutput.length > 0) refSheet.getRange(1, 16, databaseOutput.length, 4).setValues(databaseOutput);
  
  // Generate Shadow Menu
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
      var lastCat = null;
      for (var m = 0; m < items.length; m++) {
        var itm = items[m];
        var thisCat = itm.cat;
        if (thisCat !== "" && thisCat !== lastCat) { 
          menuOutput.push([parent, "--- " + thisCat + " ---"]); 
          lastCat = thisCat; 
        }
        
        var packedValue = itm.partId + " :: " + (itm.desc || "");
        menuOutput.push([parent, packedValue]);
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
    if (stopPhrases && stopPhrases.some(function(s) { return pID.indexOf(s) > -1; })) break;
    if (pID.indexOf(":") > -1) break;
    if (pID !== "" && pID !== "---") items.push([pID, desc]); 
  }
  return items;
}
