/**
 * ORDERING LIST SCRIPT
 * * Features:
 * 1. Master Sync from Source BOM (Preserves User Mappings in REF_DATA Cols W-AD).
 * 2. Module Section: Spreadsheet-driven Dependency Logic (REF_DATA Cols C & W,X,Y,Z, AA,AB, AC,AD).
 * - Electrical (Col W/X): Rotational (Based on instance count).
 * - Tooling (Col Y/Z): Stacked (Fixed, multiple items allowed).
 * - Tooling Options (REF_DATA P->S): Single Row Dropdown with Dynamic Formulas.
 * - Jigs (Col AA/AB): Stacked (Fixed, manual mapping, bottom of list).
 * - Vision (Col AC/AD): Triggered by manual mapping. Single (Fixed) or Multiple (Dropdown). New Bottom Layer.
 * 3. Shopping List Logic (Basic Tool & Pneumatic) via onEdit (Config Section).
 * 4. Vision Section: Categorized Dropdowns (REF_DATA Cols AF-AH).
 * 5. Renumbering Tool.
 */

// =========================================
// 1. LIVE TRIGGER (Handle Kit/Tool Insertion)
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
  var visionFinder = sheet.getRange("A:A").createTextFinder("VISION").matchEntireCell(true).findNext();
  
  // Define Boundaries
  var configStart = configFinder ? configFinder.getRow() + 1 : 0;
  var configEnd = moduleFinder ? moduleFinder.getRow() - 1 : 0;
  
  var moduleStart = moduleFinder ? moduleFinder.getRow() + 1 : 0;
  var moduleEnd = visionFinder ? visionFinder.getRow() - 1 : 0;
  
  // =========================================================
  // LOGIC A: CONFIG SECTION (Shopping Lists)
  // =========================================================
  if (row >= configStart && row <= configEnd && configStart > 0) {
    handleConfigSection(sheet, row, newVal, oldVal);
  }

  // =========================================================
  // LOGIC B: MODULE SECTION (Spreadsheet Dependencies)
  // =========================================================
  // NOTE: This logic is being migrated to UI. Kept here for legacy support if needed.
  if (row >= moduleStart && row <= moduleEnd && moduleStart > 0) {
    handleModuleSection(sheet, row, newVal, oldVal, moduleStart, moduleEnd);
  }
}

// =========================================
// LOGIC HANDLERS
// =========================================
function handleConfigSection(sheet, row, newVal, oldVal) {
  var BASIC_TOOL_TRIGGER = "430001-A378";
  var PNEUMATIC_TRIGGER = "430001-A714";
  // --- 1. DELETE LOGIC ---
  if (oldVal === BASIC_TOOL_TRIGGER) {
    if (row + 10 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 10);
  }
  if (oldVal === PNEUMATIC_TRIGGER) {
    if (row + 3 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 3);
  }

  // --- 2. INSERT LOGIC ---
  if (newVal === BASIC_TOOL_TRIGGER) {
    insertShoppingList(sheet, row, 10, "REF_DATA!I:I", "REF_DATA!I:J");
  }
  if (newVal === PNEUMATIC_TRIGGER) {
    insertShoppingList(sheet, row, 3, "REF_DATA!K:K", "REF_DATA!K:L");
  }
}

function handleModuleSection(sheet, row, newVal, oldVal, startRow, endRow) {
  var refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  // --- CONSTANTS: RUBBER TIP DEPENDENCY ---
  var RUBBER_TIP_PARENTS = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
  var RUBBER_TIP_SOURCE_ID = "430001-A380";
  
  var refData = refSheet.getRange("C:AD").getValues();
  var optionData = refSheet.getRange("P:Q").getValues();
  var menuData = refSheet.getRange("U:V").getValues();
  
  // --- CLEANUP (Handle Removal or Swap) ---
  if (oldVal) {
    var oldParentConfig = findParentConfig(refData, oldVal);
    if (oldParentConfig) {
      var possibleChildren = [];
      var oldToolIds = [];
      if (oldParentConfig.elecIds) possibleChildren = possibleChildren.concat(oldParentConfig.elecIds.split(';').map(function(s){ return s.trim(); }));
      if (oldParentConfig.toolIds) {
        var tIds = oldParentConfig.toolIds.split(';').map(function(s){ return s.trim(); });
        possibleChildren = possibleChildren.concat(tIds);
        oldToolIds = tIds;
      }
      if (oldParentConfig.jigIds) {
         possibleChildren = possibleChildren.concat(oldParentConfig.jigIds.split(';').map(function(s){ return s.trim(); }));
      }
      if (oldParentConfig.visionIds) {
         possibleChildren = possibleChildren.concat(oldParentConfig.visionIds.split(';').map(function(s){ return s.trim(); }));
      }
      for (var t = 0; t < oldToolIds.length; t++) {
        var currentToolId = oldToolIds[t];
        var grandChildren = getToolingOptionIDs(optionData, currentToolId);
        if (grandChildren.length > 0) {
          possibleChildren = possibleChildren.concat(grandChildren);
        }
        if (RUBBER_TIP_PARENTS.includes(currentToolId)) {
             var rtChildren = getToolingOptionIDs(optionData, RUBBER_TIP_SOURCE_ID);
             if (rtChildren.length > 0) {
                 possibleChildren = possibleChildren.concat(rtChildren);
             }
        }
      }
      var checkRow = row + 1;
      while (checkRow <= sheet.getMaxRows()) {
        var childPartID = sheet.getRange(checkRow, 4).getValue();
        if (possibleChildren.includes(childPartID) || (childPartID === "" && sheet.getRange(checkRow, 3).getValue() === "")) {
          sheet.deleteRow(checkRow);
        } else {
          break; 
        }
      }
    }
  }

  // --- INSERTION (Handle New Selection) ---
  if (newVal) {
    var config = findParentConfig(refData, newVal);
    if (!config) return; 

    // 1. Determine Electrical Kit (Rotation)
    var elecToInsert = null;
    if (config.elecIds) {
      var eIds = config.elecIds.split(';').map(function(s){ return s.trim(); });
      var eDescs = config.elecDesc.split(';').map(function(s){ return s.trim(); });
      
      var instanceCount = 0;
      var sectionIds = sheet.getRange(startRow, 4, endRow - startRow + 1, 1).getValues();
      for (var i = 0; i < sectionIds.length; i++) {
        if (sectionIds[i][0] == newVal) instanceCount++;
        if (startRow + i == row) break; 
      }
      
      var index = (instanceCount - 1) % eIds.length;
      if (eIds[index]) {
        elecToInsert = { id: eIds[index], desc: (eDescs[index] || ""), type: 'child' };
      }
    }

    var itemsToAdd = [];
    if (elecToInsert) itemsToAdd.push(elecToInsert);

    // 2. Determine Tooling Kit (Stacking) AND Grandchildren
    if (config.toolIds) {
      var tIds = config.toolIds.split(';').map(function(s){ return s.trim(); });
      var tDescs = config.toolDesc.split(';').map(function(s){ return s.trim(); });
      
      for (var t = 0; t < tIds.length; t++) {
        if (tIds[t]) {
          itemsToAdd.push({ id: tIds[t], desc: (tDescs[t] || ""), type: 'child' });
          var optionRange = getToolingOptionRange(menuData, tIds[t]);
          if (optionRange) {
             itemsToAdd.push({
               type: 'grandchild',
               parentToolId: tIds[t],
               refDataStart: optionRange.startRow, 
               refDataEnd: optionRange.endRow    
            });
          }
          if (RUBBER_TIP_PARENTS.includes(tIds[t])) {
              var rtRange = getToolingOptionRange(menuData, RUBBER_TIP_SOURCE_ID);
              if (rtRange) {
                  itemsToAdd.push({
                      type: 'rubber_tip',
                      parentToolId: RUBBER_TIP_SOURCE_ID,
                      refDataStart: rtRange.startRow,
                      refDataEnd: rtRange.endRow
                  });
              }
          }
        }
      }
    }
    
    // 3. Determine Jig Items
    if (config.jigIds) {
      var jIds = config.jigIds.split(';').map(function(s){ return s.trim(); });
      var jDescs = config.jigDesc.split(';').map(function(s){ return s.trim(); });
      for (var j = 0; j < jIds.length; j++) {
        if (jIds[j]) itemsToAdd.push({ id: jIds[j], desc: (jDescs[j] || ""), type: 'jig' });
      }
    }

    // 4. Determine Vision Items
    if (config.visionIds) {
      var vIds = config.visionIds.split(';').map(function(s){ return s.trim(); });
      vIds = vIds.filter(function(id) { return id.length > 0; });
      if (vIds.length === 1) itemsToAdd.push({ id: vIds[0], type: 'vision_fixed' });
      else if (vIds.length > 1) itemsToAdd.push({ ids: vIds, type: 'vision_select' });
    }

    if (itemsToAdd.length === 0) return;
    
    sheet.insertRowsAfter(row, itemsToAdd.length);
    var startInsertRow = row + 1;
    for (var k = 0; k < itemsToAdd.length; k++) {
      var currentRow = startInsertRow + k;
      var item = itemsToAdd[k];

      if (item.type === 'child' || item.type === 'jig') {
        sheet.getRange(currentRow, 4).setValue(item.id);
        sheet.getRange(currentRow, 5).setValue(item.desc);
        sheet.getRange(currentRow, 4).clearDataValidations(); 
      } 
      else if (item.type === 'grandchild' || item.type === 'rubber_tip') {
        var rangeNotation = "REF_DATA!V" + item.refDataStart + ":V" + item.refDataEnd;
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(rangeNotation), true)
          .setAllowInvalid(true).build();
        sheet.getRange(currentRow, 4).setDataValidation(rule);
        var formulaB = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!Q:S, 2, FALSE), "")';
        sheet.getRange(currentRow, 2).setFormula(formulaB);
        var formulaE = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!Q:S, 3, FALSE), "")';
        sheet.getRange(currentRow, 5).setFormula(formulaE);
        if(item.type === 'rubber_tip') sheet.getRange(currentRow, 2).clearContent(); 
      }
      else if (item.type === 'vision_fixed') {
        sheet.getRange(currentRow, 4).setValue(item.id);
        sheet.getRange(currentRow, 2).setFormula('=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 3, FALSE), "")');
        sheet.getRange(currentRow, 5).setFormula('=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 2, FALSE), "")');
        sheet.getRange(currentRow, 4).clearDataValidations();
      }
      else if (item.type === 'vision_select') {
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(item.ids, true).setAllowInvalid(true).build();
        sheet.getRange(currentRow, 4).setDataValidation(rule);
        sheet.getRange(currentRow, 2).setFormula('=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 3, FALSE), "")');
        sheet.getRange(currentRow, 5).setFormula('=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 2, FALSE), "")');
      }

      if (item.type !== 'vision_fixed' && item.type !== 'vision_select') {
        sheet.getRange(currentRow, 3).clearContent();
      }
      sheet.getRange(currentRow, 7).insertCheckboxes();
      sheet.getRange(currentRow, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
    }
  }
}

// --- HELPER: Find Parent Config ---
function findParentConfig(refData, parentID) {
  for (var i = 0; i < refData.length; i++) {
    if (refData[i][0] == parentID) {
      return {
        elecIds: refData[i][20],
        elecDesc: refData[i][21],
        toolIds: refData[i][22],
        toolDesc: refData[i][23],
        jigIds: refData[i][24],
        jigDesc: refData[i][25],
        visionIds: refData[i][26],
        visionDesc: refData[i][27]
      };
    }
  }
  return null;
}
function getToolingOptionIDs(optionData, parentToolID) {
  var ids = [];
  for (var i = 0; i < optionData.length; i++) {
    if (optionData[i][0] == parentToolID) {
      if(optionData[i][1]) ids.push(optionData[i][1]);
    }
  }
  return ids;
}
function getToolingOptionRange(menuData, parentToolID) {
  var startRow = -1;
  var endRow = -1;
  for (var i = 0; i < menuData.length; i++) {
    if (menuData[i][0] == parentToolID) {
      if (startRow === -1) startRow = i + 1;
      endRow = i + 1;
    }
  }
  if (startRow !== -1) return { startRow: startRow, endRow: endRow };
  return null;
}
function insertShoppingList(sheet, row, count, dropdownRef, vlookupRef) {
  sheet.insertRowsAfter(row, count);
  var startInsertRow = row + 1;
  var dropDownRange = sheet.getRange(startInsertRow, 4, count, 1);
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRef), true).setAllowInvalid(true).build();
  dropDownRange.setDataValidation(rule);
  var descRange = sheet.getRange(startInsertRow, 5, count, 1);
  var formulas = [];
  for (var i = 0; i < count; i++) {
    formulas.push(['=IFERROR(VLOOKUP(D' + (startInsertRow + i) + ', ' + vlookupRef + ', 2, FALSE), "")']);
  }
  descRange.setFormulas(formulas);
  sheet.getRange(startInsertRow, 7, count, 1).insertCheckboxes();
  sheet.getRange(startInsertRow, 9, count, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
  sheet.getRange(startInsertRow, 6, count, 1).clearContent();
}

// =========================================
// 2. STANDARD MENUS & SYNC
// =========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Refresh') 
    .addItem('Sync REF_DATA (DB Only)', 'runMasterSync') // Updated Label
    .addSeparator()
    .addItem('Renumber Kits (Tidy Up)', 'renumberKits') 
    .addToUi();

  // New Menu for Configurator
  ui.createMenu('Configurator')
    .addItem('Open Sidebar', 'openSidebar')
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ModuleConfigurator')
      .setTitle('Module Configurator');
  SpreadsheetApp.getUi().showSidebar(html);
}

function renumberKits() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORDERING LIST");
  if (!sheet) return;
  var moduleFinder = sheet.getRange("A:A").createTextFinder("MODULE").matchEntireCell(true).findNext();
  var visionFinder = sheet.getRange("A:A").createTextFinder("VISION").matchEntireCell(true).findNext();
  if (!moduleFinder || !visionFinder) return;
  var startRow = moduleFinder.getRow() + 1;
  var endRow = visionFinder.getRow() - 1;
  var range = sheet.getRange(startRow, 4, endRow - startRow + 1, 1);
  var values = range.getValues();
  var refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  var refData = refSheet.getRange("C:AD").getValues();
  var parentCounts = {};
  for (var i = 0; i < values.length; i++) {
    var parentID = values[i][0];
    var config = findParentConfig(refData, parentID);
    if (config && config.elecIds) {
      if (!parentCounts[parentID]) parentCounts[parentID] = 0;
      parentCounts[parentID]++;
      var eIds = config.elecIds.split(';').map(function(s){ return s.trim(); });
      var eDescs = config.elecDesc.split(';').map(function(s){ return s.trim(); });
      var count = parentCounts[parentID];
      var index = (count - 1) % eIds.length;
      var targetId = eIds[index];
      var targetDesc = eDescs[index] || "";
      var childRowAbs = startRow + i + 1;
      if (childRowAbs > endRow + 10) continue;
      var actualChildID = sheet.getRange(childRowAbs, 4).getValue();
      if (eIds.includes(actualChildID)) {
        if (actualChildID !== targetId) {
           sheet.getRange(childRowAbs, 4).setValue(targetId);
           sheet.getRange(childRowAbs, 5).setValue(targetDesc);
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert("Renumbering Complete", "Sequential rotation updated.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// =========================================
// 4. MASTER SYNC LOGIC (REF_DATA ONLY)
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
    
    var destSS = SpreadsheetApp.getActiveSpreadsheet();
    var refSheet = destSS.getSheetByName("REF_DATA");
    // var destSheet = destSS.getSheetByName("ORDERING LIST"); // DISABLED to protect Ordering List
    
    // A. Update Reference Data (Preserving User Mappings in W-AD)
    updateReferenceData(sourceSS, sourceSheet);

    // B. Update Shopping Lists (REF_DATA Cols I-L)
    updateShoppingLists(sourceSheet);

    // C. PROCESS TOOLING OPTIONS (REF_DATA Cols P-V)
    processToolingOptions(sourceSS, refSheet);

    // D. PROCESS VISION DATA (REF_DATA Cols AF-AH)
    processVisionData(sourceSheet, refSheet);
    
    // E. Sync Main Sheet Sections - DISABLED TO PROTECT ORDERING LIST
    // updateSection_Core(sourceSheet, destSheet, "CORE", "CORE :430000-A557", 3, 4);
    // setupDropdownSection(destSheet, "CONFIG", "REF_DATA!A:A", "REF_DATA!A:B", null);
    // setupDropdownSection(destSheet, "MODULE", "REF_DATA!C:C", "REF_DATA!C:D", null);
    // setupVisionSection(destSheet);
    
    ui.alert("Sync Complete", "REF_DATA has been updated successfully.\n(ORDERING LIST was not touched)", ui.ButtonSet.OK);
  } catch (e) {
    console.error(e);
    ui.alert("Error during Sync", e.message, ui.ButtonSet.OK);
  }
}

// =========================================
// HELPER: TOOLING ILLUSTRATION PARSER
// =========================================
function processToolingOptions(sourceSS, refSheet) {
  var toolingSheet = sourceSS.getSheetByName("Tooling Illustration");
  if (!toolingSheet) { console.warn("Tooling Illustration sheet not found."); return; }
  var lastRow = toolingSheet.getLastRow();
  var rawData = toolingSheet.getRange(1, 1, lastRow, 8).getValues();
  var databaseOutput = [];
  var currentParentID = null;
  var currentCategory = null;
  for (var i = 0; i < rawData.length; i++) {
    var colA = String(rawData[i][0]).trim();
    var colB = String(rawData[i][1]).trim();
    var colF = String(rawData[i][5]).trim();
    var colH = String(rawData[i][7]).trim();
    var match = colA.match(/\[(.*?)\]/);
    if (match && match[1]) { currentParentID = match[1]; currentCategory = null; }
    if (colB !== "") { currentCategory = colB; }
    if (currentParentID && colF !== "" && colF !== "Part ID") {
      databaseOutput.push([currentParentID, colF, (currentCategory || ""), colH]);
    }
  }
  refSheet.getRange("P:S").clearContent();
  if (databaseOutput.length > 0) { refSheet.getRange(1, 16, databaseOutput.length, 4).setValues(databaseOutput); }
  var menuOutput = [];
  if (databaseOutput.length > 0) {
    var grouped = {};
    var orderParents = [];
    for (var k = 0; k < databaseOutput.length; k++) {
      var pID = databaseOutput[k][0];
      if (!grouped[pID]) { grouped[pID] = []; orderParents.push(pID); }
      grouped[pID].push({ partId: databaseOutput[k][1], cat: databaseOutput[k][2] });
    }
    for (var p = 0; p < orderParents.length; p++) {
      var parent = orderParents[p];
      var items = grouped[parent];
      var lastCat = null;
      for (var m = 0; m < items.length; m++) {
        var itm = items[m];
        var thisCat = itm.cat;
        if (thisCat !== "" && thisCat !== lastCat) { menuOutput.push([parent, "--- " + thisCat + " (Do Not Click) ---"]); lastCat = thisCat; }
        menuOutput.push([parent, itm.partId]);
      }
    }
  }
  refSheet.getRange("U:V").clearContent();
  if (menuOutput.length > 0) { refSheet.getRange(1, 21, menuOutput.length, 2).setValues(menuOutput); }
}

function updateReferenceData(sourceSS, sourceSheet) {
  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var refSheetName = "REF_DATA";
  var refSheet = destSS.getSheetByName(refSheetName);
  if (!refSheet) { refSheet = destSS.insertSheet(refSheetName); refSheet.hideSheet(); }
  var lastRefRow = refSheet.getLastRow();
  var existingMappings = {}; 
  if (lastRefRow > 0) {
    var currentData = refSheet.getRange(1, 3, lastRefRow, 28).getValues();
    for (var i = 0; i < currentData.length; i++) {
      var pId = currentData[i][0].toString().trim();
      if (pId !== "") {
        existingMappings[pId] = {
          eId: currentData[i][20], eDesc: currentData[i][21],
          tId: currentData[i][22], tDesc: currentData[i][23],
          jId: currentData[i][24], jDesc: currentData[i][25],
          vId: currentData[i][26], vDesc: currentData[i][27]
        };
      }
    }
  }
  refSheet.getRange("A:D").clear(); refSheet.getRange("W:AD").clear(); 
  var configItems = fetchRawItems(sourceSheet, "OPTIONAL MODULE: 430001-A712", 6, 7, ["CONFIGURABLE MODULE"]);
  var moduleItems = fetchRawItems(sourceSheet, "CONFIGURABLE MODULE: 430001-A713", 6, 7, ["CONFIGURABLE VISION MODULE"]);
  if (configItems.length > 0) refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems);
  if (moduleItems.length > 0) {
    var moduleOutput = []; var mappingOutput = [];
    for (var m = 0; m < moduleItems.length; m++) {
      var mId = moduleItems[m][0].toString().trim(); var mDesc = moduleItems[m][1];
      var eId = "", eDesc = "", tId = "", tDesc = "", jId = "", jDesc = "", vId = "", vDesc = "";
      if (existingMappings[mId]) {
        eId = existingMappings[mId].eId; eDesc = existingMappings[mId].eDesc;
        tId = existingMappings[mId].tId; tDesc = existingMappings[mId].tDesc;
        jId = existingMappings[mId].jId; jDesc = existingMappings[mId].jDesc;
        vId = existingMappings[mId].vId; vDesc = existingMappings[mId].vDesc;
      }
      moduleOutput.push([mId, mDesc]);
      mappingOutput.push([eId, eDesc, tId, tDesc, jId, jDesc, vId, vDesc]);
    }
    refSheet.getRange(1, 3, moduleOutput.length, 2).setValues(moduleOutput);
    refSheet.getRange(1, 23, mappingOutput.length, 8).setValues(mappingOutput);
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

function processVisionData(sourceSheet, refSheet) {
  var textFinder = sourceSheet.createTextFinder("CONFIGURABLE VISION MODULE").matchEntireCell(false);
  var found = textFinder.findNext();
  if (!found) { console.warn("Header 'CONFIGURABLE VISION MODULE' not found."); return; }
  var startRow = found.getRow() + 1;
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < startRow) return;
  var numRows = lastRow - startRow + 1;
  var rawData = sourceSheet.getRange(startRow, 5, numRows, 3).getValues();
  var databaseOutput = [];
  var menuOutput = [];
  var currentCategory = "Uncategorized";
  var groupedData = {};
  var orderCategories = [];
  for (var i = 0; i < rawData.length; i++) {
    var cat = rawData[i][0].toString().trim();
    var id = rawData[i][1].toString().trim();
    var desc = rawData[i][2].toString().trim();
    if (id === "" && cat === "") continue; 
    if (cat !== "") { currentCategory = cat; }
    if (id !== "" && id !== "Part Number") { 
       databaseOutput.push([id, desc, currentCategory]);
       if (!groupedData[currentCategory]) { groupedData[currentCategory] = []; orderCategories.push(currentCategory); }
       groupedData[currentCategory].push(id);
    }
  }
  refSheet.getRange(1, 32, refSheet.getMaxRows(), 3).clearContent();
  refSheet.getRange(1, 36, refSheet.getMaxRows(), 2).clearContent();
  if (databaseOutput.length > 0) { refSheet.getRange(1, 32, databaseOutput.length, 3).setValues(databaseOutput); }
  if (orderCategories.length > 0) {
    for (var c = 0; c < orderCategories.length; c++) {
      var cName = orderCategories[c];
      menuOutput.push(["--- " + cName + " ---", ""]);
      var ids = groupedData[cName];
      for (var k = 0; k < ids.length; k++) { menuOutput.push(["", ids[k]]); }
    }
    if (menuOutput.length > 0) { refSheet.getRange(1, 36, menuOutput.length, 2).setValues(menuOutput); }
  }
}

// These functions were removed from runMasterSync but logic kept if needed
function updateSection_Core(sourceSheet, destSheet, destHeaderName, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc) {
  var rawItems = fetchRawItems(sourceSheet, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc, []);
  var syncItems = rawItems.map(function(item) { return [item[0], item[1], "1"]; });
  performSurgicalSync(destSheet, [{ destName: destHeaderName, items: syncItems }]);
}
function setupDropdownSection(destSheet, sectionName, dropdownRangeString, vlookupRangeString, categoryColIndex) {
  // Legacy code suppressed
}
function setupVisionSection(sheet) {
  // Legacy code suppressed
}
function performSurgicalSync(destSheet, sections) {
  // Legacy code suppressed
}
