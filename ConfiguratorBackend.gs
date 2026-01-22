/**
 * ConfiguratorBackend.gs
 * Server-side logic for the Module Configurator UI.
 * UPDATED: Phase 6 (Machine Setup - Base Module Tooling Flattened List)
 */

// --- CONSTANTS ---
var RUBBER_TIP_PARENTS_BACKEND = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
var RUBBER_TIP_SOURCE_ID_BACKEND = "430001-A380";

// NEW: Vision PC Target IDs
var VISION_PC_IDS = ["430001-A366", "430001-A367"];

/**
 * 1. Get List of Modules (ID and Description)
 */
function getModuleList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  var rawValues = sheet.getRange(1, 3, lastRow, 2).getValues();
  var modules = [];
  
  for (var i = 0; i < rawValues.length; i++) {
    var id = String(rawValues[i][0]).trim();
    var desc = String(rawValues[i][1]).trim();
    
    if (id !== "" && id !== "Part ID") {
      modules.push({ id: id, desc: desc });
    }
  }
  return modules;
}

/**
 * 1.5 Get Vision PC Options (Hybrid Strategy)
 * Fetches descriptions for specific hardcoded IDs from REF_DATA
 */
function getVisionPCOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  // Columns A & B are Part ID & Description (Cols 1 & 2)
  var rawValues = sheet.getRange(1, 1, lastRow, 2).getValues();
  
  var options = [];
  
  // Create a quick lookup map for performance
  var lookupMap = {};
  for (var i = 0; i < rawValues.length; i++) {
    var rId = String(rawValues[i][0]).trim();
    var rDesc = String(rawValues[i][1]).trim();
    if (rId) lookupMap[rId] = rDesc;
  }
  
  // Build result based on HARDCODED Target IDs
  for (var j = 0; j < VISION_PC_IDS.length; j++) {
    var targetID = VISION_PC_IDS[j];
    var foundDesc = lookupMap[targetID] || "Description not found";
    options.push({ id: targetID, desc: foundDesc });
  }
  
  return options;
}

/**
 * 1.6 Get Configurable Base Module Options
 * Fetches from REF_DATA Columns I & J
 */
function getBaseModuleOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  // Columns I (9) and J (10)
  var rawValues = sheet.getRange(1, 9, lastRow, 2).getValues();
  var options = [];

  for (var i = 0; i < rawValues.length; i++) {
    var id = String(rawValues[i][0]).trim();
    var desc = String(rawValues[i][1]).trim();
    
    // Filter Logic:
    // 1. Exclude empty IDs
    // 2. Exclude "Part ID" header
    // 3. Exclude "List-" or "---" headers
    if (id && id !== "Part ID" && !id.toUpperCase().startsWith("LIST-") && id.indexOf("---") === -1) {
       options.push({ id: id, desc: desc });
    }
  }
  return options;
}

/**
 * 1.7 Get Base Module Tooling List (Non-Paired Only)
 * Logic: Fetch U:V, Exclude IDs found in Y:Z
 */
function getBaseModuleToolingList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  // 1. Fetch Data
  // U:V = Cols 21-22 (Tooling Source)
  // Y:Z = Cols 25-26 (Exclusion List)
  var toolingRange = sheet.getRange(1, 21, lastRow, 2).getValues(); // U:V
  var exclusionRange = sheet.getRange(1, 25, lastRow, 2).getValues(); // Y:Z

  // 2. Build Exclusion Set
  var excludedIDs = new Set();
  for (var i = 0; i < exclusionRange.length; i++) {
    var row = exclusionRange[i];
    // Y (row[0]) and Z (row[1])
    [row[0], row[1]].forEach(function(cellVal) {
      if (cellVal) {
        var str = String(cellVal).trim();
        // Handle delimited IDs (e.g., "A; B; C")
        var parts = str.split(/[;\n\r]+/); // Split by semicolon or newline
        parts.forEach(function(p) {
          var clean = p.trim();
          if (clean) excludedIDs.add(clean);
        });
      }
    });
  }

  // 3. Process Tooling List (U:V)
  // We need to group by Parent ID (Col U) to avoid duplicates if U repeats
  // But wait, fetchOptionsForTool handles the children. We just need the unique Parents.
  var uniqueTools = [];
  var seenParents = new Set();
  
  // We also need the full menu data for fetchOptionsForTool later
  var menuData = toolingRange; // Reuse the U:V values

  for (var j = 0; j < toolingRange.length; j++) {
    var parentID = String(toolingRange[j][0]).trim();
    var rawDescField = String(toolingRange[j][1]).trim(); // Col V

    // Filter Logic
    if (!parentID || parentID === "Part ID" || parentID.indexOf("---") > -1 || parentID.toUpperCase().startsWith("LIST-")) {
      continue;
    }
    
    // EXCLUSION LOGIC
    if (excludedIDs.has(parentID)) {
      continue;
    }
    
    if (seenParents.has(parentID)) {
      continue; // Already processed this parent
    }
    seenParents.add(parentID);

    // 4. Parse Description from Col V
    // Format: "OptionID :: Description" OR just "Description"
    var finalDesc = rawDescField;
    if (rawDescField.indexOf("::") > -1) {
       var parts = rawDescField.split("::");
       if (parts.length > 1) {
          finalDesc = parts[1].trim(); // Take the description part
       }
    }
    
    // 5. Fetch Options
    // Check if this tool has dropdown options
    var options = fetchOptionsForTool(menuData, parentID);
    
    uniqueTools.push({
       id: parentID,
       desc: finalDesc,
       options: options
    });
  }

  return uniqueTools;
}


/**
 * 2. Get Full Details for Selected Module
 */
function getModuleDetails(moduleID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName("REF_DATA");
  var dataRange = refSheet.getRange("C:AF").getValues();
  var rowData = null;
  
  for (var i = 0; i < dataRange.length; i++) {
    if (String(dataRange[i][0]) === moduleID) {
      rowData = dataRange[i];
      break;
    }
  }
  
  if (!rowData) return { error: "Module ID not found." };
  
  // --- ELECTRICAL ---
  var elecResult = null;
  var elecIdsStr = rowData[20];
  var elecDescStr = rowData[21];
  
  if (elecIdsStr) {
    var eIds = String(elecIdsStr).split(';').map(function(s){ return s.trim(); });
    var eDescs = String(elecDescStr).split(';').map(function(s){ return s.trim(); });
    
    if (eIds.length > 0) {
      elecResult = {
        id: eIds[0],
        desc: eDescs[0] || "",
        note: "Instance 1 (Rotation A)"
      };
    }
  }
  
  // --- TOOLING ---
  var toolsResult = [];
  var toolIdsStr = rowData[22];
  var toolDescStr = rowData[23];
  
  if (toolIdsStr) {
    var tIds = String(toolIdsStr).split(';').map(function(s){ return s.trim(); });
    var tDescs = String(toolDescStr).split(';').map(function(s){ return s.trim(); });
    var menuData = refSheet.getRange("U:V").getValues();
    
    for (var k = 0; k < tIds.length; k++) {
      var tID = tIds[k];
      if (!tID) continue;
      
      var toolObj = {
        id: tID,
        desc: tDescs[k] || "",
        standardOptions: [],
        rubberTipOptions: [],
        requiresRubberTip: false
      };
      
      var stdOpts = fetchOptionsForTool(menuData, tID);
      if (stdOpts.length > 0) toolObj.standardOptions = stdOpts;
      
      if (RUBBER_TIP_PARENTS_BACKEND.includes(tID)) {
        var tipOpts = fetchOptionsForTool(menuData, RUBBER_TIP_SOURCE_ID_BACKEND);
        if (tipOpts.length > 0) {
          toolObj.requiresRubberTip = true;
          toolObj.rubberTipOptions = tipOpts;
        }
      }
      toolsResult.push(toolObj);
    }
  }
  
  // --- VISION ---
  var visionResult = { type: 'none', options: [] };
  var visionIdsStr = rowData[26];
  var visionDescStr = rowData[27];
  
  if (visionIdsStr) {
    var vIds = String(visionIdsStr).split(';').map(function(s){ return s.trim(); });
    var vDescs = String(visionDescStr).split(';').map(function(s){ return s.trim(); });
    
    var vOptions = [];
    for(var v=0; v<vIds.length; v++){
      if(vIds[v]) {
        vOptions.push({ id: vIds[v], desc: vDescs[v] || "" });
      }
    }

    if (vOptions.length === 1) {
      visionResult.type = 'fixed';
      visionResult.options = vOptions;
    } else if (vOptions.length > 1) {
      visionResult.type = 'select';
      visionResult.options = vOptions;
    }
  }
  
  // --- JIGS ---
  var jigsResult = [];
  var jigIdsStr = rowData[24];
  var jigDescStr = rowData[25];
  if (jigIdsStr) {
    var jIds = String(jigIdsStr).split(';').map(function(s){ return s.trim(); });
    var jDescs = String(jigDescStr).split(';').map(function(s){ return s.trim(); });
    for (var j = 0; j < jIds.length; j++) {
      if(jIds[j]) jigsResult.push({ id: jIds[j], desc: jDescs[j] || "" });
    }
  }

  // --- SPARE PARTS ---
  var sparePartsResult = [];
  var spIdsStr = rowData[28]; 
  var spDescStr = rowData[29]; 
  
  if (spIdsStr) {
    var spIds = String(spIdsStr).split(';').map(function(s){ return s.trim(); });
    var spDescs = String(spDescStr).split(';').map(function(s){ return s.trim(); });
    
    for (var s = 0; s < spIds.length; s++) {
      if (spIds[s]) {
        sparePartsResult.push({ id: spIds[s], desc: spDescs[s] || "" });
      }
    }
  }

  return {
    moduleID: moduleID,
    electrical: elecResult,
    tools: toolsResult,
    jigs: jigsResult,
    vision: visionResult,
    spareParts: sparePartsResult
  };
}

function fetchOptionsForTool(menuData, parentID) {
  var options = [];
  var foundParent = false;
  var currentCategory = "Standard Options"; 

  for (var i = 0; i < menuData.length; i++) {
    var rowParent = String(menuData[i][0]);
    var rowChild = String(menuData[i][1]); 

    if (rowParent === parentID) {
      foundParent = true;
      if (rowChild) {
        if (rowChild.indexOf("---") > -1) {
           currentCategory = rowChild.replace(/---/g, '').trim();
        } 
        else {
           var parts = rowChild.split('::');
           var id = parts[0].trim();
           var desc = "";
           if (parts.length > 1) desc = parts[1].trim();
           
           options.push({ id: id, desc: desc, category: currentCategory }); 
        }
      }
    } else if (foundParent) {
      if (rowParent !== "") break; 
    }
  }
  return options;
}

/**
 * 2.5 Get Layout Header Labels
 * UPDATED Phase 6: Dynamic Anchor Scanning
 */
function getLayoutHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) return null;

  // Find VCM Anchor Row dynamically
  var vcmFinder = sheet.getRange("B:B").createTextFinder("VCM").matchEntireCell(true).findNext();
  if (!vcmFinder) {
     throw new Error("Critical Error: 'VCM' anchor not found in Column B of Trial Layout. Please ensure headers exist.");
  }
  var vcmRowIndex = vcmFinder.getRow();

  // Read Headers (Anchor Row) and Options (Row below Anchor)
  // B:F = 5 columns
  var rawData = sheet.getRange(vcmRowIndex, 2, 2, 5).getValues();
  var rowCategories = rawData[0]; // Header Row
  var rowOptions = rawData[1];    // Option Row
  
  return {
    group1: {
      title: rowCategories[0], // B "VCM"
      options: [
        { colIndex: 2, label: rowOptions[0] }, // B
        { colIndex: 3, label: rowOptions[1] }  // C
      ]
    },
    group2: {
      title: rowCategories[2], // D "VALVE SET"
      options: [
        { colIndex: 4, label: rowOptions[2] }, // D
        { colIndex: 5, label: rowOptions[3] }, // E
        { colIndex: 6, label: rowOptions[4] }  // F
      ]
    }
  };
}

/**
 * 3. SAVE CONFIGURATION TO TRIAL LAYOUT
 * UPDATED Phase 6: "Look Before You Leap" Strategy (Correctly finds B10)
 */
function saveConfiguration(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Sheet 'TRIAL-LAYOUT CONFIGURATION' not found.");

  var lastRow = sheet.getLastRow();
  
  // STEP 1: Find the "CONFIGURATION" Fence
  // This header marks the absolute start of the list. We MUST write below it (or at the same level if merged).
  var configHeaderFinder = sheet.getRange("A:B").createTextFinder("CONFIGURATION").matchEntireCell(true).findNext();
  
  if (!configHeaderFinder) {
    throw new Error("Critical Error: 'CONFIGURATION' header not found in Column A or B. Cannot safely locate list start.");
  }

  // Use the exact row where "CONFIGURATION" is found.
  // Because "CONFIGURATION" (Col A) and "B10" (Col G) are on the same row,
  // we start scanning AT this row, not +1.
  var startScanRow = configHeaderFinder.getRow();
  
  if (startScanRow > lastRow) {
     // Safety check if sheet is malformed
  }

  // Scan Column G (Turret Location) starting from the header row downwards
  var scanRange = sheet.getRange(startScanRow, 7, lastRow - startScanRow + 1, 2).getValues();

  var targetRowIndex = -1;
  var startScanning = false;

  for (var i = 0; i < scanRange.length; i++) {
    var slotID = String(scanRange[i][0]).trim(); // Col G
    var desc = String(scanRange[i][1]).trim();   // Col H

    // STRICT VALIDATION: Is this a "B" slot?
    // This prevents writing to random empty cells if logic drifts.
    // e.g. "B10", "B11" matches. "B" or "Box" does not.
    var isBSlot = slotID.startsWith("B") && !isNaN(parseInt(slotID.substring(1)));

    // Case 1: We haven't found the first B10 yet.
    if (!startScanning) {
      if (isBSlot) {
        startScanning = true; // Found the first valid slot (e.g., B10)
        // CRITICAL: Do NOT 'continue'. We must process THIS row immediately.
      }
    }

    // Case 2: We are inside the list (Processing B10, B11, etc.)
    if (startScanning) {
      if (slotID === "B33" || (slotID.startsWith("B") && parseInt(slotID.substring(1)) > 32)) {
         break; // End of list (B33 is stop)
      }

      // If desc is empty, we found our target
      if (desc === "") {
        targetRowIndex = startScanRow + i; // Convert relative index i back to absolute row
        break; 
      }
    }
  }

  if (targetRowIndex === -1) {
    throw new Error("Configuration List is full (B10-B32) or could not be located at 'CONFIGURATION' header.");
  }

  // WRITE DATA (Same as before)
  sheet.getRange(targetRowIndex, 8).setValue(payload.moduleDesc);
  sheet.getRange(targetRowIndex, 9).setValue(payload.numberVision || 1);
  sheet.getRange(targetRowIndex, 11).setValue(payload.moduleID);
  sheet.getRange(targetRowIndex, 12).setValue(payload.elecID);
  sheet.getRange(targetRowIndex, 13).setValue(payload.visionID);
  sheet.getRange(targetRowIndex, 14).setValue(payload.toolOptionID);
  sheet.getRange(targetRowIndex, 15).setValue(payload.rubberTipID);
  sheet.getRange(targetRowIndex, 16).setValue(payload.jigID);
  sheet.getRange(targetRowIndex, 17).setValue(payload.sparePartsID);
  
  if (payload.layoutFlags) {
     for (var col = 2; col <= 6; col++) {
        var val = payload.layoutFlags[String(col)] === true; 
        var cell = sheet.getRange(targetRowIndex, col);
        var rule = cell.getDataValidation();
        var isCheckbox = (rule != null && rule.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX);
        if (!isCheckbox) { cell.insertCheckboxes(); }
        cell.setValue(val);
     }
  }

  var slotLabel = sheet.getRange(targetRowIndex, 7).getValue(); 
  return { status: "success", slot: slotLabel }; 
}

/**
 * 3.5 SAVE MACHINE SETUP (Phase 6)
 * Writes Vision PC, Configurable Base Module, AND Base Module Tooling.
 * Uses "Smart Insert" to manage variable list lengths.
 */
function saveMachineSetup(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Sheet 'TRIAL-LAYOUT CONFIGURATION' not found.");
  
  // 1. Save Vision PC (C2)
  if (payload.visionPC) {
    sheet.getRange(2, 3).setValue(payload.visionPC);
  }
  
  // 2. Save Configurable Base Modules
  if (payload.baseModules && Array.isArray(payload.baseModules)) {
    var startFinder = sheet.getRange("A:B").createTextFinder("Configurable Base Module").matchEntireCell(false).findNext();
    var endFinder = sheet.getRange("A:B").createTextFinder("Base Module Tooling").matchEntireCell(false).findNext();

    if (startFinder && endFinder) {
      var startRow = startFinder.getRow();
      var nextSectionRow = endFinder.getRow();
      var currentGap = nextSectionRow - startRow; 
      var itemsToSave = payload.baseModules;
      var requiredSlots = Math.max(itemsToSave.length, 3); 
      
      if (requiredSlots > currentGap) {
         var rowsNeeded = requiredSlots - currentGap;
         sheet.insertRowsBefore(nextSectionRow, rowsNeeded);
         var newRowsRange = sheet.getRange(nextSectionRow, 1, rowsNeeded, sheet.getLastColumn());
         newRowsRange.setTextRotation(0).setVerticalAlignment("middle");
      } else if (requiredSlots < currentGap) {
         var rowsToDelete = currentGap - requiredSlots;
         if (rowsToDelete > 0) sheet.deleteRows(startRow + requiredSlots, rowsToDelete);
      }
      
      for (var i = 0; i < requiredSlots; i++) {
         var targetRow = startRow + i;
         var val = (i < itemsToSave.length) ? itemsToSave[i] : ""; 
         sheet.getRange(targetRow, 3).setValue(val);
         var mergeRange = sheet.getRange(targetRow, 3, 1, 5);
         try {
           if (!mergeRange.isPartOfMerge() || mergeRange.getMergedRanges().length > 1) { mergeRange.merge(); }
           mergeRange.setVerticalAlignment("middle");
         } catch(e) {}
      }
    }
  }

  // 3. Save Base Module Tooling (NEW - FLATTENED LIST with TREE STYLE)
  if (payload.baseTooling && Array.isArray(payload.baseTooling)) {
    // A. Locate Anchors
    var startFinder = sheet.getRange("A:B").createTextFinder("Base Module Tooling").matchEntireCell(false).findNext();
    var endFinder = sheet.getRange("A:B").createTextFinder("Comment").matchEntireCell(false).findNext();

    if (startFinder && endFinder) {
      var startRow = startFinder.getRow();
      var nextSectionRow = endFinder.getRow();
      var currentGap = nextSectionRow - startRow;
      
      // B. Flatten Payload for writing
      var linesToWrite = [];
      for (var p = 0; p < payload.baseTooling.length; p++) {
         var item = payload.baseTooling[p];
         // 1. Parent ID (Row 1)
         linesToWrite.push(item.id);
         // 2. Child ID (Row 2, Indented)
         if (item.option) {
            linesToWrite.push("L " + item.option); // Tree Style
         }
      }
      
      var requiredSlots = linesToWrite.length;
      if (requiredSlots === 0) requiredSlots = 1;

      // C. Smart Insert / Delete Rows
      if (requiredSlots > currentGap) {
         var rowsNeeded = requiredSlots - currentGap;
         sheet.insertRowsBefore(nextSectionRow, rowsNeeded);
         // Apply formatting to new rows
         var newRowsRange = sheet.getRange(nextSectionRow, 1, rowsNeeded, sheet.getLastColumn());
         newRowsRange.setTextRotation(0).setVerticalAlignment("middle");
      } else if (requiredSlots < currentGap) {
         var rowsToDelete = currentGap - requiredSlots;
         if (rowsToDelete > 0) sheet.deleteRows(startRow + requiredSlots, rowsToDelete);
      }
      
      // D. Write Data
      for (var k = 0; k < requiredSlots; k++) {
         var targetRow = startRow + k;
         
         if (k < linesToWrite.length) {
           sheet.getRange(targetRow, 3).setValue(linesToWrite[k]);
         } else {
           sheet.getRange(targetRow, 3).clearContent();
         }
         
         // Merge C:G
         var mergeRange = sheet.getRange(targetRow, 3, 1, 5);
         try {
           if (!mergeRange.isPartOfMerge() || mergeRange.getMergedRanges().length > 1) { mergeRange.merge(); }
           mergeRange.setVerticalAlignment("middle");
         } catch(e) {}
      }
    }
  }
  
  return { status: "success" };
}

// =======================================================
// PHASE 4.2: EXTRACTION LOGIC ("The Brain")
// =======================================================

function buildMasterDictionary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName("REF_DATA");
  if (!refSheet) return {};

  var lastRow = refSheet.getLastRow();
  var dictionary = {};

  function addToDict(id, desc) {
    if (id && id !== "" && id !== "Part ID") {
      var cleanId = String(id).trim();
      if (!dictionary[cleanId]) { 
        dictionary[cleanId] = String(desc).trim();
      }
    }
  }

  // 1. Modules
  var moduleData = refSheet.getRange(1, 3, lastRow, 2).getValues();
  for (var i = 0; i < moduleData.length; i++) {
    addToDict(moduleData[i][0], moduleData[i][1]);
  }

  // 2. Tooling Menu
  var menuData = refSheet.getRange(1, 21, lastRow, 2).getValues();
  for (var j = 0; j < menuData.length; j++) {
    var packed = String(menuData[j][1]); 
    if (packed.indexOf("::") > -1) {
      var parts = packed.split("::");
      addToDict(parts[0], parts[1]);
    }
  }

  // 3. Manual Mapping Block
  var manualData = refSheet.getRange(1, 23, lastRow, 10).getValues(); 
  for (var k = 0; k < manualData.length; k++) {
    var row = manualData[k];
    for (var p = 0; p < 10; p += 2) {
       var rawId = row[p];
       var rawDesc = row[p+1];
       if (rawId) {
          var ids = String(rawId).split(";");
          var descs = String(rawDesc).split(";");
          for (var x = 0; x < ids.length; x++) {
             var cleanId = ids[x].trim();
             var cleanDesc = (descs[x] || "").trim();
             addToDict(cleanId, cleanDesc);
          }
       }
    }
  }
  
  return dictionary;
}

/**
 * B. Extract Data for Production
 * UPDATED PHASE 5 (Item Number Logic Change)
 */
function extractProductionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Staging Sheet Missing");

  var masterDict = buildMasterDictionary();
  var lastRow = sheet.getLastRow();
  
  // UPDATED PHASE 6: Dynamic Header Reading
  var vcmFinder = sheet.getRange("B:B").createTextFinder("VCM").matchEntireCell(true).findNext();
  if (!vcmFinder) throw new Error("Critical Error: 'VCM' anchor not found. Cannot extract configuration.");
  
  var vcmRowIndex = vcmFinder.getRow();
  
  // Read the Header Row (e.g., "Voice Coil (Direct)...") which is at vcmRowIndex
  // B:F (Cols 2-6)
  var headerData = sheet.getRange(vcmRowIndex, 2, 1, 5).getValues()[0];
  var headerRow = [null, headerData[0], headerData[1], headerData[2], headerData[3], headerData[4]]; 
  
  // Data Range: From VCM Row + 2 downwards
  var startDataRow = vcmRowIndex + 2;
  var rawData = sheet.getRange(startDataRow, 1, lastRow - startDataRow + 1, 18).getValues();

  var payload = {
    MODULE: [],
    ELECTRICAL: [],
    VISION: [],
    TOOLING: [],
    JIG: [], 
    SPARES: [],
    VCM: [],
    OTHERS: [], 
    rowsToMarkSynced: [] 
  };

  var startScanning = false;
  
  for (var i = 0; i < rawData.length; i++) {
    var row = rawData[i];
    var turretName = String(row[6]).trim(); // Col G
    var syncStatus = String(row[9]).trim(); // Col J
    var moduleID = String(row[10]).trim();  // Col K
    
    // Dynamic start check
    if (!startScanning && turretName.startsWith("B") && !isNaN(parseInt(turretName.substring(1)))) {
       startScanning = true;
    }
    
    if (turretName === "B33" || (turretName.startsWith("B") && parseInt(turretName.substring(1)) > 32)) {
      break;
    }

    if (startScanning) {
      if (moduleID === "") continue; 
      if (syncStatus === "SYNCED") continue; 

      payload.rowsToMarkSynced.push({row: startDataRow + i}); // Use absolute row
      
      var globalQty = parseInt(row[8]); 
      if (isNaN(globalQty) || globalQty < 1) globalQty = 1;
      
      // Removed hardcoded itemIndex logic

      function pushItem(category, id, descOverride, isPrimary) {
        if (!id || id === "") return;
        var cleanId = id.trim();
        var finalDesc = descOverride || masterDict[cleanId] || "Check REF_DATA";
        
        payload[category].push({
          requiresNumbering: isPrimary, // Boolean flag instead of number
          id: cleanId,
          desc: finalDesc,
          qty: globalQty 
        });
      }

      pushItem("MODULE", moduleID, String(row[7]), true);
      pushItem("ELECTRICAL", String(row[11]), null, true);
      pushItem("VISION", String(row[12]), null, true);

      // Tooling
      pushItem("TOOLING", String(row[13]), null, true); 
      pushItem("TOOLING", String(row[14]), null, false); // Secondary

      pushItem("JIG", String(row[15]), null, true);

      // Spares
      var sparesRaw = String(row[16]);
      if (sparesRaw) {
        var spareIds = sparesRaw.split(";");
        for (var s = 0; s < spareIds.length; s++) {
          var sId = spareIds[s].trim();
          pushItem("SPARES", sId, null, (s === 0)); // Only first is primary
        }
      }

      function parseHeaderData(headerText) {
        var match = headerText.match(/(\d{4,}-\w+)/);
        if (match) {
           var id = match[0];
           var desc = headerText.replace(id, "").trim();
           return { id: id, desc: desc };
        }
        return { id: "", desc: "" };
      }

      // VCM (Indices shifted because row array is 0-indexed)
      // row[1] = Col B
      if (row[1] === true) {
         var d = parseHeaderData(headerRow[1]);
         pushItem("VCM", d.id, d.desc, true); 
      }
      if (row[2] === true) {
         var d = parseHeaderData(headerRow[2]);
         pushItem("VCM", d.id, d.desc, true);
      }

      // OTHERS
      if (row[3] === true) {
         var d = parseHeaderData(headerRow[3]);
         pushItem("OTHERS", d.id, d.desc, true);
      }
      if (row[4] === true) {
         var d = parseHeaderData(headerRow[4]);
         pushItem("OTHERS", d.id, d.desc, true);
      }
      if (row[5] === true) {
         var d = parseHeaderData(headerRow[5]);
         pushItem("OTHERS", d.id, d.desc, true);
      }
    } 
  } 

  return payload;
}

// =======================================================
// PHASE 4.3: PRODUCTION INJECTION & STATUS UPDATE
// =======================================================

function injectProductionData(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ORDERING LIST");
  if (!sheet) throw new Error("ORDERING LIST Sheet Missing");

  var itemsAdded = 0;

  var sections = [
    { key: "MODULE", header: "MODULE" },
    { key: "ELECTRICAL", header: "ELECTRICAL" },
    { key: "VISION", header: "VISION" },
    { key: "VCM", header: "VCM" },
    { key: "OTHERS", header: "OTHERS" },
    { key: "SPARES", header: "SPARES" },
    { key: "JIG", header: "JIG/CALIBRATION" },
    { key: "TOOLING", header: "TOOLING" } 
  ];

  for (var i = 0; i < sections.length; i++) {
    var sec = sections[i];
    var items = payload[sec.key];
    if (items && items.length > 0) {
      itemsAdded += insertRowsIntoSection(sheet, sec.header, items);
    }
  }

  return itemsAdded;
}

/**
 * D. The Smart Fill Helper - UPDATED PHASE 5 (New Columns G, H, I)
 */
function insertRowsIntoSection(sheet, sectionHeader, items) {
  var lastRow = sheet.getLastRow();
  var rangeValues = sheet.getRange("A1:D" + lastRow).getValues();

  var startRowIndex = -1;
  var anchorRowIndex = -1;

  // 1. Find Start Header
  for (var i = 0; i < rangeValues.length; i++) {
    if (String(rangeValues[i][0]).trim().toUpperCase() === sectionHeader.toUpperCase()) {
      startRowIndex = i;
      break;
    }
  }

  if (startRowIndex === -1) {
    console.warn("Section Header '" + sectionHeader + "' not found in ORDERING LIST.");
    return 0; 
  }

  // 2. Find Next Header (Anchor)
  for (var j = startRowIndex + 1; j < rangeValues.length; j++) {
    var cellVal = String(rangeValues[j][0]).trim(); 
    if (cellVal !== "") {
       anchorRowIndex = j;
       break;
    }
  }

  if (anchorRowIndex === -1) {
     anchorRowIndex = rangeValues.length; 
  }

  // --- Find Current Max Item Number in this Zone ---
  var currentMaxNum = 0;
  for (var r = startRowIndex + 1; r < anchorRowIndex; r++) {
     var val = rangeValues[r][2]; // Column C
     if (typeof val === 'number') {
        if (val > currentMaxNum) currentMaxNum = val;
     } else if (val) {
        var parsed = parseInt(val);
        if (!isNaN(parsed) && parsed > currentMaxNum) {
           currentMaxNum = parsed;
        }
     }
  }

  // 3. Find Write Cursor
  var writeCursorIndex = -1;

  for (var k = startRowIndex + 1; k < anchorRowIndex; k++) {
    var partId = String(rangeValues[k][3]).trim();
    var itemNum = String(rangeValues[k][2]).trim();

    if (partId === "" && itemNum === "") {
      writeCursorIndex = k;
      break;
    }
  }

  if (writeCursorIndex === -1) {
    writeCursorIndex = anchorRowIndex;
  }

  // 4. Calculate Capacity & Deficit
  var availableSlots = anchorRowIndex - writeCursorIndex;
  var itemsNeeded = items.length;
  var deficit = itemsNeeded - availableSlots;

  // 5. Expand if necessary
  if (deficit > 0) {
    sheet.insertRowsBefore(anchorRowIndex + 1, deficit);
  }

  // 6. Write Data (Expanded for Phase 5)
  var startWriteRow = writeCursorIndex + 1;
  var output = [];
  var numberingCounter = currentMaxNum;

  for (var x = 0; x < items.length; x++) {
    var itemLabel = "";
    if (items[x].requiresNumbering) {
       numberingCounter++;
       itemLabel = numberingCounter;
    }

    output.push([
      itemLabel,    // Col C: Item No
      items[x].id,  // Col D: Part ID
      items[x].desc,// Col E: Description
      items[x].qty, // Col F: Qty
      false,        // Col G: Released (Checkbox Default False)
      "",           // Col H: Date (Empty)
      ""            // Col I: Type (Empty)
    ]);
  }

  // Write 7 columns (C through I)
  sheet.getRange(startWriteRow, 3, itemsNeeded, 7).setValues(output);

  // 7. Apply Data Validation (Checkbox & Dropdown)
  // Apply Checkbox to Col G
  sheet.getRange(startWriteRow, 7, itemsNeeded, 1).insertCheckboxes();

  // Apply Dropdown to Col I
  var typeRange = sheet.getRange(startWriteRow, 9, itemsNeeded, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CHARGE OUT', 'MRP'], true)
    .setAllowInvalid(true)
    .build();
  typeRange.setDataValidation(rule);

  return itemsNeeded;
}

function markRowsAsSynced(rowsToSync) {
  if (!rowsToSync || rowsToSync.length === 0) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) return;

  for (var i = 0; i < rowsToSync.length; i++) {
     var r = rowsToSync[i].row;
     sheet.getRange(r, 10).setValue("SYNCED");
  }
}

/**
 * OPTIONAL: Retroactive Fix Tool (Option B support tool, kept for safety)
 */
function initializeReleaseColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ORDERING LIST");
  var lastRow = sheet.getLastRow();
  
  // Apply to G:G (Checkboxes)
  sheet.getRange(7, 7, lastRow - 6, 1).insertCheckboxes();
  
  // Apply to I:I (Dropdowns)
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CHARGE OUT', 'MRP'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(7, 9, lastRow - 6, 1).setDataValidation(rule);
  
  SpreadsheetApp.getUi().alert("Columns Initialized.");
}
