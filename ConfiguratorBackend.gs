/**
 * ConfiguratorBackend.gs
 * Server-side logic for the Module Configurator UI.
 * UPDATED: Phase 6 (Machine Setup - 2-Level Flattened Hierarchy & Sync Update)
 * FIX: "Collapse & Rebuild" strategy + Strict Sanitation to eliminate duplication bugs.
 * FIX: Robust Anchor Detection (Regex) + Safety Check for Anchor Order.
 */

// --- CONSTANTS ---
var RUBBER_TIP_PARENTS_BACKEND = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
var RUBBER_TIP_SOURCE_ID_BACKEND = "430001-A380";

// Vision PC Target IDs
var VISION_PC_IDS = ["430001-A366", "430001-A367"];

// Complex Tooling Rules (Parent ID -> Mandatory Child ID)
var COMPLEX_TOOL_RULES = {
  "430001-A490": "430000-A748" // PUH Interface -> Assy-PUH Phoenix-v2.0
};

// Multi-Select Tools (Parent IDs that allow multiple option rows)
var MULTI_SELECT_TOOLS = ["430001-A380", "430001-A495"]; 

// NEW: Categorized Multi Tools (Complex Grouping)
var CATEGORIZED_MULTI_TOOLS = ["430001-A494"]; // Die Ejector Tooling

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
 * 1.7 Get Base Module Tooling List (Hybrid Interceptor Logic)
 * Logic: Fetch U:V, Exclude IDs found in Y:Z, Apply Complex Rules, Multi-Select & Categorized Logic
 */
function getBaseModuleToolingList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  // 1. Fetch Data
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
        var parts = str.split(/[;\n\r]+/);
        parts.forEach(function(p) {
          var clean = p.trim();
          if (clean) excludedIDs.add(clean);
        });
      }
    });
  }

  // 3. Process Tooling List (U:V)
  var uniqueTools = [];
  var seenParents = new Set();
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
    var finalDesc = rawDescField;
    if (rawDescField.indexOf("::") > -1) {
       var parts = rawDescField.split("::");
       if (parts.length > 1) {
          finalDesc = parts[1].trim(); // Take the description part
       }
    }
    
    // 5. Fetch Options & Apply Interceptor
    
    // --- CASE A: CATEGORIZED MULTI (e.g., Die Ejector) ---
    if (CATEGORIZED_MULTI_TOOLS.includes(parentID)) {
       // Fetch raw options but preserve structure
       var groups = [];
       var currentGroup = null;
       
       // Manually scan the menuData for this parentID
       for (var m = 0; m < menuData.length; m++) {
          if (String(menuData[m][0]) === parentID) {
             var line = String(menuData[m][1]).trim();
             
             if (line.startsWith("---")) {
                // New Group Detected (e.g., --- OPTIONAL EJECTOR NEEDLE [430001-A381] ---)
                var rawHeader = line.replace(/---/g, "").trim();
                
                // EXTRACT SUB-ID FROM HEADER IF PRESENT
                var groupName = rawHeader;
                var groupID = null;
                var match = rawHeader.match(/\[(.*?)\]/);
                if (match && match[1]) {
                    groupID = match[1];
                    groupName = rawHeader.replace(/\[.*?\]/, "").trim();
                }

                var groupType = (rawHeader.indexOf("OPTIONAL") > -1) ? "multi" : "single";
                
                currentGroup = {
                   name: groupName,
                   id: groupID, // Pass this to Frontend for Sub-Grouping
                   type: groupType,
                   options: []
                };
                groups.push(currentGroup);
                
             } else if (currentGroup && line) {
                // Add Item to Current Group
                var parts = line.split("::");
                var id = parts[0].trim();
                var desc = (parts.length > 1) ? parts[1].trim() : "";
                currentGroup.options.push({ id: id, desc: desc });
             }
          }
       }
       
       uniqueTools.push({
          id: parentID,
          desc: finalDesc,
          type: 'CATEGORIZED_MULTI',
          groups: groups
       });
    }
    
    // --- CASE B: COMPLEX RULES (e.g., PUH) ---
    else if (COMPLEX_TOOL_RULES[parentID]) {
       var mandatoryID = COMPLEX_TOOL_RULES[parentID];
       var allOptions = fetchOptionsForTool(menuData, parentID);
       
       var mandatoryItem = null;
       var selectableOptions = [];
       
       for (var k = 0; k < allOptions.length; k++) {
          if (allOptions[k].id === mandatoryID) {
             mandatoryItem = allOptions[k];
          } else {
             selectableOptions.push(allOptions[k]);
          }
       }
       
       uniqueTools.push({
          id: parentID,
          desc: finalDesc,
          type: 'COMPLEX', // Flag for UI
          mandatoryItem: mandatoryItem, // The Fixed Item
          options: selectableOptions // The Dropdown Items
       });
       
    } 
    // --- CASE C: MULTI-SELECT TOOLS ---
    else if (MULTI_SELECT_TOOLS.includes(parentID)) {
       var options = fetchOptionsForTool(menuData, parentID);
       uniqueTools.push({
          id: parentID,
          desc: finalDesc,
          type: 'MULTI_SELECT', // Flag for UI
          options: options
       });
    }
    // --- CASE D: STANDARD ---
    else {
       var options = fetchOptionsForTool(menuData, parentID);
       uniqueTools.push({
          id: parentID,
          desc: finalDesc,
          type: 'STANDARD',
          options: options
       });
    }
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
  var configHeaderFinder = sheet.getRange("A:B").createTextFinder("CONFIGURATION").matchEntireCell(true).findNext();
  
  if (!configHeaderFinder) {
    throw new Error("Critical Error: 'CONFIGURATION' header not found in Column A or B. Cannot safely locate list start.");
  }

  var startScanRow = configHeaderFinder.getRow();
  
  // Scan Column G (Turret Location) starting from the header row downwards
  var scanRange = sheet.getRange(startScanRow, 7, lastRow - startScanRow + 1, 2).getValues();

  var targetRowIndex = -1;
  var startScanning = false;

  for (var i = 0; i < scanRange.length; i++) {
    var slotID = String(scanRange[i][0]).trim(); // Col G
    var desc = String(scanRange[i][1]).trim();   // Col H

    // STRICT VALIDATION: Is this a "B" slot?
    var isBSlot = slotID.startsWith("B") && !isNaN(parseInt(slotID.substring(1)));

    if (!startScanning) {
      if (isBSlot) {
        startScanning = true; // Found the first valid slot (e.g., B10)
      }
    }

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

  // WRITE DATA
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
 * 3.5 SAVE MACHINE SETUP (Phase 6 - HYBRID RESTRUCTURE)
 * Implements "Form" for Top Sections, "Tree" for Tooling.
 * UPDATED: 2-Level Flattened Hierarchy (Col C & D only, Desc in E)
 * FIX: "Collapse & Rebuild" Strategy to prevent duplication bugs.
 * FIX: Sanitation to prevent empty row gaps.
 * FIX: Robust Anchor Detection (Regex: Comments?:?) and Order Safety Check.
 */
function saveMachineSetup(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Sheet 'TRIAL-LAYOUT CONFIGURATION' not found.");
  
  // 1. Save Vision PC (Merged Cell C2) - FORM STYLE
  if (payload.visionPC) {
    sheet.getRange(2, 3).setValue(payload.visionPC);
  }
  
  // 2. Save Configurable Base Modules (Merged Cells C:G) - FORM STYLE
  if (payload.baseModules && Array.isArray(payload.baseModules)) {
    var startFinder = sheet.getRange("A:B").createTextFinder("Configurable Base Module").matchEntireCell(false).findNext();
    var endFinder = sheet.getRange("A:B").createTextFinder("Base Module Tooling").matchEntireCell(false).findNext();

    if (startFinder && endFinder) {
      var startRow = startFinder.getRow();
      var nextSectionRow = endFinder.getRow();
      var currentGap = nextSectionRow - startRow; 
      var itemsToSave = payload.baseModules;
      var requiredSlots = Math.max(itemsToSave.length, 3); // Minimum 3 lines
      
      // Smart Resize (Safe here as usually small fixed list)
      if (requiredSlots > currentGap) {
         var rowsNeeded = requiredSlots - currentGap;
         sheet.insertRowsBefore(nextSectionRow, rowsNeeded);
         var newRowsRange = sheet.getRange(nextSectionRow, 1, rowsNeeded, sheet.getLastColumn());
         newRowsRange.setTextRotation(0).setVerticalAlignment("middle");
      } else if (requiredSlots < currentGap) {
         var rowsToDelete = currentGap - requiredSlots;
         if (rowsToDelete > 0) sheet.deleteRows(startRow + requiredSlots, rowsToDelete);
      }
      
      // Write & Merge Logic
      for (var i = 0; i < requiredSlots; i++) {
         var targetRow = startRow + i;
         var val = (i < itemsToSave.length) ? itemsToSave[i] : ""; 
         sheet.getRange(targetRow, 3).setValue(val);
         var mergeRange = sheet.getRange(targetRow, 3, 1, 5); // C to G
         try {
           if (!mergeRange.isPartOfMerge() || mergeRange.getMergedRanges().length > 1) { mergeRange.merge(); }
           mergeRange.setVerticalAlignment("middle");
         } catch(e) {}
      }
    }
  }

  // 3. Save Base Module Tooling (FLATTENED TREE - UNMERGED)
  // CRITICAL FIX: COLLAPSE & REBUILD STRATEGY WITH ROBUST ANCHORING
  if (payload.baseTooling && Array.isArray(payload.baseTooling)) {
    
    // --- SANITATION STEP: FILTER OUT EMPTY ROWS ---
    var sanitizedTooling = payload.baseTooling.filter(function(item) {
        var hasId = item.id && item.id.trim() !== "";
        var hasDesc = item.desc && item.desc.trim() !== "";
        var hasChildren = item.structure && item.structure.length > 0;
        return hasId || hasDesc || hasChildren;
    });

    var startFinder = sheet.getRange("A:B").createTextFinder("Base Module Tooling").matchEntireCell(false).findNext();
    
    // ROBUST ANCHORING: FIND ALL Variations (Comment, Comments, Comment:)
    // Use Regex: (?i)^Comments?:?$ -> Case insensitive, 'Comment' or 'Comments', optional ':'
    var commentFinders = sheet.getRange("A:B").createTextFinder("(?i)^Comments?:?$").useRegularExpression(true).matchEntireCell(true).findAll();
    var endFinder = null;
    
    if (commentFinders && commentFinders.length > 0) {
        // Pick the last one to be safe against matches in descriptions (though matchEntireCell protects somewhat)
        endFinder = commentFinders[commentFinders.length - 1]; 
    }
    
    if (!startFinder) {
       // If start anchor missing, we cannot proceed safely
       return { status: "error", message: "Anchor 'Base Module Tooling' not found." };
    }

    if (!endFinder) {
       throw new Error("Critical Error: Footer anchor 'Comment' (or Comment:) not found in Column A/B. Aborting save.");
    }

    if (startFinder && endFinder) {
      var startRow = startFinder.getRow();
      var endRow = endFinder.getRow();

      // SAFETY CHECK: Ensure End is actually below Start
      if (endRow <= startRow) {
         throw new Error("Layout Error: Found 'Comment' anchor at Row " + endRow + " which is above/equal to 'Base Module Tooling' at Row " + startRow + ". Please check your staging sheet anchors.");
      }
      
      // A. Build Flattened Write Queue (2-Level Logic)
      var writeQueue = [];
      
      for (var p = 0; p < sanitizedTooling.length; p++) {
         var tool = sanitizedTooling[p];
         
         // Level 1: Tool Parent -> Col C
         writeQueue.push({ 
             level: 1, 
             id: tool.id, 
             desc: tool.desc 
         });
         
         // Level 2: Flattened Children/Options
         if (tool.structure && tool.structure.length > 0) {
             for (var s = 0; s < tool.structure.length; s++) {
                 var item = tool.structure[s];
                 
                 // Strict check for option emptiness too
                 if (!item.id && !item.desc) continue;

                 if (item.type === 'option') {
                     // Standard Option -> Level 2
                     writeQueue.push({ level: 2, id: item.id, desc: item.desc });
                 } else if (item.type === 'group') {
                     // Group Header -> DISCARD (Option A)
                     // Process Children -> PROMOTE to Level 2 (Siblings)
                     if (item.children && item.children.length > 0) {
                         for (var c = 0; c < item.children.length; c++) {
                             var child = item.children[c];
                             writeQueue.push({ level: 2, id: child.id, desc: child.desc });
                         }
                     }
                 }
             }
         }
      }
      
      var requiredSlots = Math.max(writeQueue.length, 1);

      // B. NUCLEAR OPTION: COLLAPSE THEN EXPAND
      // 1. Calculate how many rows exist currently between headers
      var currentGap = endRow - startRow - 1; // Rows strictly between headers
      
      // 2. Delete ALL existing rows in the gap (if any)
      if (currentGap > 0) {
         sheet.deleteRows(startRow + 1, currentGap);
      }
      
      // 3. Insert EXACTLY needed rows (fresh canvas)
      if (requiredSlots > 0) {
         sheet.insertRowsAfter(startRow, requiredSlots);
         // Reset formatting for new rows just in case
         var newRange = sheet.getRange(startRow + 1, 1, requiredSlots, sheet.getLastColumn());
         newRange.setTextRotation(0).setVerticalAlignment("middle").setFontWeight("normal").setFontStyle("normal");
      }
      
      // C. WRITE DATA
      // Col C = Level 1 ID
      // Col D = Level 2 ID
      // Col E = Description (Moved from F)
      // Col F = Empty
      
      var outputRange = sheet.getRange(startRow + 1, 3, requiredSlots, 3); // C, D, E
      var values = [];
      var fontWeights = [];

      for (var k = 0; k < requiredSlots; k++) {
         if (k < writeQueue.length) {
            var data = writeQueue[k];
            var colC = (data.level === 1) ? data.id : "";
            var colD = (data.level === 2) ? data.id : "";
            var colE = data.desc || "";
            
            values.push([colC, colD, colE]);
            // All Bold as requested
            fontWeights.push(["bold", "bold", "bold"]);
         } else {
            // Should not happen with nuclear logic, but safe fallback
            values.push(["", "", ""]);
            fontWeights.push(["normal", "normal", "normal"]);
         }
      }
      
      outputRange.setValues(values);
      outputRange.setFontWeights(fontWeights);
      
      // Cleanup Col F just to be safe (explicit clear)
      sheet.getRange(startRow + 1, 6, requiredSlots, 1).clearContent();
      
      // Borders
      var block = sheet.getRange(startRow + 1, 3, requiredSlots, 4); // C to F
      block.setBorder(true, true, true, true, true, true, "lightgray", SpreadsheetApp.BorderStyle.SOLID);
    }
  }
  
  return { status: "success" };
}

// =======================================================
// EXTRACTION LOGIC
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
      addToDict(parts[0].trim(), parts[1].trim());
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
 * UPDATED: Includes Machine Setup (Top Section) Extraction
 */
function extractProductionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Staging Sheet Missing");

  var masterDict = buildMasterDictionary();
  var lastRow = sheet.getLastRow();

  // Initialize Payload
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

  // --- SECTION 1: MACHINE SETUP EXTRACTION (Cols C-E) ---
  // Scan between "Base Module Tooling" and "Comment"
  try {
     var toolStart = sheet.getRange("A:B").createTextFinder("Base Module Tooling").matchEntireCell(false).findNext();
     // FIX: Robust Regex Search for Comment Anchor
     var toolEndFinder = sheet.getRange("A:B").createTextFinder("(?i)^Comments?:?$").useRegularExpression(true).matchEntireCell(true).findAll();
     
     if (toolStart && toolEndFinder && toolEndFinder.length > 0) {
        var startR = toolStart.getRow(); // The header row "Base Module Tooling"
        var endR = toolEndFinder[toolEndFinder.length - 1].getRow(); // Last "Comment:"
        
        // Data is between startR and endR
        if (endR > startR + 1) {
           var setupRange = sheet.getRange(startR + 1, 3, endR - startR - 1, 3).getValues(); // C, D, E
           
           for (var m = 0; m < setupRange.length; m++) {
              var rowData = setupRange[m];
              var idC = String(rowData[0]).trim(); // Parent (Level 1)
              var idD = String(rowData[1]).trim(); // Child (Level 2)
              var descE = String(rowData[2]).trim(); // Description in E
              
              var validID = "";
              if (idC) validID = idC;
              else if (idD) validID = idD;
              
              if (validID) {
                 payload.TOOLING.push({
                    requiresNumbering: true,
                    id: validID,
                    desc: descE || masterDict[validID] || "Manual Entry",
                    qty: 1
                 });
              }
           }
        }
     }
  } catch(e) {
     console.warn("Machine Setup Extraction Failed: " + e.message);
  }
  
  // --- SECTION 2: MODULE CONFIGURATION EXTRACTION (B-Slots) ---
  // Dynamic Header Reading
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

  var startScanning = false;
  
  for (var i = 0; i < rawData.length; i++) {
    var row = rawData[i];
    var turretName = String(row[6]).trim(); // Col G
    var syncStatus = String(row[9]).trim(); // Col J
    var moduleID = String(row[10]).trim();  // Col K
    
    // Dynamic start check for B10+
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

      // Tooling - Standard extraction from Cols M & N
      pushItem("TOOLING", String(row[13]), null, true); 
      pushItem("TOOLING", String(row[14]), null, false); 

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

/**
 * C. Inject Data to Production
 */
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
 * D. The Smart Fill Helper
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

  // 6. Write Data
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
  sheet.getRange(startWriteRow, 7, itemsNeeded, 1).insertCheckboxes();

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
