/**
 * ConfiguratorBackend.gs
 * Server-side logic for the Module Configurator UI.
 * Handles fetching data from REF_DATA and calculating logic (Rotation, Rubber Tips, etc.)
 */

// --- CONSTANTS ---
// Tools that trigger the Rubber Tip Dependency logic
var RUBBER_TIP_PARENTS_BACKEND = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
// The Source ID for Rubber Tip Options
var RUBBER_TIP_SOURCE_ID_BACKEND = "430001-A380";

/**
 * 1. Get List of Modules for Sidebar Dropdown
 * Reads REF_DATA Column C (Index 3).
 */
function getModuleList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];
  
  // Column C is Index 3. 
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  var rawValues = sheet.getRange(1, 3, lastRow, 1).getValues();
  var modules = [];
  
  for (var i = 0; i < rawValues.length; i++) {
    var val = String(rawValues[i][0]).trim();
    if (val !== "" && val !== "Part ID") { // Basic cleanup
      modules.push(val);
    }
  }
  return modules;
}

/**
 * 2. Get Full Details for Selected Module
 * Calculates the "Virtual BOM" including:
 * - Electrical Rotation (Simulated as Instance 1 for now)
 * - Tooling Stacks
 * - Tooling Options (Standard)
 * - RUBBER TIP DEPENDENCY (Special)
 * - Vision (List)
 */
function getModuleDetails(moduleID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName("REF_DATA");
  
  // 1. Find the Module Row in REF_DATA (Cols C:AD)
  // C=3 ... AD=30
  var dataRange = refSheet.getRange("C:AD").getValues();
  var rowData = null;
  
  for (var i = 0; i < dataRange.length; i++) {
    if (String(dataRange[i][0]) === moduleID) {
      rowData = dataRange[i];
      break;
    }
  }
  
  if (!rowData) {
    return { error: "Module ID not found in REF_DATA." };
  }
  
  // Extract Raw Strings (Based on Index in C:AD, where C is index 0)
  // W=20, X=21, Y=22, Z=23, AA=24, AB=25, AC=26, AD=27
  var elecIdsStr = rowData[20];
  var elecDescStr = rowData[21];
  var toolIdsStr = rowData[22];
  var toolDescStr = rowData[23];
  var jigIdsStr = rowData[24];
  var jigDescStr = rowData[25];
  var visionIdsStr = rowData[26];
  
  // --- A. Process Electrical (Rotation Logic) ---
  // READ-ONLY PHASE: We default to Instance 1 (Index 0)
  var elecResult = null;
  if (elecIdsStr) {
    var eIds = String(elecIdsStr).split(';').map(function(s){ return s.trim(); });
    var eDescs = String(elecDescStr).split(';').map(function(s){ return s.trim(); });
    
    // Logic: (InstanceCount 0) % Length = Index 0
    if (eIds.length > 0) {
      elecResult = {
        id: eIds[0],
        desc: eDescs[0] || "",
        note: "Instance 1 (Rotation A)" // Visual indicator for UI
      };
    }
  }
  
  // --- B. Process Tooling (Complex Logic) ---
  var toolsResult = [];
  if (toolIdsStr) {
    var tIds = String(toolIdsStr).split(';').map(function(s){ return s.trim(); });
    var tDescs = String(toolDescStr).split(';').map(function(s){ return s.trim(); });
    
    // Fetch Shadow Menu Data (U:V) for Options Lookup
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
      
      // 1. Check Standard Options (Direct Children)
      var stdOpts = fetchOptionsForTool(menuData, tID);
      if (stdOpts.length > 0) {
        toolObj.standardOptions = stdOpts;
      }
      
      // 2. CHECK RUBBER TIP DEPENDENCY
      if (RUBBER_TIP_PARENTS_BACKEND.includes(tID)) {
        // Fetch options for the Master Tip Source (430001-A380)
        var tipOpts = fetchOptionsForTool(menuData, RUBBER_TIP_SOURCE_ID_BACKEND);
        if (tipOpts.length > 0) {
          toolObj.requiresRubberTip = true;
          toolObj.rubberTipOptions = tipOpts;
        }
      }
      
      toolsResult.push(toolObj);
    }
  }
  
  // --- C. Process Vision ---
  var visionResult = { type: 'none', options: [] };
  if (visionIdsStr) {
    var vIds = String(visionIdsStr).split(';').map(function(s){ return s.trim(); });
    vIds = vIds.filter(function(s) { return s.length > 0; });
    
    if (vIds.length === 1) {
      visionResult.type = 'fixed';
      visionResult.options = vIds; // Just one
    } else if (vIds.length > 1) {
      visionResult.type = 'select';
      visionResult.options = vIds;
    }
  }
  
  // --- D. Process Jigs ---
  var jigsResult = [];
  if (jigIdsStr) {
    var jIds = String(jigIdsStr).split(';').map(function(s){ return s.trim(); });
    var jDescs = String(jigDescStr).split(';').map(function(s){ return s.trim(); });
    for (var j = 0; j < jIds.length; j++) {
      if(jIds[j]) jigsResult.push({ id: jIds[j], desc: jDescs[j] || "" });
    }
  }

  // Construct Final JSON
  return {
    moduleID: moduleID,
    electrical: elecResult,
    tools: toolsResult,
    jigs: jigsResult,
    vision: visionResult
  };
}

/**
 * Helper: Read Shadow Menu (Cols U:V) to find options for a Parent Tool
 * Returns array of strings: ["OptionID1", "OptionID2"]
 */
function fetchOptionsForTool(menuData, parentID) {
  var options = [];
  var foundParent = false;
  
  for (var i = 0; i < menuData.length; i++) {
    var rowParent = String(menuData[i][0]);
    var rowChild = String(menuData[i][1]);
    
    if (rowParent === parentID) {
      foundParent = true;
      // Filter out Headers/separators (e.g., "--- ... ---")
      if (rowChild && rowChild.indexOf("---") === -1) {
        options.push(rowChild);
      }
    } else if (foundParent) {
      // Since menuData is grouped, if we were finding it and now the parent changed (and isn't blank), stop.
      // (Assuming blank parent means continuation of previous block, or data is strictly sorted)
      if (rowParent !== "") break; 
    }
  }
  return options;
}
