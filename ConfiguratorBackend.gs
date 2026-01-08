/**
 * ConfiguratorBackend.gs
 * Server-side logic for the Module Configurator UI.
 */

// --- CONSTANTS ---
var RUBBER_TIP_PARENTS_BACKEND = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
var RUBBER_TIP_SOURCE_ID_BACKEND = "430001-A380";

/**
 * 1. Get List of Modules (ID and Description)
 * Reads REF_DATA Cols C (ID) and D (Description).
 */
function getModuleList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  // Read Cols C and D (Index 3 and 4 in 1-based notation? No, getRange uses 1-based)
  // Col C = 3, Col D = 4.
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
 * 2. Get Full Details for Selected Module
 * Fetches IDs and Descriptions for all components.
 */
function getModuleDetails(moduleID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName("REF_DATA");
  
  // Fetch C:AD range
  var dataRange = refSheet.getRange("C:AD").getValues();
  var rowData = null;
  
  for (var i = 0; i < dataRange.length; i++) {
    if (String(dataRange[i][0]) === moduleID) {
      rowData = dataRange[i];
      break;
    }
  }
  
  if (!rowData) return { error: "Module ID not found." };
  
  // Array Indices (0-based from C):
  // W=20 (Elec ID), X=21 (Elec Desc)
  // Y=22 (Tool ID), Z=23 (Tool Desc)
  // AA=24 (Jig ID), AB=25 (Jig Desc)
  // AC=26 (Vision ID), AD=27 (Vision Desc)
  
  // --- ELECTRICAL ---
  var elecResult = null;
  var elecIdsStr = rowData[20];
  var elecDescStr = rowData[21];
  
  if (elecIdsStr) {
    var eIds = String(elecIdsStr).split(';').map(function(s){ return s.trim(); });
    var eDescs = String(elecDescStr).split(';').map(function(s){ return s.trim(); });
    
    // Default to Instance 1
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
      
      // Standard Options
      var stdOpts = fetchOptionsForTool(menuData, tID);
      if (stdOpts.length > 0) toolObj.standardOptions = stdOpts;
      
      // Rubber Tip Dependency
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
    
    // Create objects for each vision option
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

  return {
    moduleID: moduleID,
    electrical: elecResult,
    tools: toolsResult,
    jigs: jigsResult,
    vision: visionResult
  };
}

function fetchOptionsForTool(menuData, parentID) {
  var options = [];
  var foundParent = false;

  for (var i = 0; i < menuData.length; i++) {
    var rowParent = String(menuData[i][0]);
    var rowChild = String(menuData[i][1]); // This now contains "ID :: Desc"

    if (rowParent === parentID) {
      foundParent = true;
      if (rowChild && rowChild.indexOf("---") === -1) {
        
        // UNPACK LOGIC
        // We look for the delimiter "::"
        var parts = rowChild.split('::');
        var id = parts[0].trim();
        var desc = "";
        
        if (parts.length > 1) {
           desc = parts[1].trim();
        }
        
        // Push an OBJECT, not just a string
        options.push({ id: id, desc: desc }); 
      }
    } else if (foundParent) {
      // If we were finding the parent but now rowParent is empty/different (and not just an empty cell in a contiguous block), stop.
      // Logic from before: "if (rowParent !== "") break;" means if we hit a NEW parent, stop.
      if (rowParent !== "") break; 
    }
  }
  return options;
}
