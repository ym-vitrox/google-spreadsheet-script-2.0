/**
 * OrderingListHandlers.gs
 * Logic for Release Checkbox, Password Protection, and Timestamps.
 * Phase 5 Requirement
 */

function handleCheckboxEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var newVal = e.value;
  var oldVal = e.oldValue;

  // Security Const
  var secretPassword = "123";

  // Target Columns (Indexes)
  var COL_CHECKBOX = 7; // G
  var COL_DATE = 8;     // H
  var COL_TYPE = 9;     // I

  // 1. CASE: Checked (FALSE -> TRUE)
  if (newVal === "TRUE") {
    // Set Timestamp
    var dateCell = sheet.getRange(row, COL_DATE);

    // Format explicitly to d/m/yyyy (Day/Month/Year) to match Malaysia locale reqs
    dateCell.setNumberFormat("d/m/yyyy");
    dateCell.setValue(new Date()); // Writes current date object

    // Reset Release Type to Empty (Requirement: Default Empty)
    sheet.getRange(row, COL_TYPE).setValue("");

    // Lock the row now that it is released
    refreshUnprotectedRanges(sheet);
  }

  // 2. CASE: Unchecked (TRUE -> FALSE)
  else if (newVal === "FALSE") {
    // Re-lock the cell immediately to prevent unauth change
    range.setValue(true);
    SpreadsheetApp.flush(); // Force UI update

    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Security Check',
      'Please enter the password to uncheck "Release":',
      ui.ButtonSet.OK_CANCEL
    );

    // Process Password
    if (response.getSelectedButton() == ui.Button.OK) {
      if (response.getResponseText() == secretPassword) {
        // Correct Password:
        // 1. Uncheck box
        range.setValue(false);
        // 2. Clear Date
        sheet.getRange(row, COL_DATE).clearContent();
        // 3. Clear Release Type
        sheet.getRange(row, COL_TYPE).clearContent();

        // Unlock the row by refreshing protection exceptions
        refreshUnprotectedRanges(sheet);

        ui.alert("Success: Item unmarked.");
      } else {
        ui.alert("Error: Incorrect Password.");
      }
    }
  }
}

// =========================================
// ROW LOCK MANAGER
// Manages sheet-level protection for released rows.
// ONE protection object covers the entire sheet.
// Released rows are removed from unprotectedRanges (locked).
// Unreleased rows remain in unprotectedRanges (freely editable).
// Col G is ALWAYS in unprotectedRanges so it stays clickable.
// =========================================

var ROW_LOCK_PROTECTION_DESC = "ROW_LOCK_SYSTEM";

/**
 * Gets the existing ROW_LOCK_SYSTEM sheet protection, or creates it.
 * @param {Sheet} sheet - The ORDERING LIST sheet.
 * @returns {Protection} The protection object.
 */
function getOrCreateSheetProtection(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === ROW_LOCK_PROTECTION_DESC) {
      var existing = protections[i];
      // Enforce full protection on every access (not just creation)
      // This also fixes any already-created protection that had stale editor bypasses
      existing.setWarningOnly(false);
      var existingEditors = existing.getEditors();
      if (existingEditors.length > 0) {
        existing.removeEditors(existingEditors);
      }
      return existing;
    }
  }
  // None found — create it for the first time
  var protection = sheet.protect().setDescription(ROW_LOCK_PROTECTION_DESC);
  // CRITICAL: Remove all editors from the bypass list so unprotectedRanges is the sole gatekeeper.
  // Without this, all spreadsheet editors bypass the protection silently.
  protection.setWarningOnly(false);
  var editors = protection.getEditors();
  if (editors.length > 0) {
    protection.removeEditors(editors);
  }
  return protection;
}

/**
 * Recomputes and applies the unprotectedRanges on the sheet protection.
 * Rules:
 *   - Any row where Col G = TRUE (released) is LOCKED (not in exceptions).
 *   - Any row where Col G = FALSE or blank (unreleased/empty) is FREE.
 *   - Col G of ALL rows is always FREE so the checkbox stays clickable.
 * @param {Sheet} sheet - The ORDERING LIST sheet.
 */
function refreshUnprotectedRanges(sheet) {
  var protection = getOrCreateSheetProtection(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  // Read Col G (column 7) for all rows in one call
  var colGValues = sheet.getRange(1, 7, lastRow, 1).getValues();

  var unprotectedRanges = [];

  for (var i = 0; i < lastRow; i++) {
    var rowNum = i + 1; // 1-based
    var isReleased = colGValues[i][0] === true;

    if (!isReleased) {
      // Entire row is editable
      unprotectedRanges.push(sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()));
    } else {
      // Row is locked — but keep Col G clickable so the user can trigger unlock
      unprotectedRanges.push(sheet.getRange(rowNum, 7, 1, 1));
    }
  }

  // Apply the computed exception list
  if (unprotectedRanges.length > 0) {
    protection.setUnprotectedRanges(unprotectedRanges);
  }
}
