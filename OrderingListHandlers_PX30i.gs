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



        ui.alert("Success: Item unmarked.");
      } else {
        ui.alert("Error: Incorrect Password.");
      }
    }
  }
}
