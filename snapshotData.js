///
// Creates a snapshot of the data to be analyzed over time.
///

// Creates a snapshot of sheet and appends to log
function snapshotStoryTracking() {
  // get current sheet and tabs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var current = ss.getSheetByName("Pivot::StoryTracking");
  var database = ss.getSheetByName("StoryTracking");

  // count rows to snap
  var currentRows = current.getLastRow();
  var databaseRows = database.getLastRow() + 1;
  var databaseRowNew = currentRows + databaseRows - 3; // change number dependant on the depth
  var rowNew = current.getRange("A3:E" + currentRows).getValues();

  // snap rows, can run this on a trigger to be timed
  database
    .getRange("A" + databaseRows + ":E" + databaseRowNew)
    .setValues(rowNew);
}

// Creates a snapshot of sheet and appends to log
function snapshotMemberTracking() {
  // get current sheet and tabs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var current = ss.getSheetByName("Pivot::MemberTracking");
  var database = ss.getSheetByName("MemberTracking");

  // count rows to snap
  var currentRows = current.getLastRow();
  var databaseRows = database.getLastRow() + 1;
  var databaseRowNew = currentRows + databaseRows - 4;
  var rowNew = current.getRange("A4:F" + currentRows).getValues();

  // snap rows, can run this on a trigger to be timed
  database
    .getRange("A" + databaseRows + ":F" + databaseRowNew)
    .setValues(rowNew);
}

function snapshotEpicTracking() {
  // get current sheet and tabs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var current = ss.getSheetByName("Pivot::EpicTracking");
  var database = ss.getSheetByName("EpicTracking");

  // count rows to snap
  var currentRows = current.getLastRow();
  var databaseRows = database.getLastRow() + 1;
  var databaseRowNew = currentRows + databaseRows - 3; // change number dependant on the depth
  var rowNew = current.getRange("A3:E" + currentRows).getValues();

  // snap rows, can run this on a trigger to be timed
  database
    .getRange("A" + databaseRows + ":E" + databaseRowNew)
    .setValues(rowNew);
}

// Single function to run all snapshots
function snapshotData() {
  snapshotStoryTracking();
  snapshotMemberTracking();
  snapshotEpicTracking();
}
