///
// Runs a find and replace to replace the ID within the sheet with the human readable label.
///

// Find and Replace
function replaceInSheet(sheet, to_replace, replace_with, column) {
  //get the current data range values as an array
  var values = sheet.getDataRange().getValues();

  //loop over the rows in the array
  for (var row in values) {
    var replaceValueKey = values[row][column].toString();
    //use Array.map to execute a replace call on each of the cells in the row.
    var replaced_values = values[row].map(function(original_value) {
      var original_value = original_value.toString();
      if (original_value == replaceValueKey) {
        return original_value.replace(to_replace, replace_with);
      } else {
        return original_value;
      }
    });

    // Replace the original row values with the replaced values
    values[row] = replaced_values;
  }

  //write the updated values to the sheet
  sheet.getDataRange().setValues(values);
}

///
// Replace Story Sheet
///

// Replace Member Id with Name
function replaceMemberId(sheetMembers, sheetStories) {
  // Members
  var columnMemberIdIndex = 2; //column Index
  var columnMemberNameIndex = 3; //column Index
  var columnValuesMemberName = sheetMembers
    .getRange(2, columnMemberIdIndex, sheetMembers.getLastRow())
    .getValues();
  var columnValuesMemberID = sheetMembers
    .getRange(2, columnMemberNameIndex, sheetMembers.getLastRow())
    .getValues();
  var numValuesMembers = columnValuesMemberID.length;

  for (var i = 0; i < numValuesMembers; i++) {
    // Replace Id with Names
    var id = columnValuesMemberID[i].toString();
    var name = columnValuesMemberName[i].toString();
    replaceInSheet(sheetStories, name, id, 8);
  }

  for (var i = 0; i < numValuesMembers; i++) {
    // Replace Id with Names
    var id = columnValuesMemberID[i].toString();
    var name = columnValuesMemberName[i].toString();
    replaceInSheet(sheetStories, name, id, 9);
  }
}

// Replace Workflow Id with Name
function replaceWorkflowId(sheetWorkflow, sheetStories) {
  // Workflow
  var columnWorkflowIdIndex = 4; //column Index
  var columnWorkflowNameIndex = 5; //column Index
  var columnValuesWorkflowName = sheetWorkflow
    .getRange(2, columnWorkflowIdIndex, sheetWorkflow.getLastRow())
    .getValues();
  var columnValuesWorkflowID = sheetWorkflow
    .getRange(2, columnWorkflowNameIndex, sheetWorkflow.getLastRow())
    .getValues();
  var numValuesWorkflow = columnValuesWorkflowID.length;
  for (var i = 0; i < numValuesWorkflow; i++) {
    // Replace Id with Names
    var id = columnValuesWorkflowID[i].toString();
    var name = columnValuesWorkflowName[i].toString();
    replaceInSheet(sheetStories, name, id, 5);
  }
}

// Replace Projects Id with Name
function replaceProjectsId(sheetProjects, sheetStories) {
  // Projects
  var columnProjectIdIndex = 2; //column Index
  var columnProjectNameIndex = 3; //column Index
  var columnValuesProjectName = sheetProjects
    .getRange(2, columnProjectIdIndex, sheetProjects.getLastRow())
    .getValues();
  var columnValuesProjectID = sheetProjects
    .getRange(2, columnProjectNameIndex, sheetProjects.getLastRow())
    .getValues();
  var numValuesProjects = columnValuesProjectID.length;
  for (var i = 0; i < numValuesProjects; i++) {
    // Replace Id with Names
    var id = columnValuesProjectID[i].toString();
    var name = columnValuesProjectName[i].toString();
    replaceInSheet(sheetStories, name, id, 4);
  }
}

// Replace Epics Id with Name
function replaceEpicsId(sheetEpics, sheetStories) {
  // Epics
  var columnEpicIdIndex = 2; //column Index
  var columnEpicNameIndex = 3; //column Index
  var columnValuesEpicName = sheetEpics
    .getRange(2, columnEpicIdIndex, sheetEpics.getLastRow())
    .getValues();
  var columnValuesEpicID = sheetEpics
    .getRange(2, columnEpicNameIndex, sheetEpics.getLastRow())
    .getValues();
  var numValuesEpics = columnValuesEpicID.length;
  for (var i = 0; i < numValuesEpics; i++) {
    // Replace Id with Names
    var id = columnValuesEpicID[i].toString();
    var name = columnValuesEpicName[i].toString();
    replaceInSheet(sheetStories, name, id, 3);
  }
}

// Main stories call
// Split into 2 functions to avoid timeouts in Google Scripts
function runReplaceInSheetStories() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetStories = ss.getSheetByName("Stories");
  var sheetMembers = ss.getSheetByName("Members");
  var sheetProjects = ss.getSheetByName("Projects");
  var sheetEpics = ss.getSheetByName("Epics");
  var sheetWorkflow = ss.getSheetByName("Workflows");

  replaceWorkflowId(sheetWorkflow, sheetStories);
  replaceMemberId(sheetMembers, sheetStories);
  //replaceProjectsId(sheetProjects, sheetStories);
  //replaceEpicsId(sheetEpics, sheetStories);
}

function runReplaceInSheetStories1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetStories = ss.getSheetByName("Stories");
  var sheetMembers = ss.getSheetByName("Members");
  var sheetProjects = ss.getSheetByName("Projects");
  var sheetEpics = ss.getSheetByName("Epics");
  var sheetWorkflow = ss.getSheetByName("Workflows");

  //replaceWorkflowId(sheetWorkflow, sheetStories);
  //replaceMemberId(sheetMembers, sheetStories);
  replaceProjectsId(sheetProjects, sheetStories);
  replaceEpicsId(sheetEpics, sheetStories);
}

///
// Replace Epic Sheet
///

function replaceEpicMilestones(sheetMilestones, sheetEpics) {
  // Epics
  var columnIdIndex = 2; // column Index
  var columnNameIndex = 3; // column Index
  var columnValuesName = sheetMilestones
    .getRange(2, columnIdIndex, sheetMilestones.getLastRow())
    .getValues();
  var columnValuesID = sheetMilestones
    .getRange(2, columnNameIndex, sheetMilestones.getLastRow())
    .getValues();
  var numValues = columnValuesID.length;
  for (var i = 0; i < numValues; i++) {
    // Replace Id with Names
    var id = columnValuesID[i].toString();
    var name = columnValuesName[i].toString();

    replaceInSheet(sheetEpics, name, id, 11);
  }
}

function replaceEpicWorkflow(sheetWorkflowEpics, sheetEpics) {
  var columnIdIndex = 2; // column Index
  var columnNameIndex = 3; // column Index
  var columnValuesName = sheetWorkflowEpics
    .getRange(2, columnIdIndex, sheetWorkflowEpics.getLastRow())
    .getValues();
  var columnValuesID = sheetWorkflowEpics
    .getRange(2, columnNameIndex, sheetWorkflowEpics.getLastRow())
    .getValues();
  var numValues = columnValuesID.length;
  for (var i = 0; i < numValues; i++) {
    // Replace Id with Names
    var id = columnValuesID[i].toString();
    var name = columnValuesName[i].toString();
    replaceInSheet(sheetEpics, name, id, 13);
  }
}

function runReplaceInSheetEpics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMilestones = ss.getSheetByName("Milestones");
  var sheetWorkflowEpics = ss.getSheetByName("WorkflowEpics");
  var sheetEpics = ss.getSheetByName("Epics");

  replaceEpicMilestones(sheetMilestones, sheetEpics);
  replaceEpicWorkflow(sheetWorkflowEpics, sheetEpics);
}
