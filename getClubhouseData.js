///
// Base Data Fetches for the Google Sheet.
///

// Clubhouse Token Config
function apiTokenConfig() {
  return "0000000000000"; // Add your Clubhouse API token here!
}

// Clear Spreadsheet of all data
function clearSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }
}

// Builds timestamp for fetches
function buildDate() {
  var date = new Date(); // create new date for timestamp
  date.setHours(0);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);

  return date;
}

function getClubhouseData(url) {
  var response = UrlFetchApp.fetch(url); // get api endpoint
  var json = response.getContentText(); // get the response content as text
  var base = JSON.parse(json); // parse text into json

  return base;
}

// Fetches labels
function getClubhouseLabels(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Labels"); //The name of the sheet tab where you are sending the info

  clearSheet(sheet);

  var url = "https://api.clubhouse.io/api/v2/labels?token=" + apiToken;
  var base = getClubhouseData(url);

  var date = buildDate();

  var entityLength = base.length;
  for (var i = 0; i < entityLength; i++) {
    // Skip Archived Labels
    if (base[i].archived == false) {
      var stats = []; //create empty array to hold data points
      stats.push(date); // timestamp
      stats.push(base[i].id);
      stats.push(base[i].name);
      // stats.push(base[i].archived);
      stats.push(base[i].stats.num_stories_total);
      stats.push(base[i].stats.num_stories_in_progress);
      stats.push(base[i].stats.num_stories_completed);
      stats.push(base[i].stats.num_points_total);
      stats.push(base[i].stats.num_points_in_progress);
      stats.push(base[i].stats.num_points_completed);
      stats.push(base[i].updated_at);

      //append the stats array to the active sheet
      sheet.appendRow(stats);
    }
  }
}

// Fetches Projects
function getClubhouseProjects(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Projects"); //The name of the sheet tab where you are sending the info
  var range = ss
    .getSheetByName("Gantt::Epics")
    .getRange("A:D")
    .clearContent();

  clearSheet(sheet);

  // Request
  var url = "https://api.clubhouse.io/api/beta/projects?token=" + apiToken;
  var base = getClubhouseData(url);
  var date = buildDate();

  var entityLength = base.length;
  for (var i = 0; i < entityLength; i++) {
    // Skip Archived Projects
    if (base[i].archived == false) {
      var stats = []; //create empty array to hold data points
      stats.push(date); // timestamp
      stats.push(base[i].id);
      stats.push(base[i].name);
      stats.push(base[i].stats.num_stories);
      stats.push(base[i].stats.num_points);
      //append the stats array to the active sheet
      sheet.appendRow(stats);
    }
  }
}

// Fetches Epics
function getClubhouseEpics(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Epics"); // The name of the sheet tab where you are sending the info
  var range = ss
    .getSheetByName("Gantt::Epics")
    .getRange("A:D")
    .clearContent();

  var url = "https://api.clubhouse.io/api/v2/epics?token=" + apiToken;
  clearSheet(sheet);
  var base = getClubhouseData(url);

  var date = buildDate();

  var entityLength = base.length;
  for (var i = 0; i < entityLength; i++) {
    // Skip Archived Projects
    if (base[i].archived == false) {
      var stats = []; //create empty array to hold data points
      stats.push(date); // timestamp
      stats.push(base[i].id);
      stats.push(base[i].name);
      stats.push(base[i].completed);
      stats.push(base[i].stats.num_points);
      stats.push(base[i].stats.num_points_done);
      stats.push(base[i].stats.num_points_started);
      stats.push(base[i].stats.num_points_unstarted);
      stats.push(base[i].stats.num_stories_done);
      stats.push(base[i].stats.num_stories_started);
      stats.push(base[i].stats.num_stories_unstarted);
      stats.push(base[i].milestone_id);
      stats.push(base[i].completed_at);
      stats.push(base[i].epic_state_id);
      stats.push(base[i].started_at);
      stats.push(base[i].deadline);
      stats.push(base[i].state);

      // TODO :: Update this to grab all labels not just the first one
      if (base[i].labels.length && base[i].labels[0].archived == false) {
        stats.push(base[i].labels[0].id);
        stats.push(base[i].labels[0].name);
      } else {
        stats.push("");
        stats.push("");
      }

      // append the stats array to the active sheet
      sheet.appendRow(stats);

      if (base[i].started_at != null && base[i] != null) {
        var values = range.getValues();
        var start = new Date(base[i].started_at);

        if (base[i].deadline != null) {
          var deadline = new Date(base[i].deadline);
        } else if (base[i].completed_at != null) {
          var deadline = new Date(base[i].completed_at);
        } else {
          var deadline = "";
          var diffDays = "";
        }

        values[i] = [date, base[i].name, start, deadline];
        range.setValues(values);
      }
    }
  }
}

// Fetches Workflows
function getClubhouseWorkflowEpics(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("WorkflowEpics"); //The name of the sheet tab where you are sending the info

  clearSheet(sheet); // Clear

  var url = "https://api.clubhouse.io/api/v2/epic-workflow?token=" + apiToken;
  var response = UrlFetchApp.fetch(url); // get api endpoint
  var json = response.getContentText(); // get the response content as text
  var base = JSON.parse(json); //parse text into json

  var date = buildDate();
  Logger.log(base.epic_states);
  var entityLength = base.epic_states.length;

  for (var i = 0; i < entityLength; i++) {
    var workflow = base.epic_states;
    var stats = []; //create empty array to hold data points
    stats.push(date); // timestamp
    stats.push(workflow[i].id);
    stats.push(workflow[i].name);
    stats.push(workflow[i].type);
    //append the stats array to the active sheet
    sheet.appendRow(stats);
  }
}

// Fetches Members
function getClubhouseMembers(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Members"); // The name of the sheet tab where you are sending the info

  clearSheet(sheet);

  var url = "https://api.clubhouse.io/api/v2/members?token=" + apiToken;
  var base = getClubhouseData(url);

  var date = buildDate();

  var entityLength = base.length;
  for (var i = 0; i < entityLength; i++) {
    // Skip deactivated profiles
    if (base[i].profile.deactivated == false) {
      var stats = []; // create empty array to hold data points
      stats.push(date); // timestamp
      stats.push(base[i].id);
      stats.push(base[i].profile.name);
      stats.push(base[i].profile.email_address);
      //append the stats array to the active sheet
      sheet.appendRow(stats);
    }
  }
}

// Fetches Workflows
function getClubhouseWorkflows(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Workflows"); //The name of the sheet tab where you are sending the info

  clearSheet(sheet); // Clear

  var url = "https://api.clubhouse.io/api/v2/workflows?token=" + apiToken;
  var response = UrlFetchApp.fetch(url); // get api endpoint
  var json = response.getContentText(); // get the response content as text
  var base = JSON.parse(json); //parse text into json

  var date = buildDate();

  var entityLength = base[0].states.length;

  for (var i = 0; i < entityLength; i++) {
    var workflow = base[0];
    var stats = []; //create empty array to hold data points
    stats.push(date); // timestamp
    stats.push(workflow.id);
    stats.push(workflow.name);
    stats.push(workflow.states[i].id);
    stats.push(workflow.states[i].name);
    stats.push(workflow.states[i].num_stories);
    stats.push(workflow.states[i].position);

    //append the stats array to the active sheet
    sheet.appendRow(stats);
  }
}

// Fetches Milestones
function getClubhouseMilestones(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Milestones"); //The name of the sheet tab where you are sending the info

  clearSheet(sheet);

  // Request
  var url = "https://api.clubhouse.io/api/beta/milestones?token=" + apiToken;
  var base = getClubhouseData(url);
  var date = buildDate();

  var entityLength = base.length;
  for (var i = 0; i < entityLength; i++) {
    var stats = []; //create empty array to hold data points
    stats.push(date); // timestamp
    stats.push(base[i].id);
    stats.push(base[i].name);
    stats.push(base[i].started);
    stats.push(base[i].started_at);
    stats.push(base[i].updated_at);
    stats.push(base[i].completed);
    stats.push(base[i].completed_at);
    stats.push(base[i].position);
    stats.push(base[i].state);

    // TODO :: Update this to grab all catagories not just the first one
    if (base[i].categories.length && base[i].categories[0].archived == false) {
      stats.push(base[i].categories[0].id);
      stats.push(base[i].categories[0].name);
    } else {
      stats.push("");
      stats.push("");
    }

    //append the stats array to the active sheet
    sheet.appendRow(stats);
  }
}

// Loop for appending new stories to stories sheet
function getClubhouseProjectStories(
  projects,
  start,
  projectAmount,
  sheet,
  apiToken
) {
  for (var project = start; project <= projectAmount; project++) {
    var projectIdRange = projects.getRange(project, 2);
    var projectId = projectIdRange.getValue();

    var url =
      "https://api.clubhouse.io/api/v2/projects/" +
      projectId +
      "/stories?token=" +
      apiToken;
    var base = getClubhouseData(url);

    var date = buildDate();

    var entityLength = base.length;
    for (var i = 0; i < entityLength; i++) {
      // Skip Archived Projects
      if (base[i].archived == false) {
        var stats = []; // create empty array to hold data points
        stats.push(date); // timestamp
        stats.push(base[i].id);
        stats.push(base[i].story_type);
        stats.push(base[i].epic_id);
        stats.push(base[i].project_id);
        stats.push(base[i].workflow_state_id);
        stats.push(base[i].position);
        stats.push(base[i].estimate);

        // TODO :: Update this to grab all owners not just the first one
        if (base[i].owner_ids.length) {
          stats.push(base[i].owner_ids[0]);
        } else {
          stats.push("");
        }
        stats.push(base[i].requested_by_id);

        // TODO :: Update this to grab all labels not just the first one
        if (base[i].labels.length && base[i].labels[0].archived == false) {
          stats.push(base[i].labels[0].id);
          stats.push(base[i].labels[0].name);
        } else {
          stats.push("");
          stats.push("");
        }
        stats.push(base[i].updated_at);
        stats.push(base[i].name);
        stats.push(base[i].started_at);
        stats.push(base[i].created_at);
        stats.push(base[i].app_url);
        stats.push(base[i].blocked);
        stats.push(base[i].blocker);
        stats.push(base[i].deadline);

        //append the stats array to the active sheet
        sheet.appendRow(stats);
      }
    }
    Logger.log("Completed Story Pull - " + projectId);
  }
}

///
// Trigger Functions -- Setup with the script UI.
///

// NOTE:: Split into 3 functions (3 functions - 1/3, 2/3, 3/3 of project story fetches) -- built as hacky a way around the run time being capped at 30 min on the free tiers of google scripts
function getClubhouseStoriesStart(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Stories"); //The name of the sheet tab where you are sending the info

  clearSheet(sheet);

  var projects = ss.getSheetByName("Projects"); // Access project sheet for project id's
  var projectAmount = projects.getLastRow(); // Get Number of Project
  var projectAmount = Math.round(projectAmount / 3); // Split projects
  var start = 2; // Start on row 2

  getClubhouseProjectStories(projects, start, projectAmount, sheet, apiToken);
}

// Second fetch for last half of stories
function getClubhouseStoriesEnd(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Stories"); //The name of the sheet tab where you are sending the info
  var projects = ss.getSheetByName("Projects"); // Access project sheet for project id's
  var projectAmountFull = projects.getLastRow(); // Get Number of Project
  var projectAmount = Math.round(projectAmountFull / 3); // Splt projects
  var start = projectAmount + 1; // Start on half plus 1 so we do not repeat the same fetch as pervious

  getClubhouseProjectStories(
    projects,
    start,
    projectAmountFull,
    sheet,
    apiToken
  );
}

// Second fetch for last half of stories
function getClubhouseStoriesMid(apiToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  var sheet = ss.getSheetByName("Stories"); //The name of the sheet tab where you are sending the info
  var projects = ss.getSheetByName("Projects"); // Access project sheet for project id's
  var projectAmountFull = projects.getLastRow(); // Get Number of Project
  var projectAmount = Math.round(projectAmountFull / 3); // Splt projects
  var start = projectAmount + projectAmount + 1; // Start on half plus 1 so we do not repeat the same fetch as pervious

  getClubhouseProjectStories(
    projects,
    start,
    projectAmountFull,
    sheet,
    apiToken
  );
}

// Fetch to update for sprint (Excludes Stories)
function getSprint() {
  var apiToken = apiTokenConfig();
  getClubhouseLabels(apiToken);
  getClubhouseEpics(apiToken);
  getClubhouseLabels(apiToken);
  getClubhouseMilestones(apiToken);
}

// Fetch to update process
function updateProcess() {
  var apiToken = apiTokenConfig();
  getClubhouseWorkflowEpics(apiToken);
  getClubhouseWorkflows(apiToken);
  getClubhouseMembers(apiToken);
  getClubhouseProjects(apiToken);
}

// First Fetch for stories
function getStoryFetch1() {
  var apiToken = apiTokenConfig();
  getClubhouseStoriesStart(apiToken);
}

// Second half - Fetch for stories
function getStoryFetch2() {
  var apiToken = apiTokenConfig();
  getClubhouseStoriesEnd(apiToken);
}

// Second half - Fetch for stories
function getStoryFetch3() {
  var apiToken = apiTokenConfig();
  getClubhouseStoriesMid(apiToken);
}
