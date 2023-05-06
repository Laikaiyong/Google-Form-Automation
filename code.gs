/**
 * Sends a customized email for every response on a form.
 *
 * @param {Object} e - Form submit event
 */
async function onFormSubmit(e) {
  //e is a form submit event from Google Forms, use nameValuePairs to get the responses
  var itemResponses = e.namedValues;
  var firstName = itemResponses[APUBCC_MEMBERSHIP_RESPONSES_COLUMN_FIRST_NAME][0]
  var tpNumber = itemResponses[APUBCC_MEMBERSHIP_RESPONSES_COLUMN_TP_NUMBER][0]

  var discordInviteLink = "";
  var telegramInviteLink = "";

  var taskListId = await getTaskListId();
  var task = await createTask(
    taskListId,
    firstName,
    tpNumber,
    discordInviteLink,
    telegramInviteLink,
    APUBCC_OFFICIAL_TELEGRAM_INVITE_DURATION
  );
  if (task) {
    await updateApplicationStatus(tpNumber, "pass");
  } else {
    await updateApplicationStatus(tpNumber, "fail");
  }
}


/**
 * Gets the task list ID for the current month.
 *
 * @return {string} - The task list ID.
 */
async function getTaskListId() {
  var taskListId = "";

  var taskLists = Tasks.Tasklists.list();
  var taskListTitle = "New Member Signups";
  var taskListExists = false;

  for (var i = 0; i < taskLists.items.length; i++) {
    if (taskLists.items[i].title == taskListTitle) {
      taskListExists = true;
      taskListId = taskLists.items[i].id;
      break;
    }
  }

  if (!taskListExists) {
    taskListId = await createTaskList(taskListTitle);
  }

  return taskListId;
}

/**
 * Creates a task list.
 *
 * @param {string} taskListTitle - The title of the task list.
 * @return {string} - The task list ID.
 */
async function createTaskList(taskListTitle) {
  var taskList = Tasks.Tasklists.insert({
    title: taskListTitle,
  });

  return taskList.id;
}

/**
 * Creates a task.
 *
 * @param {string} taskListId - The ID of the task list.
 * @param {string} firstName - The first name of the applicant.
 * @param {string} tpNumber - The TP number of the applicant.
 * @param {string} discordInviteLink - The Discord invite link.
 * @param {string} telegramInviteLink - The Telegram invite link.
 * @param {string} telegramInviteDuration - The Telegram invite duration.
 *
 * @return {string} - The task ID.
 */
async function createTask(
  taskListId,
  firstName,
  tpNumber,
  discordInviteLink,
  telegramInviteLink,
  telegramInviteDuration
) {
  var tpEmail = `${tpNumber}@mail.apu.edu.my`
  var task = Tasks.Tasks.insert(
    {
      title: firstName.charAt(0).toUpperCase() + firstName.slice(1),
      notes:
        tpEmail +
        ", " +
        discordInviteLink +
        ", " +
        telegramInviteLink +
        ", " +
        telegramInviteDuration,
    },
    taskListId
  );

  return task.id;
}


/**
 * Updates the responses spreadsheet with the status of the application. Find row by tpMail, if pass, highlight the entire row green. If fail, highlight the entire row red.
 *
 * @param {string} tpNumber - The TP Number to identify the row to highlight
 * @param {string} status - The email send status of this TP Number
 */
function updateApplicationStatus(tpNumber, status) {
  //Find spreadsheet by ID
  var ss = SpreadsheetApp.openById(
    APUBCC_MEMBERSHIP_RESPONSES_GOOGLE_SPREADSHEET_ID
  );

  //Get the sheet by name
  var sheet = ss.getSheetByName(APUBCC_MEMBERSHIP_RESPONSES_GOOGLE_SHEET_NAME);

  //Find the row containing the tpNumber, then update the status column which is the last column by appending the status. EMAIL SENT if status is pass, EMAIL NOT SENT if status is fail.
  //Find which column number has the title 'TP Number'
  var tpNumberColumnIndex = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .indexOf(APUBCC_MEMBERSHIP_RESPONSES_COLUMN_TP_NUMBER);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var values = range.getValues();

  //Highlight the entire row green if status is pass, red if status is fail.
  for (var i = 0; i < values.length; i++) {
    if (values[i][tpNumberColumnIndex] == tpNumber) {
      if (status === "pass") {
        sheet.getRange(i + 1, 1, 1, lastColumn).setBackground("#00FF00");
      } else {
        sheet.getRange(i + 1, 1, 1, lastColumn).setBackground("#FFFF00");
      }
    }
  }
}

//A function that gets one or more highligted rows from the responses spreadsheet and send the email to the applicant, then update the status of the application to pass.
async function manualKYC() {
  //This script will be loaded into the spreadsheet, so a user will highlight the rows they want to send the email to, then run this script.
  //The script's container will be the spreadsheet, so we need to get the spreadsheet ID from the spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //Find the active sheet, then get the highlighted rows.
  var sheet = ss.getActiveSheet();
  var highlightedRows = sheet.getActiveRange().getValues();

  //Get the column number of the TP Number column
  var tpNumberColumnIndex = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .indexOf(APUBCC_MEMBERSHIP_RESPONSES_COLUMN_TP_NUMBER);

  //Get the column number of the first name column
  var firstNameColumnIndex = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .indexOf(APUBCC_MEMBERSHIP_RESPONSES_COLUMN_FIRST_NAME);

  //Get the column number of the email column
  var emailColumnIndex = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .indexOf(APUBCC_MEMBERSHIP_RESPONSES_COLUMN_EMAIL);

  //From the highlighted rows, get the TP Number, first name and email of the applicant.
  for (var i = 0; i < highlightedRows.length; i++) {
    var discordInviteLink = "https://discord.com/channels/864498442537402368/992415598048985159/1091027957360894074";
    var telegramInviteLink = "https://discord.com/channels/864498442537402368/992415598048985159/1091027957360894074";
    var tpNumber = highlightedRows[i][tpNumberColumnIndex];
    var firstName = highlightedRows[i][firstNameColumnIndex];

    //Use the Tasks API to check if there exists an item with the same name as the TP Number, if yes, then delete the item and create a new item with the same name. if no, then create a new item with the same name.

    var taskListId = await getTaskListId();
    var taskListItems = await getTaskListItems(taskListId);
    var taskListItemId = "";
    for (var j = 0; j < taskListItems.length; j++) {
      if (
        taskListItems[j].title ==
        firstName.charAt(0).toUpperCase() + firstName.slice(1)
      ) {
        taskListItemId = taskListItems[j].id;
        break;
      }
    }

    if (taskListItemId != "") {
      await deleteTaskListItem(taskListId, taskListItemId);
    }

    await createTask(
      taskListId,
      firstName.charAt(0).toUpperCase() + firstName.slice(1),
      tpNumber,
      discordInviteLink,
      telegramInviteLink,
      APUBCC_OFFICIAL_DISCORD_INVITE_DURATION
    );

    //Update the status of the application to pass
    updateApplicationStatus(tpNumber, "pass");
  }
}

/**
 * A function that gets all items in the task list and return the task list items.
 *
 * @param {string} taskListId The ID of the task list.
 */
async function getTaskListItems(taskListId) {
  var taskListItems = [];
  var nextPageToken = "";
  do {
    var taskListItemsResponse = await Tasks.Tasklists.list({
      tasklist: taskListId,
      pageToken: nextPageToken,
    });
    taskListItems = taskListItems.concat(taskListItemsResponse.items);
    nextPageToken = taskListItemsResponse.nextPageToken;
  } while (nextPageToken);

  return taskListItems;
}