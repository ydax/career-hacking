///////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////// Eazl Career Hacking™ App ///////////////////////////////////////////
////////////////////////// Enroll in the class: ///////////////////////////////////////////////
////////////////////////// On Eazl https://courses.eazl.co/p/career-hacking ///////////////////
////////////////////////// On Udemy https://www.udemy.com/course/golden-gate-bridge/ //////////
///////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////// Developed by Davis Jones @ github.com/ydax /////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////
///////////////////////// MENU & TRIGGER SETUP ////////////////////////
///////////////////////////////////////////////////////////////////////

function onOpen(e) {
  // sets up the UI
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🔍 Career Hacking™")
  .addItem("🔔 Activate Follow-up Helper", "activate")
  .addItem("📧 Set Email Address", "setEmailAddress")
  .addToUi();
}

// called from UI
function activate () {
  // look for the app's required time-based trigger and sets it up if not present
  var triggerPresent = triggerCheck();
  if (triggerPresent != true) {
    ScriptApp.newTrigger('sendUpdateEmail')
    .timeBased()
    .everyDays(1)
    .create();
  }
  var emailStatus = checkEmail();
  SpreadsheetApp.getUi().alert('✓ Your automatic follow-up reminder email trigger is set up.\n' + emailStatus);
}

///////////////////////////////////////////////////////////////////////
////////// STRUCTURE AND SEND AUTOMATIC FOLLOW-UP REMINDERS ///////////
///////////////////////////////////////////////////////////////////////

// return array of rows matching today for update
function getUpdateRows () {
  var ss = SpreadsheetApp.getActive();
  var followUpSheet = ss.getSheetByName('Following Up');
  
  // get the date array from Follow Up sheet
  var followUpDateArray = followUpSheet.getRange(4, 11, followUpSheet.getLastRow(), 1).getValues();
  
  // cycle through the array and push the row # of any dates that equal today
  var i = 4; // iterator
  var today = getTodaySubstr();
  var rowsForUpdate = [];
  followUpDateArray.forEach(function (date) {
    date = date[0].toString().substring(4, 10);
    if (date == today) {
      rowsForUpdate.push(i);
    }
    i++;
  })
  
  return rowsForUpdate;
}

function sendUpdateEmail () {
  var ss = SpreadsheetApp.getActive();
  var followUpSheet = ss.getSheetByName('Following Up');
  var updateRows = getUpdateRows();
  var textSnippets = [];
  
  if (updateRows.length != 0) {
    updateRows.forEach(function (row) {
      Logger.log(row);
      var firstName = followUpSheet.getRange(row, 1).getValue();
      var lastName = followUpSheet.getRange(row, 2).getValue();
      var title = followUpSheet.getRange(row, 3).getValue();
      var organization = followUpSheet.getRange(row, 6).getValue();
      var email = followUpSheet.getRange(row, 7).getValue();
      var linkedin = followUpSheet.getRange(row, 9).getValue();
      
      // structure snippit
      var snippet = '• ' + firstName + lastName + ', ' + title + ' with ' + organization + '\n Available contact info: ' + email + ' ' + linkedin + '\n';
      textSnippets.push(snippet);
    })
    // create and send reminder
    var today = getTodaySubstr();
    var body = 'Hi there,\nHere are your follow-up reminders for today:\n\n';
    textSnippets.forEach(function (snippet) {
      body += snippet;
    });
    body += '\nBest of luck with your job search,\nTeam Eazl';
    
    var myEmail = getMyEmail();
    
    GmailApp.sendEmail(
      myEmail,
      'Your Follow Up Reminders for ' + today,
      body);
  }
}

///////////////////////////////////////////////////////////////////////
////////////////////////////// UTILITIES //////////////////////////////
///////////////////////////////////////////////////////////////////////

// return a string w/ today's month and date
function getTodaySubstr () {
  var today = new Date().toString();
  var todaySubstr = today.substring(4, 10); // grab the month and day of the date obj
  return todaySubstr;
}

// updates the user on the status of their email address
function checkEmail () {
  var scriptProps = PropertiesService.getScriptProperties();
  var myEmail = scriptProps.getProperty('userEmail');
  if (myEmail == null) {
    return '✗ You still need to set up your email address. Go to 🔍 Career Hacking™ > 📧 Set Email Address.';
  } else {
    return 'Your follow-up reminders will be send to ' + myEmail + '. To have them sent to a different address, go to\n🔍 Career Hacking™ > 📧 Set Email Address.'
  } 
}

// see if time-based trigger is in place and add if not
function triggerCheck () {
  // get current project's triggers
  var triggers = ScriptApp.getProjectTriggers();
  var triggerFunctions = [];
  triggers.forEach(function (trigger) {
    var handlerFunction = trigger.getHandlerFunction();
    triggerFunctions.push(handlerFunction);
  });
  
  // check to see if one of the current trigger functions is sendScheduledDrafts and returns
  var appTrigger = function(functionName) { // sets up the filter
    return functionName == 'sendUpdateEmail';
  }
  var triggerPresent = triggerFunctions.some(appTrigger);
  
 return triggerPresent;
}

/* interface enabling user to set their preferred email address
⚠️ This does not give Eazl or anyone else access to your
email address */

function setEmailAddress () {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt('Enter the email address you\'d like\nyour follow-up reminders sent to:');

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    var scriptProps = PropertiesService.getScriptProperties();
    var myEmail = scriptProps.setProperty('userEmail', response.getResponseText());
    myEmail = response.getResponseText();
    ui.alert('✓ All set. You\'ll receive a follow-up reminder at ' + response.getResponseText() + ' whenever the date equals the follow-up date you\'ve set in Column K (Next Follow Up) with that person.')
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function getMyEmail () {
  var scriptProps = PropertiesService.getScriptProperties();
  var myEmail = scriptProps.getProperty('userEmail');
  return myEmail;
}