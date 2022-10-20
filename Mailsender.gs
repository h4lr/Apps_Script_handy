// This ui.createMenu is for adding a clickable button on the top on the page
// function onOpen Automatically Triggers when user opens spreadsheet. At the same time this means that 1 I am not supposed to manually execute a trigger function. This particular one is activated when you open the spreadsheet. 

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('➜ Send Emails')
      .addItem('Send Emails xxxxxx explain more', 'sendEmails')
      .addSeparator()
      .addSubMenu(ui.createMenu('Read Me')
          .addItem('How-to', 'Item2'))
      .addToUi();
}

// Returns a rectangular grid of values in a given sheet. param {string} sheetName The name of the sheet object that contains the data to be processed. return {object[][]} A two-dimensional array of // values in the sheet.

function Item2() {
  SpreadsheetApp.getUi() // For information purposes
     .alert('This Script is automatic. It transfers the Range "AO4:AS17" into the tab "Consolidation Domains" and "AO35:AS35" into the tab "Global Domain" so the timeline can be automatically built. It is activated each 15 days. However, if needed, it can be activated by the button "➜ Manually Transfer"');
}


function getData(sheetName) {
  var data = SpreadsheetApp.getActive().getSheetByName(sheetName).getDataRange().getValues();
  return data;
}

// Renders a template with values from an object. Parameter {string} template The template to render. Parameter {object} data The object containing data to render the template.
// Then return the {string} The rendered template

function renderTemplate(template, data) {
  var output = template;
  var params = template.match(/\{\{(.*?)\}\}/g);
  params.forEach(function (param) {
    var propertyName = param.slice(2,-2); //Remove the {{ and the }}
    output = output.replace(param, data[propertyName] || "");
  });
  return output;
}

//Grid of values transformed into array of objects. {object[][]} rows is An array of rows in the grid. Then return {object[]} An array of objects so each row became an object
function rowsToObjects(rows) {
  var headers = rows.shift();
  var data = [];
  rows.forEach(function (row) {
    var object  = {};
    row.forEach(function (value, index) {
      object[headers[index]] = value;
    });
    data.push(object);
  });
  return data;
}

function sendEmails() {  // Send mail for each row
  var templateData = getData("Templates");
  var emailSubjectTemplate = templateData[1][0]; //Cell A2
  var emailBodyTemplate = templateData[4][0]; //Cell A5

  var emailData = getData("Data");

  emailData = rowsToObjects(emailData);

  emailData.forEach(function (rowObject) {
    
    var mails  = Browser.inputBox('Enter "CC:" emails followed by a ","', Browser.Buttons.OK_CANCEL);  //Variable for the different mails
    var mandatory  = "etc@gmail.com,etc@gmail.com"; // I can either put this to be an input from user or fixed, I believe it is better a fixed value ()
    var subject = renderTemplate(emailSubjectTemplate, rowObject);
    var body = renderTemplate(emailBodyTemplate, rowObject);

    MailApp.sendEmail({to:mandatory,subject,body,cc: mails });
    
  
  })
}
