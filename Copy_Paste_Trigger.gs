// This ui.createMenu is for adding a clickable button on the top on the page
// function onOpen Automatically Triggers when user opens spreadsheet. At the same time this means that 1 I am not supposed to manually execute a trigger function. This particular one is activated when you open the spreadsheet. 

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('➜ Manually Transfer')
      .addItem('Transfer ETC)', 'CopyRange')
      .addSeparator()
      .addItem('Transfer ETC)', 'CopyRange2')
      .addSeparator()
      .addItem('Transfer ETC)', 'CopyRange3')
      .addSeparator()
      .addSubMenu(ui.createMenu('Read Me')
          .addItem('How-to', 'Item2'))
      .addToUi();
}

function Item2() {
  SpreadsheetApp.getUi() // For information purposes
     .alert('This Script is automatic. It transfers the Range "AO4:AS17" into the tab "Consolidation Domains" and "AO35:AS35" into the tab "Global Domain" so the timeline can be automatically built. It is activated each 15 days. However, if needed, it can be activated by the button "➜ Manually Transfer"');
}

function CopyRange() {
 var sss = SpreadsheetApp.openById('ENTERID'); //To be replaced by SpreadsheetID
 var ss = sss.getSheetByName('Data'); //Replace 'Data' with the source Sheet tab name
 var range = ss.getRange('AO4:AS17'); //Assign the range to be copied
 var data = range.getValues();

 var tss = SpreadsheetApp.openById('ENTERID'); //To be replaced by SpreadsheetID
 var ts = tss.getSheetByName('Consolidation Domains'); //Replace with destination Sheet tab name


 ts.getRange(ts.getLastRow()+1, 1,14,5).setValues(data); //you will need to define the size of the copied data see getRange()
}

function CopyRange2() {
 var sss = SpreadsheetApp.openById('ENTERID'); //To be replaced by SpreadsheetID
 var ss = sss.getSheetByName('Data'); //Replace 'Data' with the source Sheet tab name
 var range = ss.getRange('AO35:AS35'); //Assign the range to be copied
 var data = range.getValues();

 var tss = SpreadsheetApp.openById('ENTERID'); //To be replaced by SpreadsheetID
 var ts = tss.getSheetByName('Global Domain'); //Replace with destination Sheet tab name


 ts.getRange(ts.getLastRow()+1, 1,1,5).setValues(data); //you will need to define the size of the copied data see getRange()
}


function CopyRange3() {
 var sss = SpreadsheetApp.openById('ENTERID'); //To be replaced by SpreadsheetID
 var ss = sss.getSheetByName('Data'); //Replace 'Data' with the source Sheet tab name
 var range = ss.getRange('AO18:AS33'); //Assign the range to be copied
 var data = range.getValues();

 var tss = SpreadsheetApp.openById('ENTERID'); //To be replaced by SpreadsheetID
 var ts = tss.getSheetByName('Consolidation Organization'); //Replace with destination Sheet tab name


 ts.getRange(ts.getLastRow()+1, 1,16,5).setValues(data); //you will need to define the size of the copied data see getRange()
}
