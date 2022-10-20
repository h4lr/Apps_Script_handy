
function Item1() { // We create the input box for adding the data
  var folderid  = Browser.inputBox('Enter folder ID', Browser.Buttons.OK_CANCEL);
  var sh = SpreadsheetApp.getActiveSheet(); // Let's clean the Sheet before
  sh.clear();                               // Let's clean the Sheet before
  sh.appendRow(["Parent","Folder", "Name", "Update", "URL", "Type"]);  // Adding some columns
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Task started', 'Please Wait', 10);
    var parentFolder =DriveApp.getFolderById(folderid);   // DriveApp is a class. Class DriveApp allows scripts to create, find, and modify files and folders in Google Drive.
    listFiles(parentFolder,parentFolder.getName())        // getFolderById: Gets the folder with the given ID.
    listSubFolders(parentFolder,parentFolder.getName());  //Don't allow the user to see any exceptions your code throws. We use try...catch statements to intercept exceptions
  } catch (error) {                                           //then display a user-friedly error message with inline text styled in the error class from the add-ons CSS package or an alert dialog. 
    console.error(error);  // If we stop the sequence it will just display, "Finalized" instead of giving an error
    
  }
}


function listSubFolders(parentFolder,parent) {
  var childFolders = parentFolder.getFolders();
  while (childFolders.hasNext()) {
    var childFolder = childFolders.next();
    Logger.log("Fold : " + childFolder.getName());
    listFiles(childFolder,parent)
    listSubFolders(childFolder,parent + "|" + childFolder.getName());
  }
}

function listFiles(fold,parent){
  var sh = SpreadsheetApp.getActiveSheet();
  var data = [];
  var files = fold.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    data = [ 
      parent,
      fold.getName(),
      file.getName(),
      file.getLastUpdated(),
      file.getUrl(),
      file.getMimeType()
      ];
    sh.appendRow(data);
  }
} 
