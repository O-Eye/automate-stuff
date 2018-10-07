/** Based on the code from Eamon Collins from:
https://github.com/eamoncollins/GAS/commit/2f657ad24552cde61cf9ce83f785ab11a8fcffd5 */

/** Edited the script so that:

This scirpt works on Google Sheet populated manually or by using Google Forms.
Two fields are mandatory - Name and Email (2nd and 3rd column). First column is timestamp (can be populated with random data).
The scripts creates folders where not previously created, within the Google Sheet used.
Captures Folder name, Folder ID. Shares the folder with the provided email.

NOTE: not using ROOT - script does not check for ROOT of Google Drive
      but can be modified to do so. 
      You are free to use this script and suggest any modifications.
      I will be happy to learn more from your own ideas as well. */

//create a menu in the spreadsheet (visible once refreshed)
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GDrive')
    .addItem('Create and share folders from this spreadsheet', 'crtShrDrvFldrs')
    .addToUi();
}

//function to create folders
function crtShrDrvFldrs(){
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //change this as required (header in first row)   
  const START_ROW = 2;
  
  //get last populated row
  var numRows = sheet.getLastRow(); 
  
  //increase this to accommodate more rows (contributors)
  var maxRows = Math.min(numRows, 36); //whichever is smaller
  
  //change this for the Parent Folder as required, should already exists
  const FOLDER_NAME = 'SAMPLE FOLDER';
  const ROOT = 'N'; //Not used in this particular script
  const DESC = 'ADD DESCRIPTION HERE'; //change here as required, this is added to folder name
  
  //range of data to be used
  var dataRange = sheet.getRange(START_ROW, 2, maxRows, 4);
  var data = dataRange.getValues();
  
  //find the folder and set it as parent 
  var folderIter = DriveApp.getFoldersByName(FOLDER_NAME);
  if(!folderIter.hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Folder not found:' + FOLDER_NAME);
    return;
  }
  var parentFolder = folderIter.next();  
  if(folderIter.hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Folder has non-unique name!' + FOLDER_NAME);
    return;
  }
  
  //header of other columns
  SpreadsheetApp.getActiveSheet().getRange(1, 4).setValue('FolderCreated?');
  SpreadsheetApp.getActiveSheet().getRange(1, 5).setValue('FolderName');
  SpreadsheetApp.getActiveSheet().getRange(1, 6).setValue('FolderID');
  
  //main code for creating and sharing folders
  var k = 0;
  for (i in data) {  	
    var row = data[i];    
    var uname = row[0]; //column 1 of data range
    var email = row[1];
    var isCreated = row[2];    
    if(uname != '' && isCreated != 'Y') {
    
      var idNewFolder = parentFolder.createFolder(uname +' '+ DESC).setDescription(DESC).getId();      
      Utilities.sleep(250);  //pause
      
      var newFolder = DriveApp.getFolderById(idNewFolder);
      
      SpreadsheetApp.getActiveSheet().getRange((k+START_ROW), 6).setValue(idNewFolder);
      SpreadsheetApp.getActiveSheet().getRange((k+START_ROW), 5).setValue(newFolder.getName());
      newFolder.addEditor(email); //give access to the respective user
      SpreadsheetApp.getActiveSheet().getRange((k+START_ROW), 4).setValue('Y'); //iscreated and shared
      
      SpreadsheetApp.flush();     //write to cells immediately 
    }//if
    k = k + 1;
  }/for
}//main
