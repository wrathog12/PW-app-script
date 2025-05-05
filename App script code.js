function createFoldersAndLinks() {
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    
    var codes = sheet.getRange("A2:A" + sheet.getLastRow()).getValues(); 
    
    
    for (var i = 0; i < codes.length; i++) {
      var code = codes[i][0];  // Get the code from the current row
      
      if (code && sheet.getRange(i + 2, 2).getValue() === '') {  
        
        var folder = DriveApp.createFolder(code);
        
        
        var folderLink = folder.getUrl();
        
        
        sheet.getRange(i + 2, 2).setValue(folderLink);
      }
    }
  }
  