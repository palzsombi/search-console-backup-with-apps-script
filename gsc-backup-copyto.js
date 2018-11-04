/*********************
* GOOGLE APPS SCRIPT *
**********************

Munkalap kiválasztása aktuális hónap alapján
+ Adatok másolása céltartományba 


Készítette: Pál Zsombor 

Módosítandó variánsok:

- 24. sor - destinationSheetName
- 26. sor - startRow
- 27. sor - startCol
- 82. sor - destinationRangeClear

*/

function copyTo() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var destinationSheetName = 'destination';     // A 'destination' helyére a céltartomány munkalapjának neve kerül
  
  var startRow = 1;     // Másolandó adattartomány első sorának száma                                
  var startCol = 1;     // Másolandó adattartomány első oszlopának száma    
  
  var date = new Date();     // Mai nap
  
  var month = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM");     // Aktuális hónap
  var year = Utilities.formatDate(date, Session.getScriptTimeZone(), "YYYY");     // Aktuális év
  
  var lastMonth = (function()     // Előző hónap meghatározása
    {
      var date = new Date();
      date.setMonth(date.getMonth() - 1);
    
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM');
      
    })();
  
  var lastYear = (function()     // Előző év meghatározása
    {
      var date = new Date();
      date.setYear(date.getYear() - 1);
    
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'YYYY');
      
    })();
  
  // Forrás munkalap azonosítása Januártól eltérő hónap esetén
  
  if (month != 'Jan')
  {    
    var sourceSheetName = lastMonth + " " + year;
    
  }
  
  // Forrás munkalap azonosítása Január esetén
  
  if (month == 'Jan')
  { 
    var sourceSheetName = lastMonth + " " + lastYear;
    
  }
  
  // Forrás és céltartomány meghatározása
  
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  
  var lastRow = sourceSheet.getLastRow();          
  var lastCol = sourceSheet.getLastColumn();
  
  var destinationSheet = ss.getSheetByName(destinationSheetName);
  var destinationRange = destinationSheet.getRange(startRow, startCol, lastRow, lastCol);
  
  var sourceRange = sourceSheet.getRange(startRow, startCol, lastRow, lastCol);
  
  // Korábbi adatok törlése
  
  var destinationRangeClear = destinationSheet.getRange("A2:F");     // Törlendő adattartomány meghatározása
  
  destinationRangeClear.clear();
  
  // Adatok másolása
    
  sourceRange.copyTo(destinationRange);
  
  // Adatok ellenőrzése
  
  Logger.log(month);
  Logger.log(year);
  
  Logger.log(lastMonth);
  Logger.log(lastYear);
  
}
  
