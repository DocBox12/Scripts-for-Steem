function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Nasz dodatek') 
      .addItem('Pobierz przebieg', 'func') 
      .addToUi();
}

function func() {
  var ss = SpreadsheetApp.openById("LINK_DO_ARKUSZA");
  var KILOMETRY = ss.getSheetByName("Odpowiedzi");
  var arkusz_ostatnie_kilometry = ss.getSheetByName("Ostatni");
  
  var i = 2
  
  for (i; i<1000000; i++) {
    var range = KILOMETRY.getRange(i,3);
      var data_from_cell = range.getValue();
      if (data_from_cell == "") {
        break
      } 
  }
  
  var j = i-1
  var range = KILOMETRY.getRange(j,3);
  var latest_data_from_cell = range.getValue();
        
     
  var range2 = arkusz_ostatnie_kilometry.getRange(1,2);
  range2.setValue(latest_data_from_cell);
  
  Browser.msgBox(latest_data_from_cell);
  return
}
 
