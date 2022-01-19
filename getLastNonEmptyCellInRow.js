///////////////////////// PARTECIPAZIONE EVENTI FORM /////////////////////////

function getLastNonEmptyCellInRow_PEF() {
  

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sPEF = ss.getSheetByName("PartecipazioneEventiForm");
  
    var maxColumns = sPEF.getLastColumn(); // in all the sheet, conservative but ok
    var rowToCheck = sPEF.getLastRow();
  
    var rowData = sPEF.getRange(rowToCheck, 1, 1, maxColumns).getValues(); //take that row and till the maxColumn
    rowData = rowData[0]; //Get inner array of two dimensional array
  
    var rowLength = rowData.length;
    var countRows = -3; // initialise the counter as -3 because we know that the first 3 elements will be
    // the " Informazioni cronologiche |	Nome |	Nome e Cognome " coming from the form
    for (var i = 0; i < rowLength; i++) {
      var thisCellContents = rowData[i];
  
      if (thisCellContents != "") {
        countRows = countRows + 1;
      }
    }
  
    return countRows
  }
  
  
  ///////////////// COMPOSIZIONE PROGETTI FORM ///////////////////
  function getLastNonEmptyCellInRow_CPF() {
    
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sCPF = ss.getSheetByName("ComposizioneProgettiForm");
  
    var maxColumns = sCPF.getLastColumn(); // in all the sheet, conservative but ok
    var rowToCheck = sCPF.getLastRow();
  
    var rowData = sCPF.getRange(rowToCheck, 1, 1, maxColumns).getValues(); //take that row and till the maxColumn
    rowData = rowData[0]; //Get inner array of two dimensional array
  
    var rowLength = rowData.length;
    var countRows = -3; // initialise the counter as -3 because we know that the first 3 elements will be
    // the " Informazioni cronologiche |	ID & Nome Progetto |	Nome e Cognome " coming from the form
    for (var i = 0; i < rowLength; i++) {
      var thisCellContents = rowData[i];
  
      if (thisCellContents != "") {
        countRows = countRows + 1;
      }
    }
  
    return countRows
  }
  
  
  ///////////////// PROGETTI ////////////////////////////////////
  function getLastNonEmptyCellInRow_P(row) { //row
    
    // var row = 151; // debugging purposes
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sP = ss.getSheetByName("Progetti");
  
    var maxColumns = sP.getLastColumn(); // in all the sheet, conservative but ok
    // var lastRow = sP.getLastRow();
  
    const headerCount = 14; // how many header are there, filled by ClickUp, if you change something, change also here, N=14, where the team is
  
    var rowData = sP.getRange(row, 1, 1, maxColumns).getValues(); //take that row and till the maxColumn
    rowData = rowData[0]; //Get inner array of two dimensional array
    
    var rowLength = rowData.length;
    var countRows = -1; // initialise the counter as 0
    for (var i = headerCount; i < rowLength; i++) {
      var thisCellContents = rowData[i];
  
      if (thisCellContents != "") {
        countRows = countRows + 1;
      }
    }
    
    return countRows
  }