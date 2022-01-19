// TRIGGER TO LAUNCH THE DATA ELABORATION, ON FORM SUBMISSION
/* 
function onEdit(e) {
  
  var ss = e.source;
  var sheetName = ss.getActiveSheet().getName();
  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();

  if(sheetName == 'Progetti' && column == letterToColumn('Z')) { // "N" is the column with the team // "Z" is the IMPORTRANGE cell
  // because the onEdit event doesn't recognise the API edit of a cell, but should the IMPORTRANGE ones
  // so the trigger is built on the "Z" column
    appendDataElaborated_ComposizioneProgetti(row); 
  } 
}

*/

// DATA ELABORATION
function appendDataElaborated_ComposizioneProgetti(row) { //row
    // var row = 151; // for debugging purpose
    // INTRODUCTION AND DECLARATION
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sP = ss.getSheetByName("Progetti");
    const sCP = ss.getSheetByName("ComposizioneProgetti");
  
    //var lastRowP = sP.getLastRow();
  
    // DATA EXPANSION
    // find the "Team", "Expansion Area" and "Codice Progetto" headers column number, to have flexible implementation to clickup changes and future database maintenace and entity changes
    const [headerP, ...dataP] = sP.getDataRange().getDisplayValues();
    var headerSize = headerP.length;
    var idxCodiceProgetto = 0;
    var idxTeam = 0;
    var idxExpansionArea = 0;
    for (var i=0; i<=headerSize; i++) { 
      if (headerP[i] == "IMPORTRANGE") { //IMPORTRANGE
        idxIMPORTRANGE = i+1; // +1 because the counting start from 0 but the index from 1
      }
      if (headerP[i] == "Team") {
        idxTeam = i+1;
      }
      if (headerP[i] == "Expansion Area") {
        idxExpansionArea = i+1;
      }
      if (headerP[i] == "Codice Progetto") {
        idxCodiceProgetto = i+1;
      }
    }
  
    // spread the cell with all the names coming from ClickUp to the right, separating all the names 
    var NameListCell = columnToLetter(idxIMPORTRANGE)+row; //  the team, column we want to expand
    var TeamCell = columnToLetter(idxTeam)+row; //  the team, column we want to expand
    var SpreadedStartNameListCell = columnToLetter(idxExpansionArea)+row; // arbitrary started in AA, far away from the first used cells
    sP.getRange(SpreadedStartNameListCell).activate(); 
    // var formula = '=arrayformula(trim(split(' + NameListCell + ';";";true;true)))';
    var formula2 = '=IF(ISBLANK(' + TeamCell + ');"";arrayformula(trim(split(' + NameListCell + ';";";true;true))))';
    sP.getCurrentCell().setFormula([formula2]);
  
    // SAVE THE NUMBER OF THE CURRENT PROJECT UNDER ELABORATION
    var CodiceProgettoCell = columnToLetter(idxCodiceProgetto)+row;
    var CodiceProgetto = sP.getRange(CodiceProgettoCell).getValues();
  
    // APPEND PART
    // Calculate the size of the project by counting the spreaded names
    var projectSize = getLastNonEmptyCellInRow_P(row);
    
    // Put the names in column on the correct sheet and column
    for (var i=0; i<projectSize; i++) { 
      var letCol = idxExpansionArea; // arbitrary started in AA, far away from the first used cells
      var colLet = columnToLetter(letCol+i);
      var NameMembers = sP.getRange(colLet+row).getValues();
      
      // CHECK REPETITION IN THE LIST OF THE NAME
      // if positive (-> there is not the same name associated to the same project id), then proceed
      if (repetitionCheck_CP(NameMembers[0][0],CodiceProgetto[0][0])) {
        // if true -> not repeated -> then proceed with appending
        sCP.appendRow([null,CodiceProgetto[0][0],NameMembers[0][0]]); // so that it fill the 3rd column with names, the actual partecipants
        
        // AUTO ID STUFF
        var lastRowCP = sCP.getLastRow();
        var PREautoIDCell = columnToLetter(letterToColumn("A"))+(lastRowCP-1); // always the id is in the first column!
        var PREautoID = sCP.getRange(PREautoIDCell).getValues();
        var POSTautoIDCell = columnToLetter(letterToColumn("A"))+(lastRowCP);
        sCP.getRange(POSTautoIDCell).setValue(PREautoID[0][0]+1);
      }
    }
  
    // CELEBRATION
    ss.toast('Complimenti stellina!      L\'area audit ti ringrazia della gentile collaborazione e ti augura una buona giornata!')
  
  }