// TRIGGER TO LAUNCH THE DATA ELABORATION, ON FORM SUBMISSION
function onFormSubmit_PE() {
    appendDataElaborated()
   }
   
   // DATA ELABORATION FROM THE FORM SUBMISSION
   function appendDataElaborated() {
   
     // INTRODUCTION AND DECLARATION
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     const sPE = ss.getSheetByName("PartecipazioneEventi");
     const sPEF = ss.getSheetByName("PartecipazioneEventiForm");
   
     var lastRowPEF = sPEF.getLastRow();
   
     // DATA EXPANSION
     // spread the cell with all the names coming from the form to the right, separating all the names 
     var NameListCell = columnToLetter(letterToColumn("C"))+lastRowPEF; //  it's always the C column we want to expand
     var SpreadedStartNameListCell = columnToLetter(letterToColumn("G"))+lastRowPEF; // arbitrary started in G
     sPEF.getRange(SpreadedStartNameListCell).activate(); 
     var formula = '=arrayformula(trim(split(' + NameListCell + ';",";true;true)))';
     sPEF.getCurrentCell().setFormula([formula]);
     
     // SAVE THE NAME OF THE CURRENT EVENT UNDER ELABORATION
     var NameEventCell = columnToLetter(letterToColumn("B"))+lastRowPEF;
     var NameEvent = sPEF.getRange(NameEventCell).getValues();
   
     // APPEND PART
     // put the names in column on the correct sheet and column
     var partecipantNumber = getLastNonEmptyCellInRow_PEF()
     for (var i=0; i<partecipantNumber; i++) { 
       var letCol = letterToColumn("G"); // arbitrary started in G
       var colLet = columnToLetter(letCol+i);
       var NamePartecipants = sPEF.getRange(colLet+lastRowPEF).getValues();
   
       // CHECK REPETITION IN THE LIST OF THE NAME
       // if positive (-> there is not the same name associated to the same event name), then proceed
       if (repetitionCheck_PE(NamePartecipants[0][0],NameEvent[0][0])) {
         // if true -> not repeated -> then proceed with appending
         sPE.appendRow([null,NameEvent[0][0],NamePartecipants[0][0]]); // so that it fill the 3rd column with names, the actual partecipants and the 2nd column with the Name of the Event
   
         // AUTO ID STUFF
         var lastRowPE = sPE.getLastRow();
         var PREautoIDCell = columnToLetter(letterToColumn("A"))+(lastRowPE-1);
         var PREautoID = sPE.getRange(PREautoIDCell).getValues();
         var POSTautoIDCell = columnToLetter(letterToColumn("A"))+(lastRowPE);
         sPE.getRange(POSTautoIDCell).setValue(PREautoID[0][0]+1);
       }
     }
   
     // CELEBRATION
     ss.toast('Complimenti stellina!      L\'area audit ti ringrazia della gentile collaborazione e ti augura una buona giornata!')
   
   }