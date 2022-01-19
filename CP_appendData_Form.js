// TRIGGER TO LAUNCH THE DATA ELABORATION, ON FORM SUBMISSION
function onFormSubmit_CP() {
    appendDataElaborated_CP()
   }
   
   // DATA ELABORATION FROM THE FORM SUBMISSION
   function appendDataElaborated_CP() {
   
     // INTRODUCTION AND DECLARATION
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     const sCP = ss.getSheetByName("ComposizioneProgetti");
     const sCPF = ss.getSheetByName("ComposizioneProgettiForm");
   
     var lastRowCPF = sCPF.getLastRow();
   
     // DATA EXPANSION
     // spread the cell with all the names coming from the form to the right, separating all the names 
     var NameListCell = columnToLetter(letterToColumn("C"))+lastRowCPF; //  it's always the C column we want to expand
     var SpreadedStartNameListCell = columnToLetter(letterToColumn("G"))+lastRowCPF; // arbitrary started in G
     sCPF.getRange(SpreadedStartNameListCell).activate(); 
     var formula = '=arrayformula(trim(split(' + NameListCell + ';",";true;true)))';
     sCPF.getCurrentCell().setFormula([formula]);
   
     // SAVE THE NAME OF THE CURRENT PROJECT UNDER ELABORATION
     var NameProjectCell = columnToLetter(letterToColumn("B"))+lastRowCPF;
     var ProjectCodeAndName = sCPF.getRange(NameProjectCell).getValues();
     var ProjectCodeAndNameArray = ProjectCodeAndName[0][0].split(' ');
   
     // APPEND PART
     // put the names in column on the correct sheet and column
     var memberNumber = getLastNonEmptyCellInRow_CPF()
     for (var i=0; i<memberNumber; i++) { 
       var letCol = letterToColumn("G"); // arbitrary started in G
       var colLet = columnToLetter(letCol+i);
       var NameMembers = sCPF.getRange(colLet+lastRowCPF).getValues();
   
       // CHECK REPETITION IN THE LIST OF THE NAME
       // if positive (-> there is not the same name associated to the same event name), then proceed
       if (repetitionCheck_CP(NameMembers[0][0],ProjectCodeAndNameArray[0])) {
         // if true -> not repeated -> then proceed with appending
         sCP.appendRow([null,ProjectCodeAndNameArray[0],NameMembers[0][0]]); // so that it fill the 3rd column with names, the actual partecipants and the 2nd column with the Name of the Event
   
         // AUTO ID STUFF
         var lastRowCP = sCP.getLastRow();
         var PREautoIDCell = columnToLetter(letterToColumn("A"))+(lastRowCP-1);
         var PREautoID = sCP.getRange(PREautoIDCell).getValues();
         debugger
         var POSTautoIDCell = columnToLetter(letterToColumn("A"))+(lastRowCP);
         sCP.getRange(POSTautoIDCell).setValue(PREautoID[0][0]+1);
       }
     }
   
     // CELEBRATION
     ss.toast('Complimenti stellina!      L\'area audit ti ringrazia della gentile collaborazione e ti augura una buona giornata!')
   
   }