// TIME TRIGGER TO UPDATE THE FORM
// PROGRAMMATICALLY WAY, INSTEAD OF GOING TO THE TRIGGER PAGE AND SET UP MANUALLY
// THE PROBLEM WITH THIS ONE IS THAT IT CREATE MANY AND MANY TRIGGERS!
/*
function timeDrivenPopulateFormTrigger_CP() { 
  ScriptApp
    .newTrigger("populateGoogleForm_ComposizioneProgetti")
    .timeBased()
    .everyDays(1)
    //.everyHours(12)
    //.everyMinutes(1) // debugging purposes
    .create();
}
*/

// GET DATA FROM THE SHEETS "PROGETTI" AND "SOCI" TO POPULATE THE FORM
function getDataFromGoogleSheets_ComposizioneProgetti() {
    // video tutorial: https://www.youtube.com/watch?v=Z-gCwZ0lXd8
    
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetSoci = ss.getSheetByName("Soci");
      const sheetProgetti = ss.getSheetByName("Progetti");
      const [headerSoci, ...dataSoci] = sheetSoci.getDataRange().getDisplayValues(); //take the same type of data as reported in the cell -> text
      const [headerProgetti, ...dataProgetti] = sheetProgetti.getDataRange().getDisplayValues(); //take the same type of data as reported in the cell -> text
    
      const choicesSoci = {}
      const choicesProgetti = {}
      headerSoci.forEach(
        function(title, index) {
          choicesSoci[title] = dataSoci.map(row => row[index]).filter(e => e !== "");
        }
      );
      headerProgetti.forEach(
        function(title, index) {
          choicesProgetti[title] = dataProgetti.map(row => row[index]).filter(e => e !== "");
        }
      );
    
      return [choicesSoci, choicesProgetti];
    
    }
    
    // ACTUALLY POPULATE THE FORM
    function populateGoogleForm_ComposizioneProgetti() {
    // video tutorial: https://www.youtube.com/watch?v=Z-gCwZ0lXd8
    
      const GoogleFormID = "1TTcKqJxUup6SkjOcrhLjpG5psomhQgdHLfxF-nbizj8";
      const googleForm = FormApp.openById(GoogleFormID);
      const items = googleForm.getItems();
      const [choicesSoci, choicesProgetti] = getDataFromGoogleSheets_ComposizioneProgetti();
    
      // SELECT ONLY ACTIVE ASSOCIATES
      var NumeroTotaleSoci = choicesSoci['Nome e Cognome'].length;
      var SociAttiviArray = []; 
      var SociAttivi = {};
      var idx = 0;
      for (var i=0; i<NumeroTotaleSoci; i++) {
        if (choicesSoci['Stato'][i] == "socio" || choicesSoci['Stato'][i] == "in prova") {
          SociAttiviArray[idx] = choicesSoci['Nome e Cognome'][i];
          idx = idx+1;
        }
      }
      SociAttivi['Nome e Cognome'] = SociAttiviArray.sort(); // to make them alphabetically ordered regardless of their ID
    
      // SORT THE Progetti NAME BY THE DATE, SHOWING FIRST THE LAST OCCURRED
      var NumeroTotaleProgetti = choicesProgetti['Nome'].length;
      var ENDArray = [];
      var ProgettiArray = [];
      var Temp = [];
      var ProgettiSorted = {};
      var idx = 0;
      for (var i=0; i<NumeroTotaleProgetti; i++) {
        ProgettiEl = choicesProgetti['Nome'][i];
        NumberEl = choicesProgetti['Codice Progetto'][i];
        ENDArray[i] = [ProgettiEl,NumberEl];
      }
      ProgettiArray = ENDArray.reverse(function (a, b) { // sort in reverse order, first the last project
        return a[1] - b[1]; // the 2nd column are the project number code
      });
      for (var i=0; i<NumeroTotaleProgetti; i++) {
        Temp[i] = ProgettiArray[i][1] + " - " + ProgettiArray[i][0];
      }
      ProgettiSorted['ID & Nome Progetto'] = Temp;
    
      // APPLY THE ASSOCIATE NAMES AND THE EVENT NAMES TO THE FORM
      items.forEach(function(item) {
        const itemTitle = item.getTitle();
    
        // STUFF RELATED TO THE ASSOCIATE NAME
        if (itemTitle in SociAttivi) { // it was choicesSoci
          // to check if the item is in the list or not
          // if you have options of choices in that google sheet, move to the next step
          const itemType = item.getType(); // check which type the choice is because it's different 
          // for checkboxes, dropdown or whatever
          switch (itemType) {
            case FormApp.ItemType.CHECKBOX:
              item.asCheckboxItem().setChoiceValues(SociAttivi[itemTitle]); // it was choicesSoci[itemTitle]
              break;
            case FormApp.ItemType.LIST:
              item.asListItem().setChoiceValues(SociAttivi[itemTitle]);
              break;
            case FormApp.ItemType.MULTIPLE_CHOICE:
              item.asMultipleChoiceItem().setChoiceValues(SociAttivi[itemTitle]);
              break;
            default:
              Logger.log("Ignore question", itemTitle);
            }
          }
    
        // STUFF RELATED TO THE Progetti NAME
        if (itemTitle in ProgettiSorted) { // it was choicesEventi
        // to check if the item is in the list or not
        // using ProgettiSorted we have the freedom to use any combination of label instead of the one coming from Clickup or stored in the Supremo
        // attention that if you get easier the form, you complicate the appen analysis later!
        // if you have options of choices in that google sheet, move to the next step
        const itemType = item.getType(); // check which type the choice is because it's different 
        // for checkboxes, dropdown or whatever
        switch (itemType) {
          case FormApp.ItemType.CHECKBOX:
            item.asCheckboxItem().setChoiceValues(ProgettiSorted[itemTitle]); // it was choicesEventi[itemTitle]
            break;
          case FormApp.ItemType.LIST:
            item.asListItem().setChoiceValues(ProgettiSorted[itemTitle]);
            break;
          case FormApp.ItemType.MULTIPLE_CHOICE:
            item.asMultipleChoiceItem().setChoiceValues(ProgettiSorted[itemTitle]); 
            break;
          default:
            Logger.log("Ignore question", itemTitle);
          }
        }
      })
    
    }