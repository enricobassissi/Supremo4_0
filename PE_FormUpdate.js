// TIME TRIGGER TO UPDATE THE FORM
// PROGRAMMATICALLY WAY, INSTEAD OF GOING TO THE TRIGGER PAGE AND SET UP MANUALLY
// THE PROBLEM WITH THIS ONE IS THAT IT CREATE MANY AND MANY TRIGGERS!
/*
function timeDrivenPopulateFormTrigger_PE() { 
  ScriptApp
    .newTrigger("populateGoogleForm_PartecipazioneEventi")
    .timeBased()
    .everyHours(12)
    .create();
}
/*

/*
maybe a good idea to implement, one day
function checkIfTriggerExists(eventType, handlerFunction) {
var triggers = ScriptApp.getProjectTriggers();
var triggerExists = false;
triggers.forEach(function (trigger) {
  if(trigger.getEventType() === eventType &&
    trigger.getHandlerFunction() === handlerFunction)
    triggerExists = true;
});
return triggerExists;
}
*/

// GET DATA FROM THE SHEETS "EVENTI" AND "SOCI" TO POPULATE THE FORM
function getDataFromGoogleSheets_PartecipazioneEventi() {
    // video tutorial: https://www.youtube.com/watch?v=Z-gCwZ0lXd8
    
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetSoci = ss.getSheetByName("Soci");
      const sheetEventi = ss.getSheetByName("Eventi");
      const [headerSoci, ...dataSoci] = sheetSoci.getDataRange().getDisplayValues(); //take the same type of data as reported in the cell -> text
      const [headerEventi, ...dataEventi] = sheetEventi.getDataRange().getDisplayValues(); //take the same type of data as reported in the cell -> text
    
      const choicesSoci = {}
      const choicesEventi = {}
      headerSoci.forEach(
        function(title, index) {
          choicesSoci[title] = dataSoci.map(row => row[index]).filter(e => e !== "");
        }
      );
      headerEventi.forEach(
        function(title, index) {
          choicesEventi[title] = dataEventi.map(row => row[index]).filter(e => e !== "");
        }
      );
    
      return [choicesSoci, choicesEventi];
    
    }
    
    // ACTUALLY POPULATE THE FORM
    function populateGoogleForm_PartecipazioneEventi() {
    // video tutorial: https://www.youtube.com/watch?v=Z-gCwZ0lXd8
    
      var GoogleFormID = '1v8twZhCtasOhASK85r0Dk-_WlTdUBizWhsxzxEedUpE';
      var googleForm = FormApp.openById(GoogleFormID);
      var items = googleForm.getItems();
      var [choicesSoci, choicesEventi] = getDataFromGoogleSheets_PartecipazioneEventi();
    
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
      
      // SORT THE EVENT NAME BY THE DATE, SHOWING FIRST THE LAST OCCURRED
      var NumeroTotaleEventi = choicesEventi['Nome Evento'].length;
      var ENDArray = [];
      var EventiArray = [];
      var Temp = [];
      var EventiSorted = {};
      var idx = 0;
      for (var i=0; i<NumeroTotaleEventi; i++) {
        EventiEl = choicesEventi['Nome Evento'][i];
        DatesEl = choicesEventi['Data Value'][i];
        ENDArray[i] = [EventiEl,DatesEl];
      }
      EventiArray = ENDArray.reverse(function (a, b) { // sort in reverse order, first the last event
        return a[1] - b[1]; // the 2nd column are the dates
      });
      for (var i=0; i<NumeroTotaleEventi; i++) {
        Temp[i] = EventiArray[i][0];
      }
      EventiSorted['Nome Evento'] = Temp; 
      // same name as the name used in the form for naming the event name
      // best practice, keep them all the same in the supremo and in the form
    /*
    
      // APPLY THE ASSOCIATE NAMES AND THE EVENT NAMES TO THE FORM
    
      items[1].asCheckboxItem().setChoiceValues(SociAttivi['Nome e Cognome']); 
          // it was choicesSoci[itemTitle]
    
      // STUFF RELATED TO THE EVENT NAME
     
      
      items[0].asListItem().setChoiceValues(EventiSorted['Nome Evento']); //EventiSorted['Nome Evento']
    
        
      //debugger
     */
    
    
    
      // APPLY THE ASSOCIATE NAMES AND THE EVENT NAMES TO THE FORM
      items.forEach(function(item) {
        var itemTitle = item.getTitle();
        var aaa = 0;
        // STUFF RELATED TO THE ASSOCIATE NAME
        if (itemTitle in SociAttivi) { // it was choicesSoci
          aaa = 1;
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
          
        // STUFF RELATED TO THE EVENT NAME
        else if (itemTitle in EventiSorted) { // it was choicesEventi
        aaa = 2;
          // to check if the item is in the list or not
          // if you have options of choices in that google sheet, move to the next step
          const itemType = item.getType(); // check which type the choice is because it's different 
          // for checkboxes, dropdown or whatever
          switch (itemType) {
            case FormApp.ItemType.CHECKBOX:
              item.asCheckboxItem().setChoiceValues(EventiSorted[itemTitle]); // it was choicesEventi[itemTitle]
              break;
            case FormApp.ItemType.LIST:
              item.asListItem().setChoiceValues(EventiSorted[itemTitle]);
              break;
            case FormApp.ItemType.MULTIPLE_CHOICE:
              item.asMultipleChoiceItem().setChoiceValues(EventiSorted[itemTitle]); 
              break;
            default:
              Logger.log("Ignore question", itemTitle);
          }
        }
      });
    
    }