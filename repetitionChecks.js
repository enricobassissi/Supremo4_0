// CHECK TO CONTROL THAT THE NAME AND EVENT COUPLE YOU ARE ADDING WAS NOT ALREADY ENTERED ANOTHER TIME
function repetitionCheck_PE(nameSocio, nameEvento) {

    // return: true -> not repeated, false -> repeated
  
    // INTRODUCTION AND DECLARATION
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sPE = ss.getSheetByName("PartecipazioneEventi");
  
    // get all PartecipazioneEventi
    const [headerPE, ...dataPE] = sPE.getDataRange().getDisplayValues();
    const allPartecipantsEver = {}
    headerPE.forEach(
      function(title, index) {
        allPartecipantsEver[title] = dataPE.map(row => row[index]);
      }
    );
  
    // check che il nome che passi non è in quelli già presenti per l'evento per cui c'è già
    var esito = true;
    var countAll = allPartecipantsEver['Nome e Cognome'].length;
    for (var i=0; i<countAll; i++) { 
      // controlla che i nomi di allPartecipantsEver['xxx'] siano gli stessi dei campi del supremo 
      if (allPartecipantsEver['Nome e Cognome'][i] == nameSocio && allPartecipantsEver['Nome Evento'][i] == nameEvento) {
        esito = false;
      }
    }
  
    return esito
    
  }
  
  // CHECK TO CONTROL THAT THE NAME AND PROJECT ID COUPLE YOU ARE ADDING WAS NOT ALREADY ENTERED ANOTHER TIME
  // USEFUL ALSO IF THERE IS A CHANGE/ADDITION TO THE TEAM AND CLICKUP DOES IT AUTOMATICALLY
  function repetitionCheck_CP(name, projectID) {
  
    // output: true -> not repeated, false -> repeated
  
    // INTRODUCTION AND DECLARATION
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sCP = ss.getSheetByName("ComposizioneProgetti");
  
    // get all ComposizioneProgetti
    const [headerCP, ...dataCP] = sCP.getDataRange().getDisplayValues();
    const allProjectMembersEver = {}
    headerCP.forEach(
      function(title, index) {
        allProjectMembersEver[title] = dataCP.map(row => row[index]);
      }
    );
  
    // check che il nome che passi non è in quelli già presenti per il progetto per cui c'è già
    var esito = true;
    var countAll = allProjectMembersEver['Nome e Cognome'].length;
    for (var i=0; i<countAll; i++) { 
      // controlla che i nomi di allProjectMembersEver['xxx'] siano gli stessi dei campi del supremo 
      if (allProjectMembersEver['Nome e Cognome'][i] == name && allProjectMembersEver['Codice Progetto'][i] == projectID) {
        esito = false;
      }
    }
  
    return esito
    
  }
  