/********************
 * podklady pro vyúčtování závodu
 ********************/
function zavod(id, garant) {
  if (id) var eventId = id
  else generujVyuctovaniZavodu();
    
  // informace o závodu, nastavení jména záložky
  var infoZavod = getEventInfo(eventId);
  
  // rekurzivně zpracujeme všechny etapy do extra záložek
  //if (infoZavod.etapy.length > 0) {
  //  infoZavod.etapy.forEach(function(value, index, array) { zavod(value); });
  //  return;
  //}
  
  //vytvorime zalozku pro novy zavod
  var as = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = as.insertSheet();
  
  eventName = infoZavod.datum +"-" + infoZavod.name +" (" + id +")";
  sheet.setName(eventName);
  
  headerColor = "#dfe3ee";
  
  // záhlaví s informacemi o závodu
  sheet.appendRow([" "]);
  sheet.appendRow(["", "Garant závodu", garant]);
  sheet.appendRow(["", "Datum závodu", infoZavod.datum]);
  sheet.appendRow(["", "Název závodu", infoZavod.name]);
  sheet.appendRow(["", "Disciplina", infoZavod.sport + " - " + infoZavod.disciplina]);
  sheet.appendRow(["", "ORIS", constructOrisEventURL(id)]);
  sheet.getRange(sheet.getLastRow(), sheet.getLastColumn()).setShowHyperlink(true);
  sheet.appendRow([" "]);
  
  // formatujeme prvni sloupec, oddělvoací řádku a barvu
  sheet.getRange("B2:B6").setBackground(headerColor);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(3, 150);
  
  //vytvoříme záhlaví tabulky s informacemio startujících
  sheet.appendRow(["", "ID závodníka", "Jméno a příjmení", "Kategorie", "Termín přihlášky", "Startovné", "Startoval?", "Hradí klub?"])
  sheet.getRange(sheet.getLastRow(),2,1,7).setBackground(headerColor);
  
  // zjistime kdo z prihlasenych startoval
  var startovali = kdoStartoval(eventId);
  
  // zjistime kdo je prihlaseny, v jake kategorii atd.
  var url = 'https://oris.orientacnisporty.cz/API/?format=json&method=getEventEntries&eventid='+eventId+'&clubid=OPI';
  
  var j = UrlFetchApp.fetch(url).getContentText();
  var parsedEventInfo = JSON.parse(j);
  
  for (x in parsedEventInfo.Data)
  {
    var regNo = parsedEventInfo.Data[x].RegNo;
    var name = parsedEventInfo.Data[x].Name;
    var fee = parsedEventInfo.Data[x].Fee;
    var terminPrihlasky = parsedEventInfo.Data[x].EntryStop;
    var kategorie = parsedEventInfo.Data[x].ClassDesc;
    var bezel = startovali.bezel(regNo);
    var platiKlub = placenoKlubem(regNo, terminPrihlasky, kategorie, bezel, (infoZavod.etapy.length>0));
    values = ["", regNo, name, kategorie, terminPrihlasky, fee, bezel, platiKlub];
    sheet.appendRow(values);
  }
  
  // setřídíme podle kategorie
  sheet.getRange("B9:M"+sheet.getLastRow()).sort([{column: 4, ascending: true}]);
  // fixujeme záhlaví
  sheet.setFrozenRows(6);
}

/********************
 * informace o zavodu ve forme JSON objektu
 ********************/
function getEventInfo(eventId) {
  var url = 'https://oris.orientacnisporty.cz/API/?format=json&method=getEvent&id='+eventId;
  
  var json = UrlFetchApp.fetch(url).getContentText();
  var j = JSON.parse(json);
 
  var eventName = j.Data.Name
  var eventDate = new Date(j.Data.Date);
  var poradajiciOddil = j.Data.Org1.Abbr;
  var typZavodu = j.Data.Discipline.NameCZ;
  var typOB = j.Data.Sport.NameCZ;
  var noStages = new Number(j.Data.Stages);
  
  // etapovy zavod - potrebujeme predat vsechny etapy ke zpracovani
  var stages = [];
  for (i = 1; i <= noStages; i++) {
    Logger.log("Data.Stage"+i+"="+j.Data["Stage"+i]);
    stages.push(j.Data["Stage"+i]);
  }
  
  //extrahuj datum ve formatu dd.mm.yyyy
  var mesic = new Number(eventDate.getMonth())+1;
  var datumZavodu = eventDate.getDate()+"."+mesic+"."+eventDate.getYear();
  
  //priprav navratovy objekt
  var eventInfo = {id: eventId, name: eventName, oddil: poradajiciOddil, datum: datumZavodu, sport: typOB, disciplina: typZavodu, etapy: stages};
  
  return eventInfo;
}

/********************
 * seznam kdo startoval
 ********************/
function kdoStartoval(eventId) {
  var url = 'https://oris.orientacnisporty.cz/API/?format=json&method=getEventResults&eventid='+eventId+'&clubid=OPI';
  
  var json = UrlFetchApp.fetch(url).getContentText();
  var j = JSON.parse(json);
  var startovali = [];
 
  for (x in j.Data) {
    startovali.push(j.Data[x].RegNo);
  }
  
  rc = {bezeli: startovali, bezel : function(reg) { return (startovali.lastIndexOf(reg) > -1) ? "ANO" : "NE" }}
  
  return rc;
}


/********************
 * spocitani veku zavodnik
 ********************/
function getAge(regNo) {
  var s = "NA"
  if (regNo.length == 7 && regNo.lastIndexOf("OPI") == 0) {
     s = regNo.substring("OPI".length, "OPI".length+2);
     
     if (Number(s) < 30) s = "20"+s;
     else s = "19"+s;
  
     currentYear = (new Date()).getFullYear();
     s = (Number(currentYear) - Number(s))
  }
    
  return s;
}

/********************
 * ma zavodnik narok na proplaceni startovneho?
 ********************/
function placenoKlubem(regNo, terminPrihlasky, classDesc, bezel, etapovy) {
  vekZavodnika = getAge(regNo);
  s = "NE";
  
  //pokud je ve vekove kategorii 11 az 20 let
  if (bezel == "ANO" && vekZavodnika > 10 && vekZavodnika < 21) {
    if (etapovy) s = "ZKONTROLUJ ETAPY";
    else s = "ANO";
    // pokud startoval v P,T nebo faborkach nebo se prihlasil v druhem ci dalsim terminu pak zkontroluj zda ma narok
    if (classDesc.indexOf("P") > -1 || 
        classDesc.indexOf("T") > -1 || 
        classDesc.indexOf("F") > -1 || 
        classDesc.indexOf("N") > -1 ||
        classDesc.indexOf("HDR") > -1 ||
        classDesc.indexOf("10L") > -1 ||
        Number(terminPrihlasky) > 1) {
        s = "ASI NE"
    }
  }
    
  return s;
}


/********************
 * vytvori seznam zebrickovych zavodů na daný rok?
 ********************/
function zebrickoveZavody(rok) {
  if (rok == null) rok = (new Date()).getYear();
  
  var urlCelostatni = 'https://oris.orientacnisporty.cz/API/?format=json&datefrom='+rok+'-01-01&dateto='+rok+'-12-31&sport=1,2&level=1,2,3&method=getEventList';
  var urlJihoceske = 'https://oris.orientacnisporty.cz/API/?format=json&datefrom='+rok+'-01-01&dateto='+rok+'-12-31&sport=1&level=4&rg=JČ&method=getEventList';
  
  
  //vytvorime zalozku pro novy zavod
  var as = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = as.insertSheet("Závody "+rok);
  
  headerColor = "#dfe3ee";
  
  // záhlaví 
  sheet.appendRow(["", "Datum závodu","ORIS ID", "Sport", "Disciplina", "Název závodu","Soutěž", "Pořádá", "Garant", "Startovné celkem", "Startovné oddíl"]);
  
  // formatujeme prvni sloupec, oddělovací řádku a barvu
  sheet.getRange("B1:L1").setBackground(headerColor);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(6, 230);
  
  var j = UrlFetchApp.fetch(urlCelostatni).getContentText();
  var parsedEventInfo = JSON.parse(j);
  
  for (x in parsedEventInfo.Data)
  {
    var datumZavodu = parsedEventInfo.Data[x].Date;
    var idZavodu = parsedEventInfo.Data[x].ID;
    var sport = parsedEventInfo.Data[x].Sport.NameCZ;
    var disciplina = parsedEventInfo.Data[x].Discipline.NameCZ;
    var nazevZavodu = parsedEventInfo.Data[x].Name;
    var soutez = parsedEventInfo.Data[x].Level.ShortName;
    var poradajiciKlub = parsedEventInfo.Data[x].Org1.Abbr;
    
    values = ["", datumZavodu,idZavodu, sport, disciplina, nazevZavodu, soutez, poradajiciKlub];
    sheet.appendRow(values);
  }
  
  j = UrlFetchApp.fetch(urlJihoceske).getContentText();
  parsedEventInfo = JSON.parse(j);
  
  for (x in parsedEventInfo.Data)
  {
    var datumZavodu = parsedEventInfo.Data[x].Date;
    var idZavodu = parsedEventInfo.Data[x].ID;
    var sport = parsedEventInfo.Data[x].Sport.NameCZ;
    var disciplina = parsedEventInfo.Data[x].Discipline.NameCZ;
    var nazevZavodu = parsedEventInfo.Data[x].Name;
    var soutez = parsedEventInfo.Data[x].Level.ShortName;
    var poradajiciKlub = parsedEventInfo.Data[x].Org1.Abbr;
    
    values = ["", datumZavodu,idZavodu, sport, disciplina, nazevZavodu, soutez, poradajiciKlub];
    sheet.appendRow(values);
    //
    // tady dodělat pridání radku a nastaveni URL pres setFormula()....
    //
  }
  
  Logger.log(sheet.getLastRow());
  
  // setřídíme podle datumu
  sheet.getRange("B2:M"+sheet.getLastRow()).sort([{column: 2, ascending: true}]);
  // fixujeme záhlaví
  sheet.setFrozenRows(1);
  // formatujeme datum
  sheet.getRange("B2:B"+sheet.getLastRow()).setNumberFormat("dd.mm.yyyy");
}

/********************
 * rozšíření menu formuláře
 ********************/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OPI')
      .addItem('Generuj vyúčtování vybraného závodu ', 'generujVyuctovaniVybranehoZavodu')
      .addSeparator()
      .addItem('Generuj vyúčtování závodu ', 'generujVyuctovaniZavodu')
      .addItem('Generuj žebříčkové závody ', 'generujZebrickoveZavody')
      .addToUi();
}


/********************
 * UI dialog pro spuštění generování seznamu žebříčkových závodů
 ********************/
function generujZebrickoveZavody() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      'Generuj žebříčkové závody.',
      'Zadej rok (YYYY) pro který chceš seznam vytvořit:',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    zebrickoveZavody(text);
    ui.alert('Seznam závodů pro rok '+text+' je vytvořen.');
  } 
}


/********************
 * UI dialog pro generování vyúčtování závodu
 ********************/
function generujVyuctovaniZavodu() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      'Generuj vyúčtování závodu.',
      'Zadej ORIS ID závodu:',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    zavod(text);
    ui.alert('Podklady pro vyúčtování závodu č.'+text+' jsou vygenerovány.');
  } 
}

/********************
 * Spusť vyúčtování pro vybraný závod ze seznamu
 ********************/
function generujVyuctovaniVybranehoZavodu() {
  var ui = SpreadsheetApp.getUi();
  
  // předpokládáme že jsme na záložce se seznamem závodu
  var row = SpreadsheetApp.getActiveRange().getRowIndex();
  var id = SpreadsheetApp.getActiveSheet().getRange("C"+row).getValue();
  var nid = new Number(id);
  
  var garant = SpreadsheetApp.getActiveSheet().getRange("I"+row).getValue();
  
  if (id === null || isNaN(nid)) {
    ui.alert('ID závodu není k dispozici. \n\n Umistěte kurzor na řádek se závodem pro který chcete spustit generování vyúčtování.');
    return;
  }
  
  zavod(id, garant);
  
  ui.alert('Podklady pro vyúčtování závodu č.'+id+' jsou vygenerovány.');
   
}


function constructOrisEventURL(id) {
  return "https://oris.orientacnisporty.cz/Zavod?id="+id;
}