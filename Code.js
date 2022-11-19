//
// podklady pro vyúčtování závodu
//
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

  var sheetName = constructTabName(infoZavod);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  // ošetříme situaci kdy záložka existuje/neexistuje
  if (sheet == null) {
    //pokud ještě neexistuje pak vytvorime zalozku pro novy zavod
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    sheet.setName(sheetName);
  } else {
    //pokud existuje pak se zjitíme zda ji smazat a znovu vytvořit; pokud ne pak konec
    if (vymazatSheet(sheetName)) {
      sheet.clear();
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    } else return;
  }

  // záhlaví s informacemi o závodu
  sheet.appendRow([" "]);
  sheet.appendRow(["", "Garant závodu", garant, "", "Datum závodu", infoZavod.datum]);
  sheet.appendRow(["", "Název závodu", infoZavod.name, "", "Disciplina", infoZavod.sport + " - " + infoZavod.disciplina]);

  sheet.appendRow(["", "Přihlášky", "", "", "Doplňkové služby", ""]);

  let formula = "=HYPERLINK(\"" + constructOrisPrehledPrihlasenychURL(id) + "\";\"ORIS Přihlášky\")"
  sheet.getRange(sheet.getLastRow(), sheet.getLastColumn()-3).setFormula(formula);
  formula = "=HYPERLINK(\"" + constructOrisDoplnkoveSluzbyURL(id) + "\";\"ORIS Služby\")"
  sheet.getRange(sheet.getLastRow(), sheet.getLastColumn()).setFormula(formula);

  const headerColor = "#dfe3ee";

  // formatujeme prvni sloupec, oddělovací řádku a barvu
  sheet.getRange("B2:B4").setBackground(headerColor);
  sheet.getRange("E2:E4").setBackground(headerColor);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(3, 150);

  //vytvoříme záhlaví tabulky s informacemi o startujících
  sheet.appendRow(["", "ID závodníka", "Jméno a příjmení", "Kategorie", "Termín přihlášky", "Startovné", "Startoval?", "Hradí startovné klub?", "Doplňkové služby"])
  sheet.getRange(sheet.getLastRow(), 2, 1, 8).setBackground(headerColor);

  // zjistime kdo z prihlasenych startoval
  const startovali = kdoStartoval(eventId);

  // zjistime kdo mel objednané doplnkove sluzby
  const sluzby = doplnkoveSluzby(eventId);

  // zjistime kdo je prihlaseny, v jake kategorii atd.
  const url = 'https://oris.orientacnisporty.cz/API/?format=json&method=getEventEntries&eventid=' + eventId + '&clubid=OPI';

  const j = UrlFetchApp.fetch(url).getContentText();
  const parsedEventInfo = JSON.parse(j);

  // vypíšeme jednotlivé závodníky ...
  // .. pozor na štafety ST
  for (x in parsedEventInfo.Data) {
    let regNo = (infoZavod.disciplinaZkratka == "ST") ? "N/A" : parsedEventInfo.Data[x].RegNo;
    let name = parsedEventInfo.Data[x].Name;
    let fee = parsedEventInfo.Data[x].Fee;
    let terminPrihlasky = parsedEventInfo.Data[x].EntryStop;
    let kategorie = parsedEventInfo.Data[x].ClassDesc;
    let bezel = (infoZavod.disciplinaZkratka == "ST") ? "N/A" : startovali.bezel(regNo);
    let platiKlub = (infoZavod.disciplinaZkratka == "ST") ? "ANO" : placenoKlubem(regNo, terminPrihlasky, kategorie, bezel, (infoZavod.etapy.length > 0));
    let spaniJidlo = (infoZavod.disciplinaZkratka == "ST") ? "Koukni do ORISu" : sluzby.kolik(regNo);
    let values = ["", regNo, name, kategorie, terminPrihlasky, fee, bezel, platiKlub, spaniJidlo];
    sheet.appendRow(values);
  }

  // setřídíme podle kategorie
  const lastRowIndex = sheet.getLastRow();
  sheet.getRange("B9:M" + lastRowIndex).sort([{ column: 4, ascending: true }]);
  // formatujeme CZK
  sheet.getRange("F2:F" + lastRowIndex).setNumberFormat("#,##0.00 [$Kč]");
  sheet.getRange("I2:I" + lastRowIndex).setNumberFormat("#,##0.00 [$Kč]");

  // fixujeme záhlaví
  sheet.setFrozenRows(5);

  sheet.autoResizeColumns(2, 8);

  return sheet.getSheetId();
}

//
// informace o zavodu ve forme JSON objektu
 //
function getEventInfo(eventId) {
  const url = 'https://oris.orientacnisporty.cz/API/?format=json&method=getEvent&id=' + eventId;

  const json = UrlFetchApp.fetch(url).getContentText();
  const j = JSON.parse(json);

  const eventName = j.Data.Name
  const eventDate = new Date(j.Data.Date);
  const poradajiciOddil = j.Data.Org1.Abbr;
  const typZavodu = j.Data.Discipline.NameCZ;
  const typZavoduShort = j.Data.Discipline.ShortName;
  const typOB = j.Data.Sport.NameCZ;
  const noStages = new Number(j.Data.Stages);

  // etapovy zavod - potrebujeme predat vsechny etapy ke zpracovani
  const stages = [];
  for (i = 1; i <= noStages; i++) {
    stages.push(j.Data["Stage" + i]);
  }

  //extrahuj datum ve formatu dd.mm.yyyy
  const mesic = new Number(eventDate.getMonth()) + 1;
  const datumZavodu = eventDate.getDate() + "." + mesic + "." + eventDate.getYear();

  //priprav navratovy objekt
  const eventInfo = { id: eventId, name: eventName, oddil: poradajiciOddil, datum: datumZavodu, sport: typOB, disciplina: typZavodu, disciplinaZkratka: typZavoduShort, etapy: stages };

  return eventInfo;
}

//
// seznam kdo startoval
 //
function kdoStartoval(eventId) {
  const url = 'https://oris.orientacnisporty.cz/API/?format=json&method=getEventResults&eventid=' + eventId + '&clubid=OPI';

  var json = UrlFetchApp.fetch(url).getContentText();
  var j = JSON.parse(json);
  var startovali = [];

  for (x in j.Data) {
    startovali.push(j.Data[x].RegNo);
  }

  rc = { bezeli: startovali, bezel: function (reg) { return (startovali.lastIndexOf(reg) > -1) ? "ANO" : "NE" } }

  return rc;
}

//
// doplňkové služby - ubytování, jídlo, ....
 //
function doplnkoveSluzby(eventId) {
  const url = 'https://oris.orientacnisporty.cz/API/?format=json&method=getEventServiceEntries&eventid=' + eventId + '&clubid=OPI';

  var json = UrlFetchApp.fetch(url).getContentText();
  var j = JSON.parse(json);
  var sluzby = [];

  for (x in j.Data) {
    sluzby.push([j.Data[x].RegNo, j.Data[x].TotalFee]);
  }

  var rc = {
    doplnkoveSluzby: sluzby,
    // pro regNo zavodnika posčítá doplňkové služby
    kolik: function (reg) {
      soucet = 0;
      this.doplnkoveSluzby.forEach(function (value, index, array) {
        if (value[0] == reg) soucet = soucet + Number(value[0, 1]);
      });
      return soucet;
    }
  }

  return rc;
}


//
 //spocitani veku zavodnik na základě jeho ORIS registračního ID
 // např. OPI6900 - kde 69 je rok naorzení
 //
function getAge(regNo) {
  let s = "NA"
  if (regNo != null && regNo.length == 7 && regNo.lastIndexOf("OPI") == 0) {
    s = regNo.substring("OPI".length, "OPI".length + 2);

    if (Number(s) < 30) s = "20" + s;
    else s = "19" + s;

    currentYear = (new Date()).getFullYear();
    s = (Number(currentYear) - Number(s))
  }

  return s;
}

//
// ma zavodnik narok na proplaceni startovneho?
// ... podle aktuálních provozních pravidel klubu
//
function placenoKlubem(regNo, terminPrihlasky, classDesc, bezel, etapovy) {
  const vekZavodnika = getAge(regNo);
  let s = "NE";

  if (bezel == "NE" && vekZavodnika > 10 && vekZavodnika < 21) {
      s = "MOŽNÁ, prověř důvod proč neběžel?"
  }
  
  //pokud je ve vekove kategorii 11 az 20 let
  if (bezel == "ANO" && vekZavodnika < 21) {
    if (etapovy) s = "ZKONTROLUJ ETAPY";
    else s = "ANO";
    // pokud startoval v P,T nebo faborkach nebo se prihlasil v druhem 
    // ci dalsim terminu pak zkontroluj v ORISu zda ma narok
    if (vekZavodnika > 10 && (classDesc.indexOf("P") > -1 ||
      classDesc.indexOf("T") > -1 ||
      classDesc.indexOf("F") > -1 ||
      classDesc.indexOf("N") > -1 ||
      classDesc.indexOf("HDR") > -1 ||
      classDesc.indexOf("10L") > -1)) {
      s = "PROVĚŘ"
    }
    if (Number(terminPrihlasky) > 1) {
      s = "NE, 2. termín přihlášky"
    }
  }

  return s;
}


//
// vytvori seznam zebrickovych zavodů na daný rok?
 //
function zebrickoveZavody(rok) {
  if (rok == null) rok = (new Date()).getYear();

  const urlCelostatni = 'https://oris.orientacnisporty.cz/API/?format=json&datefrom=' + rok + '-01-01&dateto=' + rok + '-12-31&sport=1,2&level=1,2,3&rg=Č&method=getEventList';
  const urlMCR = 'https://oris.orientacnisporty.cz/API/?format=json&datefrom=' + rok + '-01-01&dateto=' + rok + '-12-31&sport=1,2&level=1,2,3&rg=ČR&method=getEventList';
  const urlJihoceske = 'https://oris.orientacnisporty.cz/API/?format=json&datefrom=' + rok + '-01-01&dateto=' + rok + '-12-31&sport=1&level=4,11&rg=JČ&method=getEventList';

  //vytvorime zalozku pro novy zavod
  const as = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = as.insertSheet("Závody " + rok);

  const headerColor = "#dfe3ee";

  // záhlaví 
  sheet.appendRow(["", "Datum závodu", "ORIS ID", "Sport", "Disciplina", "Název závodu", "Soutěž", "Pořádá", "Garant", "Startovné celkem", "Startovné oddíl", "Uzavřeno"]);

  // formatujeme prvni sloupec, oddělovací řádku a barvu
  sheet.getRange("B1:L1").setBackground(headerColor);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(6, 230);

  populateZebrickoveZavody(urlCelostatni, sheet);
  populateZebrickoveZavody(urlMCR, sheet);
  populateZebrickoveZavody(urlJihoceske, sheet);

  // setřídíme podle datumu
  sheet.getRange("B2:M" + sheet.getLastRow()).sort([{ column: 2, ascending: true }]);
  // fixujeme záhlaví
  sheet.setFrozenRows(1);
  // formatujeme datum
  sheet.getRange("B2:B" + sheet.getLastRow()).setNumberFormat("dd.mm.yyyy");
}

//
// Plní záložku záznamy žebříčkových závodů
//
function populateZebrickoveZavody(url, sheet) {
  const j = UrlFetchApp.fetch(url).getContentText();
  const parsedEventInfo = JSON.parse(j);

  for (x in parsedEventInfo.Data) {
    const datumZavodu = parsedEventInfo.Data[x].Date;
    const idZavodu = parsedEventInfo.Data[x].ID;
    const sport = parsedEventInfo.Data[x].Sport.NameCZ;
    const disciplina = parsedEventInfo.Data[x].Discipline.NameCZ;
    const nazevZavodu = parsedEventInfo.Data[x].Name;
    const soutez = parsedEventInfo.Data[x].Level.ShortName;
    const poradajiciKlub = parsedEventInfo.Data[x].Org1.Abbr;

    const values = ["", datumZavodu, idZavodu, sport, disciplina, nazevZavodu, soutez, poradajiciKlub];
    sheet.appendRow(values);
  }
}

//
// nastavení aktivní záložky (první záložka) a rozšíření menu formuláře
//
function onOpen() {

  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheets[0]);

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('OPI')
    .addItem('Generuj vyúčtování vybraného závodu ', 'generujVyuctovaniVybranehoZavodu')
    .addSeparator()
    .addItem('Generuj vyúčtování závodu ', 'generujVyuctovaniZavodu')
    .addItem('Generuj žebříčkové závody ', 'generujZebrickoveZavody')
    .addToUi();
}


//
// UI dialog pro spuštění generování seznamu žebříčkových závodů
//
function generujZebrickoveZavody() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    'Generuj žebříčkové závody.',
    'Zadej rok (YYYY) pro který chceš seznam vytvořit:',
    ui.ButtonSet.OK_CANCEL);

  const button = result.getSelectedButton();
  const text = result.getResponseText();
  if (button == ui.Button.OK) {
    zebrickoveZavody(text);
    ui.alert('Seznam závodů pro rok ' + text + ' je vytvořen.');
  }
}


//
// UI dialog pro generování vyúčtování závodu
//
function generujVyuctovaniZavodu() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    'Generuj vyúčtování závodu.',
    'Zadej ORIS ID závodu:',
    ui.ButtonSet.OK_CANCEL);

  const button = result.getSelectedButton();
  const text = result.getResponseText();
  if (button == ui.Button.OK) {
    zavod(text);
    ui.alert('Podklady pro vyúčtování závodu č.' + text + ' jsou vygenerovány.');
  }
}

//
// Spusť vyúčtování pro vybraný závod ze seznamu
 //
function generujVyuctovaniVybranehoZavodu() {
  const ui = SpreadsheetApp.getUi();

  // předpokládáme že jsme na záložce se seznamem závodu
  const zebrickoveZavody = SpreadsheetApp.getActiveSheet();
  const row = zebrickoveZavody.getActiveRange().getRowIndex();
  const id = zebrickoveZavody.getRange("C" + row).getValue();
  const nazevZavodu = zebrickoveZavody.getRange("F" + row).getValue();
  const garant = zebrickoveZavody.getRange("I" + row).getValue();

  var eventId = new Number(id);
  if (id === null || isNaN(eventId)) {
    ui.alert('ID závodu není k dispozici. \n\n Umistěte kurzor na řádek se závodem pro který chcete spustit generování vyúčtování.');
    return;
  }

  // informace o závodu
  const infoZavod = getEventInfo(eventId);

  const sheetId = zavod(eventId, garant);

  if (sheetId != null || sheetId != undefined) {
    //nastavime link na záložku s vyúčtováním
    zebrickoveZavody.getRange("F" + row).setFormula("=HYPERLINK(\"#gid=" + sheetId + "\";\"" + nazevZavodu + "\")");
    ui.alert('Podklady pro vyúčtování závodu č.' + eventId + ' jsou vygenerovány.');
  }
}

//
// UI dialog pro potvrzení zda přegenrovat již existující záložku
//
function vymazatSheet(sheetName) {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'Záložka \"' + sheetName + '\" již existuje.',
    'Chcete ji znovu naplnit?',
    ui.ButtonSet.YES_NO);

  return (result == ui.Button.YES);
}

//
// Konstruuje ORIS URL závodu 
//
function constructOrisEventURL(id) {
  return "https://oris.orientacnisporty.cz/Zavod?id=" + id;
}

//
// Konstruuje ORIS URL pro doplňkové služby OPI
//
function constructOrisDoplnkoveSluzbyURL(id) {
  return "https://oris.orientacnisporty.cz/DoplnkoveSluzby?id=" + id + "#105";
}

//
// Konstruuje ORIS URL pro přehled přihlášených za OPI
//
function constructOrisPrehledPrihlasenychURL(id) {
  return "https://oris.orientacnisporty.cz/PrehledPrihlasenych?id=" + id + "&mode=clubs#105"
}

//
// Konstruuje jméno záložky (mělo by být unikátní)
//
function constructTabName(eventInfo) {
  const eventName = eventInfo.datum + "-" + eventInfo.name + " (" + eventInfo.id + ")";
  return eventName;
}