interface MapS2I {
  [others: string]: number;
}
interface MapS2S {
  [others: string]: string;
}
interface HeaderMap {
  [others: string]: MapS2I;
}

let inited = false;
let headers: HeaderMap = {};
let reisenSheet: GoogleAppsScript.Spreadsheet.Sheet;
let buchungenSheet: GoogleAppsScript.Spreadsheet.Sheet;

// Indices are 1-based!!
let mailIndex: number; // E-Mail-Adresse
let herrFrauIndex: number; // Anrede
let herrFrau1Index: number; // Anrede 1
let nameIndex: number; // Name
let name1Index: number; // Name 1
let zustimmungsIndex: number; // Zustimmung zur SEPA-Lastschrift
let bestätigungsIndex: number; // Bestätigung (der Teilnahmebedingungen)
let verifikationsIndex: number; // Verifikation (der Email-Adresse)
let anmeldebestIndex: number; // Anmeldebestätigung (gesendet)
let tourIndexB: number; // Bei welchen Touren möchten Sie mitfahren?
let einzelnIndex: number; // Reisen Sie alleine oder zu zweit?

const tourFrage = "Bei welchen Touren möchten Sie mitfahren?";
const einzelnFrage = "Reisen Sie alleine oder zu zweit?";

interface Event {
  namedValues: { [others: string]: string[] };
  range: GoogleAppsScript.Spreadsheet.Range;
  [others: string]: any;
}

function isEmpty(str: string | undefined) {
  return !str || 0 === str.length;
}

function test() {
  init();
  let e: Event = {
    namedValues: {
      "Vorname 1": ["Michael"],
      "Name 1": ["Uhlenberg"],
      "Anrede 1": ["Herr"],
      "Vorname 2": ["Antonia"],
      "Name 2": ["Ruhdorfer"],
      "Anrede 2": ["Frau"],
      "E-Mail-Adresse": ["michael.uhlenberg@t-online.de"],
      "Gleiche Adresse wie Teilnehmer 1 ?": ["ja"],
      "IBAN-Kontonummer": ["DE91100000000123456789"],
    },
    range: buchungenSheet.getRange(2, 1, 1, buchungenSheet.getLastColumn()),
  };
  e.namedValues[einzelnFrage] = ["Zu zweit (DZ)"];
  e.namedValues[tourFrage] = [
    // "Fahrradtour um den Gardasee vom 1.5. bis 12.5.",
    // "Transalp von Salzburg nach Venedig vom 2.5. bis 13.5.",
    "Das malerische Havelland entdecken vom 5.5. bis 16.5.",
    // Durch die ungarische Steppe vom 4.5. bis 15.5.",
    //"Entlang der Drau vom 3.5. bis 14.5",
  ];
  dispatch(e);
}

function init() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  for (let sheet of sheets) {
    let sheetName = sheet.getName();
    let sheetHeaders: MapS2I = {};
    // Logger.log("sheetName %s", sheetName);
    headers[sheetName] = sheetHeaders;
    let numCols = sheet.getLastColumn();
    // Logger.log("numCols %s", numCols);
    let row1Vals = sheet.getRange(1, 1, 1, numCols).getValues();
    // Logger.log("sheetName %s row1 %s", sheetName, row1Vals);
    for (let i = 0; i < numCols; i++) {
      let v: string = row1Vals[0][i];
      if (isEmpty(v)) continue;
      sheetHeaders[v] = i + 1;
    }
    // Logger.log("sheet %s %s", sheetName, sheetHeaders);

    if (sheet.getName() == "Reisen") {
      reisenSheet = sheet;
    }
    if (sheet.getName() == "Buchungen") {
      buchungenSheet = sheet;
      mailIndex = sheetHeaders["E-Mail-Adresse"];
      herrFrauIndex = sheetHeaders["Anrede"];
      herrFrau1Index = sheetHeaders["Anrede 1"];
      nameIndex = sheetHeaders["Name"];
      name1Index = sheetHeaders["Name 1"];
      bestätigungsIndex = sheetHeaders["Bestätigung"];
      verifikationsIndex = sheetHeaders["Verifikation"];
      tourIndexB = sheetHeaders[tourFrage];
      einzelnIndex = sheetHeaders[einzelnFrage];
      if (verifikationsIndex == null) {
        verifikationsIndex = addColumn(sheet, sheetHeaders, "Verifikation");
      }
      anmeldebestIndex = sheetHeaders["Anmeldebestätigung"];
      if (anmeldebestIndex == null) {
        anmeldebestIndex = addColumn(sheet, sheetHeaders, "Anmeldebestätigung");
      }
    }
    inited = true;
  }
}

// add a cell in row 1 with a new column title, return its index
function addColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetHeaders: MapS2I,
  title: string
): number {
  let max = 0;
  for (let sh in sheetHeaders) {
    if (sheetHeaders[sh] > max) max = sheetHeaders[sh];
  }
  if (max >= sheet.getMaxColumns()) {
    sheet.insertColumnAfter(max);
  }
  max += 1;
  sheet.getRange(1, max).setValue(title);
  sheetHeaders[title] = max;
  return max;
}

function anredeText(herrFrau: string, name: string) {
  if (herrFrau === "Herr") {
    return "Sehr geehrter Herr " + name;
  } else {
    return "Sehr geehrte Frau " + name;
  }
}

function heuteString() {
  return Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    "YYYY-MM-dd HH:mm:ss"
  );
}

function attachmentFiles() {
  let thisFileId = SpreadsheetApp.getActive().getId();
  let thisFile = DriveApp.getFileById(thisFileId);
  let parent = thisFile.getParents().next();
  let grandPa = parent.getParents().next();
  let attachmentFolder = grandPa
    .getFoldersByName("Anhänge für Anmelde-Bestätigung")
    .next();
  let PDFs = attachmentFolder.getFilesByType("application/pdf"); // MimeType.PDF
  let files = [];
  while (PDFs.hasNext()) {
    files.push(PDFs.next());
  }
  return files; // why not use PDFs directly??
}

function tourPreis(einzeln: boolean, reise: string) {
  let reisenRows = reisenSheet.getLastRow() - 1; // first row = headers
  let reisenCols = reisenSheet.getLastColumn();
  let reisenVals = reisenSheet
    .getRange(2, 1, reisenRows, reisenCols)
    .getValues();
  let reisenNotes = reisenSheet
    .getRange(2, 1, reisenRows, reisenCols)
    .getNotes();
  let tourIndexR = headers["Reisen"]["Reise"];
  let dzPreisIndex = headers["Reisen"]["DZ-Preis"];
  let ezPreisIndex = headers["Reisen"]["EZ-Preis"];

  let betrag = 0;
  for (let i = 0; i < reisenRows; i++) {
    if (!isEmpty(reisenNotes[i][0])) continue;
    if (reisenVals[i][tourIndexR - 1] === reise) {
      betrag = einzeln
        ? reisenVals[i][ezPreisIndex - 1]
        : 2 * reisenVals[i][dzPreisIndex - 1];
      return betrag;
    }
  }
  return 0;
}

function anmeldebestätigung() {
  if (!inited) init();
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Buchungen") {
    SpreadsheetApp.getUi().alert(
      "Bitte eine Zeile im Sheet 'Buchungen' selektieren"
    );
    return;
  }
  let curCell = sheet.getSelection().getCurrentCell();
  if (!curCell) {
    SpreadsheetApp.getUi().alert("Bitte zuerst Teilnehmerzeile selektieren");
    return;
  }
  let row = curCell.getRow();
  if (row < 2 || row > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Teilnehmerzeile selektieren"
    );
    return;
  }
  let rowValues = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  let rowNote = sheet.getRange(row, 1).getNote();
  if (!isEmpty(rowNote)) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile hat eine Notiz und ist deshalb ungültig"
    );
    return;
  }
  if (isEmpty(rowValues[verifikationsIndex])) {
    SpreadsheetApp.getUi().alert("Email-Adresse nicht verifiziert");
    return;
  }
  if (!isEmpty(rowValues[anmeldebestIndex - 1])) {
    SpreadsheetApp.getUi().alert("Die Reise wurde schon bestätigt");
    return;
  }
  // setting up mail
  let emailTo: string = rowValues[mailIndex - 1];
  let subject: string = "Bestätigung Ihrer Buchung";
  let einzeln: boolean = rowValues[einzelnIndex - 1].startsWith("Allein");
  let herrFrau = einzeln
    ? rowValues[herrFrauIndex - 1]
    : rowValues[herrFrau1Index - 1];
  let name = einzeln ? rowValues[nameIndex - 1] : rowValues[name1Index - 1];
  // Anrede
  let anrede: string = anredeText(herrFrau, name);
  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    "emailBestätigung.html"
  );

  let reise: string = rowValues[tourIndexB - 1];

  let betrag: number = tourPreis(einzeln, reise);
  template.anrede = anrede;
  template.reise =
    reise + (einzeln ? " für eine Person" : " für zwei Personen");
  template.betrag = betrag;

  SpreadsheetApp.getUi().alert(
    herrFrau +
      " " +
      name +
      " bucht für die Reise '" +
      reise +
      "' ein " +
      (einzeln ? "EZ" : "DZ") +
      " für " +
      betrag +
      "Euro"
  );

  let htmlText: string = template.evaluate().getContent();
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "Mehrtagestouren ADFC München e.V.",
    replyTo: "michael.uhlenberg@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
  // update sheet
  sheet.getRange(row, anmeldebestIndex).setValue(heuteString());
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("ADFC-MTT")
    .addItem("Anmeldebestätigung senden", "anmeldebestätigung")
    .addItem("Update", "update")
    // .addItem("Test", "test")
    .addToUi();
}

function dispatch(e: Event) {
  if (!inited) init();
  // let keys = Object.keys(e);
  // Logger.log("verif2", keys);
  // for (let key of keys) {
  //   Logger.log("key %s val %s keys %s", key, e[key], Object.keys(e[key]));
  // }
  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  Logger.log("dispatch sheet", sheet.getName(), range.getA1Notation());
  if (sheet.getName() == "Test") checkBuchung(e);
  if (sheet.getName() == "Buchungen") checkBuchung(e);
  if (sheet.getName() == "Email-Verifikation") verifyEmail();
}

function verifyEmail() {
  let ssheet = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ssheet.getSheetByName("Email-Verifikation");
  if (evSheet.getLastRow() < 2) return;
  // It is a big nuisance that getSheetValues with a row count of 0 throws an error, instead of returning an empty list.
  let evalues = evSheet.getSheetValues(
    2,
    1,
    evSheet.getLastRow() - 1,
    evSheet.getLastColumn()
  ); // Mit dieser Email-Adresse

  let numRows = buchungenSheet.getLastRow();
  if (numRows < 2) return;
  let bvalues = buchungenSheet.getSheetValues(
    2,
    1,
    numRows - 1,
    buchungenSheet.getLastColumn()
  );
  Logger.log("bvalues %s", bvalues);

  for (let bx in bvalues) {
    let bxi = +bx; // confusingly, bx is initially a string, and is interpreted as A1Notation in sheet.getRange(bx) !
    let brow = bvalues[bxi];
    if (
      !isEmpty(brow[mailIndex - 1]) &&
      isEmpty(brow[verifikationsIndex - 1])
    ) {
      let baddr = brow[1];
      for (let ex in evalues) {
        let erow = evalues[ex];
        if (erow[1] != "Ja" || isEmpty(erow[2])) continue;
        let eaddr = erow[2];
        if (eaddr != baddr) continue;
        // Buchungen[Verifiziert] = Email-Verif[Zeitstempel]
        buchungenSheet.getRange(bxi + 2, verifikationsIndex).setValue(erow[0]);
        brow[verifikationsIndex - 1] = erow[0];
        break;
      }
    }
  }
}

function checkBuchung(e: Event) {
  let keys = Object.keys(e);
  Logger.log("checkBuch", keys, typeof e);
  for (let key of keys) {
    Logger.log("key %s val %s", key, e[key]);
  }

  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  let row = range.getRow();
  let cellA = range.getCell(1, 1);
  Logger.log("sheet %s row %s cellA %s", sheet, row, cellA.getA1Notation());

  let ibanNV = e.namedValues["IBAN-Kontonummer"][0];
  let iban = ibanNV.replace(/\s/g, "").toUpperCase();
  let emailTo = e.namedValues["E-Mail-Adresse"][0];
  Logger.log("iban=%s emailTo=%s %s", iban, emailTo, typeof emailTo);
  if (!isValidIban(iban)) {
    sendWrongIbanEmail(emailTo, iban);
    cellA.setNote("Ungültige IBAN");
    return;
  }
  if (iban != ibanNV) {
    let cellIban = range.getCell(1, headers["Buchungen"]["IBAN-Kontonummer"]);
    cellIban.setValue(iban);
  }
  // Die Zellen Zustimmung und Bestätigung sind im Formular als Pflichtantwort eingetragen
  // und können garnicht anders als gesetzt sein. Sonst hier prüfen analog zu IBAN.

  let einzeln = e.namedValues[einzelnFrage][0].startsWith("Alleine");
  let personen = einzeln ? "einer Person" : "zwei Personen";
  let restCol = einzeln
    ? headers["Reisen"]["EZ-Rest"]
    : headers["Reisen"]["DZ-Rest"];
  let touren: Array<string> = [];
  let tourenNV: Array<string> = e.namedValues[tourFrage];
  // actually, tourenNV = e.g. ["tour1, tour2, tour3"], i.e. just one element
  for (let tour of tourenNV) {
    let tourenList = tour.split(","); // that's the reason why we don't want commas within a tour title!
    for (let tlItem of tourenList) {
      touren.push(tlItem.trim());
    }
  }
  // Logger.log("check2", einzeln, personen, restCol, touren, touren.length);
  if (touren.length == 0) {
    Logger.log("check4");
    // cannot happen, answer is mandatory
    return;
  }

  // for each tour besides the first create a new row, and put the tour into it
  let buchungsRowNumbers = [row];
  if (touren.length > 1) {
    let numCols = sheet.getLastColumn();
    let tourCellNo = headers["Buchungen"][tourFrage];
    for (let i = 1; i < touren.length; i++) {
      let toRow = sheet.getLastRow() + 1;
      if (toRow >= sheet.getMaxRows()) {
        sheet.insertRowAfter(toRow);
      }
      let toRange = sheet.getRange(toRow, 1, 1, numCols);
      range.copyTo(toRange);
      let tourCell = toRange.getCell(1, tourCellNo);
      tourCell.setValue(touren[i]);
      buchungsRowNumbers.push(toRow);
    }
    // put the first tour into the original row
    let tourCell = range.getCell(1, tourCellNo);
    tourCell.setValue(touren[0]);
  }

  let msgs = [];
  let reisen: Array<Array<string>> = reisenSheet.getSheetValues(
    2,
    1,
    reisenSheet.getLastRow(),
    reisenSheet.getLastColumn()
  );
  let restChanged = false;
  for (let i = 0; i < touren.length; i++) {
    let tourFound = false;
    for (let j = 0; j < reisen.length; j++) {
      if (reisen[j][0] == touren[i]) {
        tourFound = true;
        let rest = reisenSheet.getRange(2 + j, restCol).getValue();
        if (rest <= 0) {
          msgs.push(
            "Die Reise '" +
              touren[i] +
              "' für " +
              personen +
              " ist leider ausgebucht."
          );
          sheet.getRange(buchungsRowNumbers[i], 1).setNote("Ausgebucht");
        } else {
          msgs.push(
            "Sie sind für die Reise '" +
              touren[i] +
              "' für " +
              personen +
              " vorgemerkt."
          );
          reisenSheet.getRange(2 + j, restCol).setValue(rest - 1);
          restChanged = true;
        }
        break;
      }
    }
    if (!tourFound) {
      Logger.log("tour '" + touren[i] + " nicht im Reisen-Sheet!?");
    }
  }
  if (msgs.length == 0) {
    Logger.log("keine Touren gefunden!?");
    return;
  }
  if (restChanged) {
    updateForm();
  }
  Logger.log("msgs: ", msgs, msgs.length);
  sendeAntwort(e, msgs, sheet, buchungsRowNumbers);
}

function sendeAntwort(
  e: Event,
  msgs: Array<string>,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  buchungsRowNumbers: Array<number>
) {
  let emailTo = e.namedValues["E-Mail-Adresse"][0];
  Logger.log("emailTo=" + emailTo);

  let templateFile = "emailVerif.html";

  // do we already know this email address?
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ss.getSheetByName("Email-Verifikation");
  let numRows = evSheet.getLastRow();
  let evalues =
    numRows < 2
      ? []
      : evSheet.getSheetValues(2, 1, evSheet.getLastRow() - 1, 3);
  for (let i = 0; i < evalues.length; i++) {
    // Mit dieser Email-Adresse
    if (evalues[i][2] == emailTo) {
      templateFile = "emailReply.html"; // yes, don't ask for verification
      for (let j = 0; j < buchungsRowNumbers.length; j++) {
        sheet
          .getRange(buchungsRowNumbers[j], verifikationsIndex)
          .setValue(evalues[i][0]);
      }
      break;
    }
  }

  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    templateFile
  );
  template.anrede = anrede(e);
  template.msgs = msgs;
  template.verifLink =
    "https://docs.google.com/forms/d/e/1FAIpQLScLF2jogdsQGOI_A4gvGVvmrasN6pS5MZgY7xvqSjMB87F6uw/viewform?usp=pp_url&entry.1398709542=Ja&entry.576071197=" +
    encodeURIComponent(emailTo);

  let htmlText: string = template.evaluate().getContent();
  let subject =
    templateFile === "emailVerif.html"
      ? "Bestätigung Ihrer Email-Adresse"
      : "Bestätigung Ihrer Anmeldung";
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "Mehrtagestouren ADFC München e.V.",
    replyTo: "michael.uhlenberg@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
}

function anrede(e: Event) {
  // if Name is not set, nv["Name"] has value [""], i.e. not null, not [], not [null]!
  let anredeA: Array<string> = e.namedValues["Anrede"];
  if (anredeA == null || anredeA.length == 0 || isEmpty(anredeA[0])) {
    anredeA = e.namedValues["Anrede 1"];
  }
  let anrede: string = anredeA[0];

  let vornameA: Array<string> = e.namedValues["Vorname"];
  if (vornameA == null || vornameA.length == 0 || isEmpty(vornameA[0])) {
    vornameA = e.namedValues["Vorname 1"];
  }
  let vorname: string = vornameA[0];

  let nameA: Array<string> = e.namedValues["Name"];
  if (nameA == null || nameA.length == 0 || isEmpty(nameA[0])) {
    nameA = e.namedValues["Name 1"];
  }
  let name: string = nameA[0];

  if (anrede == "Herr") {
    anrede = "Sehr geehrter Herr ";
  } else {
    anrede = "Sehr geehrte Frau ";
  }
  Logger.log("anrede", anrede, vorname, name);
  return anrede + vorname + " " + name;
}

function update() {
  if (!inited) init();
  verifyEmail();
  updateZimmerReste();
  updateForm();
}

function updateZimmerReste() {
  let reisenRows = reisenSheet.getLastRow() - 1; // first row = headers
  let reisenCols = reisenSheet.getLastColumn();
  let reisenVals = reisenSheet
    .getRange(2, 1, reisenRows, reisenCols)
    .getValues();
  let reisenNotes = reisenSheet
    .getRange(2, 1, reisenRows, reisenCols)
    .getNotes();

  let buchungenRows = buchungenSheet.getLastRow() - 1; // first row = headers
  let buchungenCols = buchungenSheet.getLastColumn();
  let buchungenVals: any[][];
  let buchungenNotes: string[][];
  // getRange with 0 rows throws an exception instead of returning an empty array
  if (buchungenRows == 0) {
    buchungenVals = [];
    buchungenNotes = [];
  } else {
    buchungenVals = buchungenSheet
      .getRange(2, 1, buchungenRows, buchungenCols)
      .getValues();
    buchungenNotes = buchungenSheet
      .getRange(2, 1, buchungenRows, buchungenCols)
      .getNotes();
  }

  let ezimmer: MapS2I = {};
  let dzimmer: MapS2I = {};
  for (let b = 0; b < buchungenRows; b++) {
    if (!isEmpty(buchungenNotes[b][0])) continue;
    let tour = buchungenVals[b][tourIndexB - 1];
    let einzeln = buchungenVals[b][einzelnIndex - 1].startsWith("Alleine");
    let zimmer = einzeln ? ezimmer : dzimmer;
    let anzahl: number = zimmer[tour];
    if (anzahl == null) {
      zimmer[tour] = 1;
    } else {
      zimmer[tour] = anzahl + 1;
    }
  }

  let tourIndexR = headers["Reisen"]["Reise"];
  let dzAnzahlIndex = headers["Reisen"]["DZ-Anzahl"];
  let ezAnzahlIndex = headers["Reisen"]["EZ-Anzahl"];
  let dzRestIndex = headers["Reisen"]["DZ-Rest"];
  let ezRestIndex = headers["Reisen"]["EZ-Rest"];
  for (let r = 0; r < reisenRows; r++) {
    if (!isEmpty(reisenNotes[r][0])) continue;
    let tour = reisenVals[r][tourIndexR - 1];
    let dzAnzahl: number = reisenVals[r][dzAnzahlIndex - 1];
    let ezAnzahl: number = reisenVals[r][ezAnzahlIndex - 1];
    let dzGebucht: number = dzimmer[tour];
    let ezGebucht: number = ezimmer[tour];
    if (dzGebucht == null) dzGebucht = 0;
    if (ezGebucht == null) ezGebucht = 0;
    let dzRest: number = dzAnzahl - dzGebucht;
    let ezRest: number = ezAnzahl - ezGebucht;
    if (dzRest < 0) {
      SpreadsheetApp.getUi().alert(
        "DZ der Reise '" + tour + "' sind überbucht!"
      );
      dzRest = 0;
    }
    if (ezRest < 0) {
      SpreadsheetApp.getUi().alert(
        "EZ der Reise '" + tour + "' sind überbucht!"
      );
      ezRest = 0;
    }
    let dzRestR: number = reisenVals[r][dzRestIndex - 1];
    if (dzRest !== dzRestR) {
      reisenSheet.getRange(2 + r, dzRestIndex).setValue(dzRest);
      SpreadsheetApp.getUi().alert(
        "DZ-Rest der Reise '" +
          tour +
          "' von " +
          dzRestR +
          " auf " +
          dzRest +
          " geändert!"
      );
    }
    let ezRestR: number = reisenVals[r][ezRestIndex - 1];
    if (ezRest !== ezRestR) {
      reisenSheet.getRange(2 + r, ezRestIndex).setValue(ezRest);
      SpreadsheetApp.getUi().alert(
        "EZ-Rest der Reise '" +
          tour +
          "' von " +
          ezRestR +
          " auf " +
          ezRest +
          " geändert!"
      );
    }
  }
}

function updateForm() {
  let reisenHdrs = headers["Reisen"];
  let reisenRows = reisenSheet.getLastRow() - 1; // first row = headers
  let reisenCols = reisenSheet.getLastColumn();
  let reisenVals = reisenSheet
    .getRange(2, 1, reisenRows, reisenCols)
    .getValues();
  let reisenNotes = reisenSheet
    .getRange(2, 1, reisenRows, reisenCols)
    .getNotes();
  // Logger.log("reisen %s %s", reisenVals.length, reisenVals);
  let reisenObjs = [];
  for (let i = 0; i < reisenVals.length; i++) {
    if (!isEmpty(reisenNotes[i][0])) continue;
    let reisenObj: MapS2S = {};
    for (let hdr in reisenHdrs) {
      let idx = reisenHdrs[hdr];
      // Logger.log("hdr %s %s", hdr, idx);
      reisenObj[hdr] = reisenVals[i][idx - 1];
    }
    let ok = true;
    // check if all cells of Reise row are nonempty
    for (let hdr in reisenHdrs) {
      if (isEmpty(reisenObj[hdr])) ok = false;
    }
    if (ok) {
      ok = +reisenObj["DZ-Rest"] > 0 || +reisenObj["EZ-Rest"] > 0;
    }
    if (ok) reisenObjs.push(reisenObj);
  }
  Logger.log("reisenObjs=%s", reisenObjs);

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let formUrl = ss.getFormUrl();
  // Logger.log("formUrl2 %s", formUrl);
  let form: GoogleAppsScript.Forms.Form = FormApp.openByUrl(formUrl);
  let items = form.getItems();
  let reisenItem: GoogleAppsScript.Forms.CheckboxItem = null;
  for (let item of items) {
    //   let itemType = item.getType();
    //   Logger.log("title %s it %s %s", item.getTitle(), itemType, item.getIndex());
    if (item.getTitle() === tourFrage) {
      reisenItem = item.asCheckboxItem();
      break;
    }
  }
  if (reisenItem == null) {
    SpreadsheetApp.getUi().alert(
      'Das Formular hat keine Frage "Bei welchen Touren ...?"'
    );
    return;
  }
  let choices = [];
  let descs = [];
  for (let reiseObj of reisenObjs) {
    let mr: string = reiseObj["Reise"];
    mr = mr.replace(",", ""); // mehrere Buchungen werden durch Komma getrennt, s.o.
    let desc =
      mr +
      ", EZ " +
      reiseObj["EZ-Preis"] +
      "€, " +
      reiseObj["EZ-Rest"] +
      " frei, DZ " +
      reiseObj["DZ-Preis"] +
      "€, " +
      reiseObj["DZ-Rest"] +
      " frei";
    Logger.log("mr %s desc %s", mr, desc);
    descs.push(desc);
    let choice = reisenItem.createChoice(mr);
    choices.push(choice);
  }
  let beschreibung =
    "Sie können eine oder mehrere Touren ankreuzen.\nBitte beachten Sie die Anzahl noch freier Zimmer!\n" +
    descs.join("\n");
  reisenItem.setHelpText(beschreibung);
  reisenItem.setChoices(choices);
}

function sendWrongIbanEmail(empfaenger: string, iban: string) {
  var subject = "Falsche IBAN";
  var body =
    "Die von Ihnen bei der Buchung von ADFC Mehrtagestouren übermittelte IBAN " +
    iban +
    " ist leider falsch! Bitte wiederholen Sie die Buchung mit einer korrekten IBAN.";
  GmailApp.sendEmail(empfaenger, subject, body);
}

let ibanLen: MapS2I = {
  NO: 15,
  BE: 16,
  DK: 18,
  FI: 18,
  FO: 18,
  GL: 18,
  NL: 18,
  MK: 19,
  SI: 19,
  AT: 20,
  BA: 20,
  EE: 20,
  KZ: 20,
  LT: 20,
  LU: 20,
  CR: 21,
  CH: 21,
  HR: 21,
  LI: 21,
  LV: 21,
  BG: 22,
  BH: 22,
  DE: 22,
  GB: 22,
  GE: 22,
  IE: 22,
  ME: 22,
  RS: 22,
  AE: 23,
  GI: 23,
  IL: 23,
  AD: 24,
  CZ: 24,
  ES: 24,
  MD: 24,
  PK: 24,
  RO: 24,
  SA: 24,
  SE: 24,
  SK: 24,
  VG: 24,
  TN: 24,
  PT: 25,
  IS: 26,
  TR: 26,
  FR: 27,
  GR: 27,
  IT: 27,
  MC: 27,
  MR: 27,
  SM: 27,
  AL: 28,
  AZ: 28,
  CY: 28,
  DO: 28,
  GT: 28,
  HU: 28,
  LB: 28,
  PL: 28,
  BR: 29,
  PS: 29,
  KW: 30,
  MU: 30,
  MT: 31,
};

function isValidIban(iban: string) {
  if (!iban.match(/^[\dA-Z]+$/)) return false;
  let len = iban.length;
  if (len != ibanLen[iban.substr(0, 2)]) return false;
  iban = iban.substr(4) + iban.substr(0, 4);
  let s = "";
  for (let i = 0; i < len; i += 1) s += parseInt(iban.charAt(i), 36);
  let m = +s.substr(0, 15) % 97;
  s = s.substr(15);
  for (; s; s = s.substr(13)) m = +("" + m + s.substr(0, 13)) % 97;
  return m == 1;
}
