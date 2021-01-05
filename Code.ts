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
let kurseSheet: GoogleAppsScript.Spreadsheet.Sheet;
let buchungenSheet: GoogleAppsScript.Spreadsheet.Sheet;

// Indices are 1-based!!
// Buchungen:
let mailIndex: number; // E-Mail-Adresse
let mitgliederNrIndex: number;
let studentIndex: number;
let herrFrauIndex: number; // Anrede
let nameIndex: number; // Name
let zustimmungsIndex: number; // Zustimmung zur SEPA-Lastschrift
let bestätigungsIndex: number; // Bestätigung (der Teilnahmebedingungen)
let verifikationsIndex: number; // Verifikation (der Email-Adresse)
let anmeldebestIndex: number; // Anmeldebestätigung (gesendet)
let kurseIndex: number; // Welche Kurse möchten Sie belegen?
let bezahltIndex: number; // Bezahlt

// Kurse:
let kursNummerIndex: number; // Kursnummer
let kursTitelIndex: number; // Kurstitel
let kursDatumIndex: number; // Kursdatum
let kursPlätzeIndex: number; // Kursplätze
let restPlätzeIndex: number; // Restplätze

// map Buchungen headers to print headers
let printCols = new Map([
  ["Vorname", "Vorname"],
  ["Name", "Nachname"],
  ["ADFC-Mitgliedsnummer", "Mitglied"],
  ["Studieren Sie?", "Student"],
  ["Telefonnummer für Rückfragen", "Telefon"],
  ["Anmeldebestätigung", "Bestätigt"],
  ["Bezahlt", "Bezahlt"],
]);

const kursFrage = "Welche Kurse möchten Sie belegen?";

interface Event {
  namedValues: { [others: string]: string[] };
  range: GoogleAppsScript.Spreadsheet.Range;
  [others: string]: any;
}

function isEmpty(str: string | undefined | null) {
  return !str || 0 === str.length; // I think !str is sufficient...
}

function test() {
  init();
  let e: Event = {
    namedValues: {
      Vorname: ["Michael"],
      Name: ["Uhlenberg"],
      Anrede: ["Herr"],
      "E-Mail-Adresse": ["michael.uhlenberg@t-online.de"],
      "IBAN-Kontonummer": ["DE91100000000123456789"],
      "ADFC-Mitgliedsnummer": ["1234"],
      "Studieren Sie?": ["Nein"],
      [kursFrage]: [
        "K712: Ich mach' das schon - Pannenhilfe für Mütter und Töchter am 01.08.2020",
        "K713: Grundkurs für Räder mit Kettenschaltung am 11.08.2020",
        //  "K710: Praxiskurs - Kettenschaltung am 24.06.2020"
      ],
    },
    range: buchungenSheet.getRange(2, 1, 1, buchungenSheet.getLastColumn()),
    // range: SpreadsheetApp.getActiveSpreadsheet()
    //  .getSheetByName("Email-Verifikation")
    //  .getRange(2, 1, 1, 3),
  };
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

    if (sheet.getName() == "Kurse") {
      kurseSheet = sheet;
      kursNummerIndex = sheetHeaders["Kursnummer"];
      kursTitelIndex = sheetHeaders["Kurstitel"];
      kursDatumIndex = sheetHeaders["Kursdatum"];
      kursPlätzeIndex = sheetHeaders["Kursplätze"];
      restPlätzeIndex = sheetHeaders["Restplätze"];
    }
    if (sheet.getName() == "Buchungen") {
      buchungenSheet = sheet;
      mailIndex = sheetHeaders["E-Mail-Adresse"];
      mitgliederNrIndex = sheetHeaders["ADFC-Mitgliedsnummer"];
      studentIndex = sheetHeaders["Studieren Sie?"];
      herrFrauIndex = sheetHeaders["Anrede"];
      nameIndex = sheetHeaders["Name"];
      bestätigungsIndex = sheetHeaders["Bestätigung"];
      kurseIndex = sheetHeaders[kursFrage];
      verifikationsIndex = sheetHeaders["Verifikation"];
      if (verifikationsIndex == null) {
        verifikationsIndex = addColumn(sheet, sheetHeaders, "Verifikation");
      }
      anmeldebestIndex = sheetHeaders["Anmeldebestätigung"];
      if (anmeldebestIndex == null) {
        anmeldebestIndex = addColumn(sheet, sheetHeaders, "Anmeldebestätigung");
      }
      bezahltIndex = sheetHeaders["Bezahlt"];
      if (bezahltIndex == null) {
        bezahltIndex = addColumn(sheet, sheetHeaders, "Bezahlt");
      }
    }
  }
  inited = true;
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

function kursPreis(mitglied: boolean, student: boolean) {
  if (!mitglied && !student) return 15;
  if (!mitglied && student) return 12.5;
  if (mitglied && !student) return 10;
  if (mitglied && student) return 7.5;
  return 0;
}

function anmeldebestätigung() {
  Logger.log("anmeldebestätigung");
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
  if (isEmpty(rowValues[verifikationsIndex - 1])) {
    SpreadsheetApp.getUi().alert("Email-Adresse nicht verifiziert");
    return;
  }
  if (!isEmpty(rowValues[anmeldebestIndex - 1])) {
    SpreadsheetApp.getUi().alert("Der Kurs wurde schon bestätigt");
    return;
  }
  // setting up mail
  let emailTo: string = rowValues[mailIndex - 1];
  let subject: string = "Bestätigung Ihrer Buchung";
  let herrFrau = rowValues[herrFrauIndex - 1];
  let name = rowValues[nameIndex - 1];
  let mitglied = !isEmpty(rowValues[mitgliederNrIndex - 1]);
  let student = rowValues[studentIndex - 1] === "Ja";

  let anrede: string = anredeText(herrFrau, name);
  let kurs: string = rowValues[kurseIndex - 1];
  let betrag: number = kursPreis(mitglied, student);
  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    "emailBestätigung.html"
  );
  template.anrede = anrede;
  template.kurse = ["Sie sind für den Kurs '" + kurs + "' angemeldet."];
  template.betrag = betrag;

  let htmlText: string = template.evaluate().getContent();
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "Technikkurse ADFC München e.V.",
    replyTo: "michael.uhlenberg@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
  // update sheet
  sheet.getRange(row, anmeldebestIndex).setValue(heuteString());
}

function anmeldebestätigungen(
  buchungenMap: Map<string, number[]>,
  buchungenVals: any[][]
) {
  Logger.log("anmeldebestätigungen");
  let subject: string = "Bestätigung Ihrer Buchung";
  for (let [emailTo, rows] of buchungenMap) {
    let anrede: string;
    let einzelBetrag = 0;
    let betrag = 0;
    let kurse: string[] = [];
    for (let row of rows) {
      let rowValues = buchungenVals[row];
      if (einzelBetrag === 0) {
        let herrFrau = rowValues[herrFrauIndex - 1];
        let name = rowValues[nameIndex - 1];
        let mitglied = !isEmpty(rowValues[mitgliederNrIndex - 1]);
        let student = rowValues[studentIndex - 1] === "Ja";
        anrede = anredeText(herrFrau, name);
        einzelBetrag = kursPreis(mitglied, student);
      }
      let kurs: string = rowValues[kurseIndex - 1];
      kurse.push("Sie sind für den Kurs '" + kurs + "' angemeldet.");
      betrag += einzelBetrag;
    }
    if (betrag === 0) continue;

    // setting up mail
    let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
      "emailBestätigung.html"
    );
    template.anrede = anrede;
    template.kurse = kurse;
    template.betrag = betrag;

    let htmlText: string = template.evaluate().getContent();
    let textbody = "HTML only";
    let options = {
      htmlBody: htmlText,
      name: "Technikkurse ADFC München e.V.",
      replyTo: "michael.uhlenberg@adfc-muenchen.de",
    };
    GmailApp.sendEmail(emailTo, subject, textbody, options);
    // update sheet
    for (let row of rows) {
      buchungenSheet
        .getRange(row + 2, anmeldebestIndex)
        .setValue(heuteString());
    }
  }
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("ADFC-TK")
    // .addItem("Test", "test")
    .addItem("Anmeldebestätigung senden", "anmeldebestätigung")
    .addItem("Update", "update")
    .addItem("Kursteilnehmer drucken", "printKursMembers")
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
  if (sheet.getName() == "Buchungen") checkBuchung(e);
  if (sheet.getName() == "Email-Verifikation") verifyEmail();
}

function verifyEmail() {
  Logger.log("verifyEmail");
  let ssheet = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ssheet.getSheetByName("Email-Verifikation");

  let numRows = evSheet.getLastRow() - 1;
  if (numRows < 1) return;
  let evalues = evSheet.getSheetValues(
    2,
    1,
    numRows,
    evSheet.getLastColumn() // = 3
  );

  numRows = buchungenSheet.getLastRow() - 1;
  if (numRows < 1) return;
  let numCols = buchungenSheet.getLastColumn();
  let buchungenVals = buchungenSheet.getSheetValues(2, 1, numRows, numCols);
  Logger.log("buchungenVals %s", buchungenVals);

  let buchungenNotes = buchungenSheet.getRange(2, 1, numRows, 1).getNotes();

  let buchungenMap = new Map<string, number[]>();
  for (let bx in buchungenVals) {
    let bxi = +bx; // confusingly, bx is initially a string, and is interpreted as A1Notation in sheet.getRange(bx) !
    let brow = buchungenVals[bxi];
    let baddr = (brow[mailIndex - 1] as string).toLowerCase();
    if (isEmpty(baddr)) continue;
    if (!isEmpty(buchungenNotes[bxi][0])) {
      continue;
    }
    for (let ex in evalues) {
      let erow = evalues[ex];
      if (isEmpty(erow[0])) continue;
      let eaddr = (erow[2] as string).toLowerCase();
      if (eaddr != baddr) continue;
      if (erow[1] != "Ja" || isEmpty(eaddr)) continue;
      if (isEmpty(brow[verifikationsIndex - 1])) {
        // Buchungen[Verifiziert] = Email-Verif[Zeitstempel]
        buchungenSheet.getRange(bxi + 2, verifikationsIndex).setValue(erow[0]);
        brow[verifikationsIndex - 1] = erow[0];
      }
      if (isEmpty(brow[anmeldebestIndex - 1])) {
        let rows = buchungenMap.get(baddr);
        if (rows == null) {
          rows = [];
          buchungenMap.set(baddr, rows);
        }
        rows.push(bxi);
      }
      break;
    }
  }
  anmeldebestätigungen(buchungenMap, buchungenVals);
}

function isVerified(emailTo: string, buchungsRowNumbers: number[]): boolean {
  // do we already know this email address?
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ss.getSheetByName("Email-Verifikation");
  let numRows = evSheet.getLastRow() - 1;
  if (numRows < 1) return false;
  let evalues = evSheet.getSheetValues(2, 1, numRows, 3);
  emailTo = emailTo.toLowerCase();
  for (let i = 0; i < evalues.length; i++) {
    let erow = evalues[i];
    if (isEmpty(erow[0])) continue;
    // Mit dieser Email-Adresse
    if ((erow[2] as string).toLowerCase() == emailTo) {
      for (let j = 0; j < buchungsRowNumbers.length; j++) {
        buchungenSheet
          .getRange(buchungsRowNumbers[j], verifikationsIndex)
          .setValue(erow[0]);
      }
      Logger.log("verifyEmail returns true");
      return true;
    }
  }
  Logger.log("verifyEmail returns false");
  return false;
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

function checkBuchung(e: Event) {
  Logger.log("checkBuchung");
  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  let row = range.getRow();
  let cellA = range.getCell(1, 1);
  Logger.log("sheet %s row %s cellA %s", sheet, row, cellA.getA1Notation());

  let ibanNV = e.namedValues["IBAN-Kontonummer"][0];
  let iban = ibanNV.replace(/\s/g, "").toUpperCase();
  let emailTo = e.namedValues["E-Mail-Adresse"][0];
  let anredeE = anrede(e);
  let mitglied = !isEmpty(e.namedValues["ADFC-Mitgliedsnummer"][0]);
  let student = e.namedValues["Studieren Sie?"][0] === "Ja";
  Logger.log(
    "iban=%s emailTo=%s Anrede %s Mitglied %s Student %s",
    iban,
    emailTo,
    anredeE,
    mitglied,
    student
  );
  if (!isValidIban(iban)) {
    sendWrongIbanEmail(emailTo, anredeE, iban);
    cellA.setNote("Ungültige IBAN");
    return;
  }
  if (iban != ibanNV) {
    let cellIban = range.getCell(1, headers["Buchungen"]["IBAN-Kontonummer"]);
    cellIban.setValue(iban);
  }
  // Die Zellen Zustimmung und Bestätigung sind im Formular als Pflichtantwort eingetragen
  // und können garnicht anders als gesetzt sein. Sonst hier prüfen analog zu IBAN.

  let kurse: Array<string> = [];
  let kurseNV: Array<string> = e.namedValues[kursFrage];
  // actually, kurseNV = e.g. ["kurs1, kurs2, kurs3"], i.e. just one element
  for (let kurs of kurseNV) {
    let kurseList = kurs.split(","); // that's the reason why we don't want commas within a kurs title!
    for (let tlItem of kurseList) {
      kurse.push(tlItem.trim());
    }
  }
  // Logger.log("check2", einzeln, personen, restPlätzeIndex, kurse, kurse.length);
  if (kurse.length == 0) {
    Logger.log("check4");
    // cannot happen, answer is mandatory
    return;
  }

  // for each kurs besides the first create a new row, and put the kurs into it
  let buchungsRowNumbers = [row];
  if (kurse.length > 1) {
    let numCols = sheet.getLastColumn();
    let kursCellNo = headers["Buchungen"][kursFrage];
    for (let i = 1; i < kurse.length; i++) {
      let toRow = sheet.getLastRow() + 1;
      if (toRow >= sheet.getMaxRows()) {
        sheet.insertRowAfter(toRow);
      }
      let toRange = sheet.getRange(toRow, 1, 1, numCols);
      range.copyTo(toRange);
      let kursCell = toRange.getCell(1, kursCellNo);
      kursCell.setValue(kurse[i]);
      buchungsRowNumbers.push(toRow);
    }
    // put the first kurs into the original row
    let kursCell = range.getCell(1, kursCellNo);
    kursCell.setValue(kurse[0]);
  }

  let verified = isVerified(emailTo, buchungsRowNumbers);
  let msgs = [];
  let kurseVals: Array<Array<string | Date>> = kurseSheet.getSheetValues(
    2,
    1,
    kurseSheet.getLastRow(),
    kurseSheet.getLastColumn()
  );
  let restChanged = false;
  let betrag = 0;
  let einzelPreis = kursPreis(mitglied, student);
  for (let i = 0; i < kurse.length; i++) {
    let kursFound = false;
    for (let j = 0; j < kurseVals.length; j++) {
      let kurseRow = kurseVals[j];
      if (!kurseRow[0]) continue;
      let kurs = kursNTD(
        kurseRow[kursNummerIndex - 1] as string,
        kurseRow[kursTitelIndex - 1] as string,
        date2Str(kurseRow[kursDatumIndex - 1] as Date)
      );
      if (kurs === kurse[i]) {
        kursFound = true;
        let rest = kurseSheet.getRange(2 + j, restPlätzeIndex).getValue();
        if (rest <= 0) {
          msgs.push("Der Kurs '" + kurs + "' ist leider ausgebucht.");
          sheet.getRange(buchungsRowNumbers[i], 1).setNote("Ausgebucht");
        } else {
          msgs.push(
            "Sie sind für den Kurs '" +
              kurs +
              (verified ? "' angemeldet." : "' vorgemerkt.")
          );
          kurseSheet.getRange(2 + j, restPlätzeIndex).setValue(rest - 1);
          restChanged = true;
          betrag += einzelPreis;
        }
        break;
      }
    }
    if (!kursFound) {
      Logger.log("Kurs '" + kurse[i] + " nicht im Kurse-Sheet!?");
    }
  }
  if (msgs.length == 0) {
    Logger.log("keine Kurse gefunden!?");
    return;
  }
  if (restChanged) {
    updateForm();
  }
  Logger.log("msgs: ", msgs, msgs.length);
  sendeAntwort(emailTo, verified, anredeE, betrag, msgs);
  if (verified) {
    let heute = heuteString();
    for (let row of buchungsRowNumbers) {
      sheet.getRange(row, anmeldebestIndex).setValue(heute);
    }
  }
}

function sendeAntwort(
  emailTo: string,
  verified: boolean,
  anredeE: string,
  betrag: number,
  msgs: Array<string>
) {
  Logger.log("sendeAntwort emailTo=", emailTo);

  let templateFile = verified ? "emailBestätigung.html" : "emailVerif.html";
  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    templateFile
  );
  template.anrede = anredeE;
  template.kurse = msgs;
  template.betrag = betrag;
  template.verifLink =
    "https://docs.google.com/forms/d/e/1FAIpQLSeEcceEKaHoGzwdw2qJlu0fpAKkhECG5CnQhi1jVXOcwt-6sw/viewform?usp=pp_url&entry.1398709542=Ja&entry.576071197=" +
    encodeURIComponent(emailTo);

  let htmlText: string = template.evaluate().getContent();
  let subject = verified
    ? "Bestätigung Ihrer Anmeldung"
    : "Bestätigung Ihrer Email-Adresse";
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "Mehrtageskurse ADFC München e.V.",
    replyTo: "michael.uhlenberg@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
}

function update() {
  if (!inited) init();
  verifyEmail();
  updateKursReste();
  updateForm();
}

function date2Str(ddate: Date): string {
  let sdate: string = Utilities.formatDate(
    ddate,
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    "dd.MM.YYYY"
  );
  return sdate;
}

function kursNTD(kursNummer: string, kursTitel: string, kursDatum: string) {
  return kursNummer + ": " + kursTitel + " am " + kursDatum;
}

function updateKursReste() {
  let kurseRows = kurseSheet.getLastRow() - 1; // first row = headers
  let kurseCols = kurseSheet.getLastColumn();
  let kurseVals = kurseSheet.getRange(2, 1, kurseRows, kurseCols).getValues();
  let kurseNotes = kurseSheet.getRange(2, 1, kurseRows, 1).getNotes();

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
    buchungenNotes = buchungenSheet.getRange(2, 1, buchungenRows, 1).getNotes();
  }

  let kursplätze: MapS2I = {};
  for (let b = 0; b < buchungenRows; b++) {
    if (!isEmpty(buchungenNotes[b][0])) continue;
    let kurs = buchungenVals[b][kurseIndex - 1];
    let anzahl: number = kursplätze[kurs];
    if (anzahl == null) {
      kursplätze[kurs] = 1;
    } else {
      kursplätze[kurs] = anzahl + 1;
    }
  }

  for (let r = 0; r < kurseRows; r++) {
    if (!isEmpty(kurseNotes[r][0])) continue;
    let kursNummer: string = kurseVals[r][kursNummerIndex - 1];
    let kursTitel: string = kurseVals[r][kursTitelIndex - 1];
    let kursDatum: Date = kurseVals[r][kursDatumIndex - 1];
    let kursPlätze: number = kurseVals[r][kursPlätzeIndex - 1];
    let restPlätze: number = kurseVals[r][restPlätzeIndex - 1];
    let kurs = kursNTD(kursNummer, kursTitel, date2Str(kursDatum));

    let kursGebucht: number = kursplätze[kurs];
    if (kursGebucht == null) kursGebucht = 0;
    let kursRest: number = kursPlätze - kursGebucht;
    if (kursRest < 0) {
      SpreadsheetApp.getUi().alert("Der Kurs '" + kurs + "' ist überbucht!");
      kursRest = 0;
    }
    if (kursRest !== restPlätze) {
      kurseSheet.getRange(2 + r, restPlätzeIndex).setValue(kursRest);
      SpreadsheetApp.getUi().alert(
        "Restplätze des Kurs '" +
          kurs +
          "' von " +
          restPlätze +
          " auf " +
          kursRest +
          " geändert!"
      );
    }
  }
}

function updateForm() {
  let kurseHdrs = headers["Kurse"];
  let kurseRows = kurseSheet.getLastRow() - 1; // first row = headers
  let kurseCols = kurseSheet.getLastColumn();
  let kurseVals = kurseSheet.getRange(2, 1, kurseRows, kurseCols).getValues();
  let kurseNotes = kurseSheet.getRange(2, 1, kurseRows, 1).getNotes();
  // Logger.log("kurse %s %s", kurseVals.length, kurseVals);
  let kurseObjs = [];
  for (let i = 0; i < kurseVals.length; i++) {
    if (!isEmpty(kurseNotes[i][0])) continue;
    let kursObj: MapS2S = {};
    for (let hdr in kurseHdrs) {
      let idx = kurseHdrs[hdr];
      if (hdr == "Kursdatum") {
        kursObj[hdr] = date2Str(kurseVals[i][idx - 1]);
      } else {
        kursObj[hdr] = kurseVals[i][idx - 1];
      }
    }
    let ok = true;
    // check if all cells of Kurs row are nonempty
    for (let hdr in kurseHdrs) {
      if (isEmpty(kursObj[hdr])) {
        Logger.log("In Kurse Zeile mit leerem Feld ", hdr);
        ok = false;
      }
    }
    if (ok) {
      ok = +kursObj["Restplätze"] > 0;
    }
    if (ok) kurseObjs.push(kursObj);
  }
  Logger.log("kurseObjs=%s", kurseObjs);

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let formUrl = ss.getFormUrl();
  // Logger.log("formUrl2 %s", formUrl);
  let form: GoogleAppsScript.Forms.Form = FormApp.openByUrl(formUrl);
  let items = form.getItems();
  let kurseItem: GoogleAppsScript.Forms.CheckboxItem = null;
  for (let item of items) {
    //   let itemType = item.getType();
    //   Logger.log("title %s it %s %s", item.getTitle(), itemType, item.getIndex());
    if (item.getTitle() === kursFrage) {
      kurseItem = item.asCheckboxItem();
      break;
    }
  }
  if (kurseItem == null) {
    SpreadsheetApp.getUi().alert(
      'Das Formular hat keine Frage "' + kursFrage + '"!'
    );
    return;
  }
  let choices = [];
  let descs = [];
  for (let kursObj of kurseObjs) {
    let mr: string = kursNTD(
      kursObj["Kursnummer"],
      kursObj["Kurstitel"],
      kursObj["Kursdatum"]
    );
    mr = mr.replace(",", ""); // mehrere Buchungen werden durch Komma getrennt, s.o.
    let desc = mr + ", freie Plätze: " + kursObj["Restplätze"];
    // Logger.log("desc %s", desc);
    descs.push(desc);
    let choice = kurseItem.createChoice(mr);
    choices.push(choice);
  }
  let beschreibung =
    "Sie können einen oder mehrere Kurse ankreuzen. Bitte beachten Sie die Anzahl noch freier Plätze!\n\n" +
    descs.join("\n");
  kurseItem.setHelpText(beschreibung);
  kurseItem.setChoices(choices);
}

function sendWrongIbanEmail(empfaenger: string, anrede: string, iban: string) {
  Logger.log("sendWrongIbanEmail");
  let subject = "Falsche IBAN";
  let body =
    anrede +
    ",\nDie von Ihnen bei der Buchung von ADFC Technikkurse übermittelte IBAN " +
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

// I need any2str because a date copied to temp sheet showed as date.toString().
// A ' in front of the date came too late.
function any2Str(val: any): string {
  if (typeof val == "object" && "getUTCHours" in val) {
    return Utilities.formatDate(
      val,
      SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
      "dd.MM.YYYY"
    );
  }
  return val.toString();
}

function printKursMembers() {
  Logger.log("printKursMembers");
  if (!inited) init();
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Kurse") {
    SpreadsheetApp.getUi().alert(
      "Bitte eine Zeile im Sheet 'Kurse' selektieren"
    );
    return;
  }
  let curCell = sheet.getSelection().getCurrentCell();
  if (!curCell) {
    SpreadsheetApp.getUi().alert(
      "Bitte zuerst eine Zeile im Sheet 'Kurse' selektieren"
    );
    return;
  }
  let row = curCell.getRow();
  if (row < 2 || row > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Kurszeile selektieren"
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
  let kurs = kursNTD(
    rowValues[kursNummerIndex - 1],
    rowValues[kursTitelIndex - 1],
    date2Str(rowValues[kursDatumIndex - 1])
  );

  let buchungenRows = buchungenSheet.getLastRow() - 1; // first row = headers
  let buchungenCols = buchungenSheet.getLastColumn();
  let buchungenVals: any[][];
  let buchungenNotes: string[][];
  // getRange with 0 rows throws an exception instead of returning an empty array
  if (buchungenRows < 1) {
    SpreadsheetApp.getUi().alert("Keine Buchungen gefunden");
    return;
  }
  let rows: string[][] = [];
  buchungenVals = buchungenSheet
    .getRange(2, 1, buchungenRows, buchungenCols)
    .getValues();
  buchungenNotes = buchungenSheet.getRange(2, 1, buchungenRows, 1).getNotes();

  let bHdrs = headers["Buchungen"];
  // first row of temp sheet: the headers
  {
    let row: string[] = [];
    for (let [_, v] of printCols) {
      row.push(v);
    }
    rows.push(row);
  }
  for (let b = 0; b < buchungenRows; b++) {
    if (!isEmpty(buchungenNotes[b][0])) continue;
    let brow = buchungenVals[b];
    if (brow[kurseIndex - 1] === kurs) {
      let row: string[] = [];
      for (let [k, _] of printCols) {
        //for the ' see https://stackoverflow.com/questions/13758913/format-a-google-sheets-cell-in-plaintext-via-apps-script
        // otherwise, telefon number 089... is printed as 89
        let val = any2Str(brow[bHdrs[k] - 1]);
        row.push("'" + val);
      }
      rows.push(row);
    }
  }

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.insertSheet(kurs);
  for (let row of rows) sheet.appendRow(row);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  let range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  sheet.setActiveSelection(range);
  printSelectedRange();
  Utilities.sleep(10000);
  ss.deleteSheet(sheet);
}

function objectToQueryString(obj: any) {
  return Object.keys(obj)
    .map(function (key) {
      return Utilities.formatString("&%s=%s", key, obj[key]);
    })
    .join("");
}

// see https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
let PRINT_OPTIONS = {
  size: 7, // paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B
  fzr: false, // repeat row headers
  portrait: true, // false=landscape
  fitw: true, // fit window or actual size
  gridlines: false, // show gridlines
  printtitle: true,
  sheetnames: true,
  pagenum: "UNDEFINED", // CENTER = show page numbers / UNDEFINED = do not show
  attachment: false,
};

let PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function printSelectedRange() {
  SpreadsheetApp.flush();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let range = sheet.getActiveRange();

  let gid = sheet.getSheetId();
  let printRange = objectToQueryString({
    c1: range.getColumn() - 1,
    r1: range.getRow() - 1,
    c2: range.getColumn() + range.getWidth() - 1,
    r2: range.getRow() + range.getHeight() - 1,
  });
  let url = ss.getUrl();
  Logger.log("url1", url);
  let x = url.indexOf("/edit?");
  url = url.slice(0, x);
  url = url + "/export?format=pdf" + PDF_OPTS + printRange + "&gid=" + gid;
  Logger.log("url2", url);
  let htmlTemplate = HtmlService.createTemplateFromFile("print.html");
  htmlTemplate.url = url;

  let ev = htmlTemplate.evaluate();

  SpreadsheetApp.getUi().showModalDialog(
    ev.setHeight(10).setWidth(100),
    "Drucke Auswahl"
  );
}
