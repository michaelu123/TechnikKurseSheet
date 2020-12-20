let inited = false;
let headers = {};
let reisenSheet: GoogleAppsScript.Spreadsheet.Sheet = null;

// Indices are 1-based!!
let mailIndex = 0; // E-Mail-Adresse
let herrFrauIndex = 0; // Anrede
let nameIndex = 0; // Name
let zustimmungsIndex = 0; // Zustimmung zur SEPA-Lastschrift
let bestätigungsIndex = 0; // Bestätigung (der Teilnahmebedingungen)
let verifikationsIndex = 0; // Verifikation (der Email-Adresse)
let anmeldebestIndex = 0; // Anmeldebestätigung (gesendet)

function test() {
  init();
  let e = {
    namedValues: {
      "Reisen Sie alleine oder zu zweit?": ["Alleine (EZ)"],
      Name: ["Uhlenberg"],
      Vorname: ["Michael"],
      "Bei welchen Touren möchten Sie mitfahren?": [
        // "Fahrradtour um den Gardasee vom 1.5. bis 12.5.",
        // "Transalp von Salzburg nach Venedig vom 2.5. bis 13.5.",
        "Entlang der Drau vom 3.5. bis 14.5",
      ],
      Anrede: ["Herr"],
      "E-Mail-Adresse": ["michael.uhlenberg@t-online.de"],
      "Gleiche Adresse wie Teilnehmer 1 ?": [],
      "IBAN-Kontonummer": ["DE15700202702530131478"],
    },
  };
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Buchungen");
  e["range"] = sheet.getRange(5, 1, 1, sheet.getLastColumn());
  dispatch(e);
}

function init() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  for (let sheet of sheets) {
    let sheetName = sheet.getName();
    let sheetHeaders = {};
    // Logger.log("sheetName %s", sheetName);
    headers[sheetName] = sheetHeaders;
    let numCols = sheet.getLastColumn();
    // Logger.log("numCols %s", numCols);
    let row1Vals = sheet.getRange(1, 1, 1, numCols).getValues();
    // Logger.log("sheetName %s row1 %s", sheetName, row1Vals);
    for (let i = 0; i < numCols; i++) {
      let v = row1Vals[0][i];
      if (!v) continue;
      sheetHeaders[v] = i + 1;
    }
    Logger.log("sheet %s %s", sheetName, sheetHeaders);

    if (sheet.getName() == "Reisen") {
      reisenSheet = sheet;
    }
    if (sheet.getName() == "Buchungen") {
      mailIndex = sheetHeaders["E-Mail-Adresse"];
      herrFrauIndex = sheetHeaders["Anrede"];
      nameIndex = sheetHeaders["Name"];
      bestätigungsIndex = sheetHeaders["Bestätigung"];
      verifikationsIndex = sheetHeaders["Verifikation"];
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

function addColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetHeaders: {},
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

//#########################################################
function anredeText(herrFrau, name) {
  if (herrFrau === "Herr") {
    return "Sehr geehrter Herr " + name;
  } else {
    return "Sehr geehrte Frau " + name;
  }
}

//#########################################################
function kursName(fileName) {
  let nameString = fileName.toString(fileName);
  let codeName = "";
  codeName = nameString.substring(0, 8);
  return codeName;
}

//#########################################################
function kursTermine(code) {
  let text = "";
  text =
    "Interner Fehler: Für diesen Kurs ist kein Kursdatum im Skript hinterlegt";
  SpreadsheetApp.getUi().alert(text);
  return text;
}

//#########################################################
function heuteString() {
  return Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    "YYYY-MM-dd HH:mm:ss"
  );
}

//#########################################################
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

//#########################################################
function mailSchicken() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let anmeldeSheets = SpreadsheetApp.openById(
    "1Qx1t4dPqbrZYQ8cck9kPv_P8RPQfZHAjxeWr2nJavkQ"
  ).getSheets(); // Textbausteine
  let anmeldeTexte = anmeldeSheets[0];
  let textColumn = 2;
  let subjectRow = 3;
  let bodyRow = 4;
  let subjectTemplate = anmeldeTexte
    .getRange(subjectRow, textColumn)
    .getValue()
    .toString();
  let bodyTemplate = anmeldeTexte
    .getRange(bodyRow, textColumn)
    .getValue()
    .toString();
  let row = sheet.getSelection().getCurrentCell().getRow();
  if (row < 2 || row > sheet.getLastColumn()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Teilnehmerzeile selektieren"
    );
    return;
  }
  // update sheet
  sheet.getRange(row, anmeldebestIndex).setValue(heuteString());

  // setting up mail
  let code = kursName(sheet.getName());
  let empfaenger = sheet.getRange(row, mailIndex).getValue();
  let subject = Utilities.formatString(subjectTemplate, code);
  // Anrede
  let anrede = anredeText(
    sheet.getRange(row, herrFrauIndex).getValue(),
    sheet.getRange(row, nameIndex).getValue()
  );
  // Zahlungstext
  let zahlungsText = "";
  // let zahlungsart = sheet.getRange(row, ZahlungsartIndex).getValue().toString();
  // if (zahlungsart === "Überweisung") {
  //   let überweisungsRow = 5;
  //   let überweisungsTemplate = anmeldeTexte
  //     .getRange(überweisungsRow, textColumn)
  //     .getValue()
  //     .toString();
  //   zahlungsText = Utilities.formatString(
  //     überweisungsTemplate,
  //     Kursgebühr,
  //     code
  //   );
  // } else {
  //   let sepaRow = 6;
  //   let sepaTemplate = anmeldeTexte
  //     .getRange(sepaRow, textColumn)
  //     .getValue()
  //     .toString();
  //   zahlungsText = Utilities.formatString(sepaTemplate, Kursgebühr);
  // }
  // Zusammensetzung Body
  let body = Utilities.formatString(
    bodyTemplate,
    anrede,
    code,
    kursTermine(code),
    zahlungsText
  );

  // setting up mail
  GmailApp.sendEmail(empfaenger, subject, body, {
    name: "Radfahrschule ADFC München e.V.",
    replyTo: "radfahrschule@adfc-muenchen.de",
    attachments: attachmentFiles(),
  });
}

//#########################################################
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("ADFC-MTT")
    .addItem("Anmeldebestätigung senden", "mailSchicken")
    .addItem("Update", "update")
    .addItem("Test", "test")
    .addToUi();
}

function dispatch(e) {
  // let keys = Object.keys(e);
  // Logger.log("verif2", keys);
  // for (let key of keys) {
  //   Logger.log("key %s val %s keys %s", key, e[key], Object.keys(e[key]));
  // }
  let range: GoogleAppsScript.Spreadsheet.Range = e["range"];
  Logger.log("A1 %s", range.getA1Notation());
  let sheet = range.getSheet();
  Logger.log("dispatch sheet %s", sheet.getName());
  if (sheet.getName() == "Test") checkBestellung(e);
  if (sheet.getName() == "Buchungen") checkBestellung(e);
  if (sheet.getName() == "Email-Verifikation") verifyEmail();
}

function verifyEmail() {
  let ssheet = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ssheet.getSheetByName("Email-Verifikation");
  let evalues = evSheet.getSheetValues(
    2,
    1,
    evSheet.getLastRow() - 1,
    evSheet.getLastColumn()
  ); // Mit dieser Email-Adresse

  let buchungenSheet = ssheet.getSheetByName("Buchungen");
  let numRows = buchungenSheet.getLastRow();
  if (numRows <= 1) return;
  let rvalues = buchungenSheet.getSheetValues(
    2,
    1,
    numRows - 1,
    buchungenSheet.getLastColumn()
  );
  Logger.log("rvalues %s", rvalues);

  for (let bx in rvalues) {
    let bxi = +bx; // confusingly, bx is initially a string, and is interpreted as A1Notation in sheet.getRange(bx) !
    let rrow = rvalues[bxi];
    if (rrow[mailIndex - 1] != "" && rrow[verifikationsIndex - 1] == "") {
      let raddr = rrow[1];
      for (let ex in evalues) {
        let erow = evalues[ex];
        if (erow[1] != "Ja" || erow[2] == "") continue;
        let eaddr = erow[2];
        if (eaddr != raddr) continue;
        // Bestellungen[Verifiziert] = Email-Verif[Zeitstempel]
        buchungenSheet.getRange(bxi + 2, verifikationsIndex).setValue(erow[0]);
        rrow[verifikationsIndex - 1] = erow[0];
        break;
      }
    }
  }
}

function checkBestellung(e) {
  let keys = Object.keys(e);
  Logger.log("checkBest", keys, typeof e);
  for (let key of keys) {
    Logger.log("key %s val %s", key, e[key]);
  }

  let range: GoogleAppsScript.Spreadsheet.Range = e["range"];
  let sheet = range.getSheet();
  let row = range.getRow();
  let cellA = range.getCell(1, 1);
  Logger.log("sheet %s row %s cellA %s", sheet, row, cellA.getA1Notation());

  let iban = e.namedValues["IBAN-Kontonummer"][0];
  let emailTo = e.namedValues["E-Mail-Adresse"][0];
  Logger.log("iban=%s emailTo=%s %s", iban, emailTo, typeof emailTo);
  if (!isValidIban(iban)) {
    sendWrongIbanEmail(emailTo, iban);
    cellA.setNote("Ungültige IBAN");
    return;
  }

  // Die Zellen Zustimmung und Bestätigung sind im Formular als Pflichtantwort eingetragen
  // und können garnicht anders als gesetzt sein. Sonst hier prüfen analog zu IBAN.

  let einzeln = e.namedValues[
    "Reisen Sie alleine oder zu zweit?"
  ][0].startsWith("Alleine");
  let restCol = einzeln
    ? headers["Reisen"]["EZ-Rest"]
    : headers["Reisen"]["DZ-Rest"];
  let tourFrage = "Bei welchen Touren möchten Sie mitfahren?";
  let touren: Array<string> = e.namedValues[tourFrage];
  if (touren.length == 0) {
    // cannot happen, answer is mandatory
    return;
  }
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
    for (let j = 0; j < reisen.length; j++) {
      if (reisen[j][0] == touren[i]) {
        let rest = reisenSheet.getRange(2 + j, restCol).getValue();
        if (rest <= 0) {
          msgs.push("Die Reise '" + touren[i] + "' ist leider ausgebucht.");
          sheet.getRange(buchungsRowNumbers[i], 1).setNote("Ausgebucht");
        } else {
          msgs.push("Sie sind für die Reise '" + touren[i] + "' vorgemerkt.");
          reisenSheet.getRange(2 + j, restCol).setValue(rest - 1);
          restChanged = true;
        }
      }
    }
  }
  if (restChanged) {
    updateForm();
  }
  Logger.log("msgs: ", msgs);
  sendeAntwort(e, msgs.join("<br />"), sheet, buchungsRowNumbers);
}

function sendeAntwort(
  e,
  msg: String,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  buchungsRowNumbers: Array<number>
) {
  let emailTo = e.namedValues["E-Mail-Adresse"][0];
  Logger.log("emailTo=" + emailTo);

  let templateFile = "emailVerif.html";

  // do we already know this email address?
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ss.getSheetByName("Email-Verifikation");
  let evalues = evSheet.getSheetValues(2, 1, evSheet.getLastRow() - 1, 3);
  for (let i = 0; i < evalues.length; i++) {
    // Mit dieser Email-Adresse
    if (evalues[i][2] == emailTo) {
      templateFile = "emailReply.html"; // yes, don't ask for verification
      let verifiedCol = headers["Buchungen"]["Verifikation"];
      for (let j = 0; j < buchungsRowNumbers.length; j++) {
        sheet
          .getRange(buchungsRowNumbers[j], verifiedCol)
          .setValue(evalues[i][0]);
      }
      break;
    }
  }

  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    templateFile
  );
  template.anrede = anrede(e);
  template.msg = msg;
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

function anrede(e) {
  let res = "";
  let anrede = e.namedValues["Anrede"][0];
  let vorname = e.namedValues["Vorname"];
  if (vorname == null || vorname.length == 0)
    vorname = e.namedValues["Vorname 1"];
  let name = e.namedValues["Name"];
  if (name == null || name.length == 0) name = e.namedValues["Name 1"];
  if (anrede == "Herr") {
    anrede = "Sehr geehrter Herr ";
  } else {
    anrede = "Sehr geehrte Frau ";
  }
  return anrede + vorname + " " + name;
}

function update() {
  if (!inited) {
    init();
  }
  updateRest();
  updateForm();
}

function updateRest() {}

function updateForm() {
  let reisenHdrs: {} = headers["Reisen"];
  let reisenRows = reisenSheet.getLastRow() - 1; // first row = headers
  let reisenCols = reisenSheet.getLastColumn();
  let reisenVals = reisenSheet
    .getRange(2, 1, reisenRows, reisenCols)
    .getValues();
  Logger.log("reisen %s %s", reisenVals.length, reisenVals);
  let reisenObjs = [];
  for (let i = 0; i < reisenVals.length; i++) {
    let reisenObj = {};
    for (let hdr in reisenHdrs) {
      let idx = reisenHdrs[hdr];
      // Logger.log("hdr %s %s", hdr, idx);
      reisenObj[hdr] = reisenVals[i][idx - 1];
    }
    let ok = true;
    // check if all cells of Reise row are nonempty
    for (let hdr in reisenHdrs) {
      if (reisenObj[hdr] === "") ok = false;
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
    if (item.getTitle().startsWith("Bei welchen Touren")) {
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
    mr = mr.replace(",", ""); // mehrere Buchungen werden durch Komma getrennt
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

let ibanLen = {
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
  iban = iban.replace(/\s/g, "").toUpperCase();
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

/*
19.12.2020, 07:16:02	Info	key namedValues val {Reisen Sie alleine oder zu zweit?=[Alleine (EZ)], Name=[Uhlenberg, ], Bei welchen Touren möchten Sie mitfahren?=[Fahrradtour um den Gardasee vom 1.5. bis 12.5., Transalp von Salzburg nach Venedig vom 2.5. bis 13.5., Entlang der Drau vom 3.5. bis 14.5], Telefonnummer für Rückfragen 2=[], Bestätigung=[Ich habe die Teilnahmebedingungen zur Kenntnis genommen und verstanden.], Anrede=[Herr, ], Postleitzahl=[81479, ], Zeitstempel=[19.12.2020 07:16:01], Straße und Hausnummer 2=[], Straße und Hausnummer=[Ludwigshöher Str., ], IBAN-Kontonummer=[DE44ZZZ00000793122], E-Mail-Adresse=[michael.uhlenberg@t-online.de], Postleitzahl 2=[], Anrede 2=[], Telefonnummer für Rückfragen=[015771574094, ], Zustimmung zur SEPA-Lastschrift=[Ich stimme der SEPA-Lastschrift zu], Ort 2=[], Vorname 2=[], Name der Bank (optional)=[hvb], Name des Kontoinhabers=[muh], Ort=[München, ], Name 2=[], Vorname=[Michael, ], Gleiche Adresse wie Teilnehmer 1 ?=[], =[]} keys [Zustimmung zur SEPA-Lastschrift, Anrede, Straße und Hausnummer, Gleiche Adresse wie Teilnehmer 1 ?, Postleitzahl 2, Vorname 2, Bei welchen Touren möchten Sie mitfahren?, Ort, Zeitstempel, IBAN-Kontonummer, Name der Bank (optional), Name 2, Telefonnummer für Rückfragen 2, Bestätigung, E-Mail-Adresse, Straße und Hausnummer 2, Ort 2, Name des Kontoinhabers, Vorname, , Name, Postleitzahl, Reisen Sie alleine oder zu zweit?, Telefonnummer für Rückfragen, Anrede 2]
19.12.2020, 07:16:02	Info	key range val Range keys [columnEnd, columnStart, rowEnd, rowStart]
19.12.2020, 07:16:02	Info	key source val Spreadsheet keys []
19.12.2020, 07:16:02	Info	key triggerUid val 5721330 keys [0, 1, 2, 3, 4, 5, 6]
19.12.2020, 07:16:02	Info	key values val [19.12.2020 07:16:01, michael.uhlenberg@t-online.de, Fahrradtour um den Gardasee vom 1.5. bis 12.5., Transalp von Salzburg nach Venedig vom 2.5. bis 13.5., Entlang der Drau vom 3.5. bis 14.5, Alleine (EZ), Herr, Michael, Uhlenberg, 81479, München, Ludwigshöher Str., 015771574094, , , , , , , , , , , , , , , , muh, hvb, DE44ZZZ00000793122, Ich stimme der SEPA-Lastschrift zu, Ich habe die Teilnahmebedingungen zur Kenntnis genommen und verstanden., ] keys [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]

*/
