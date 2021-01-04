# Tabellen, Formulare und Skripten für die ADFC München Technikkurse

## Links

Das Formular kann unter https://forms.gle/BstMopCR9vYiiiQL8 ausgefüllt werden. Die Backend-Tabelle dazu kann unter
https://docs.google.com/spreadsheets/d/1jm8GL-Xblyh7vORDvWljbBgz0hbFXArrWpl96WCZVGU/edit?usp=sharing
angesehen werden.

Beides muß später, wie schon bei den Anmeldungen zur RFS und den Saisonkarten, in die Ablage eines neu zu schaffenden Benutzers Technikkurse@adfc-muenchen.de kopiert werden, damit dieser Benutzer als Absender in den verschickten Emails erscheint. Momentan steckt alles noch in meiner privaten Ablage.

## Neuerungen

Das Formular hat keinen Limiter und kein Skript mehr. Die Email-Verifikation wird von der Backend-Tabelle aus gesteuert. Die Kurse im Formular speisen sich aus einem Sheet der Backend-Tabelle. Das Formular wird nach jeder Buchung auf den aktuellen Stand der freien Kursplätze gebracht. Von der Tabelle aus kann das Formular auch über einen Update-Knopf aktualisiert werden, wenn z.B. Kurse hinzugefügt oder gelöscht werden.

## Die Backend-Tabelle

Die Backend-Tabelle hat die 3 Sheets Email-Verifikation, Buchungen, Kurse.

### Notizen

Für die Buchungen- und Kursetabelle gilt: Steht in der Spalte A einer Tabelle eine Notiz, wird diese Zeile ignoriert. Buchungen mit einer ungültigen IBAN bekommen z.B. eine Notiz "IBAN ungültig". Buchungen für einen Kurs dessen Anzahl freier Plätze 0 ist, bekommen eine Notiz "Ausgebucht".
Kurse, die stattgefunden haben, können mit "Abgeschlossen" o.ä. notiert werden. Stornierte Kurse können als "Storniert" notiert werden, usw. Damit soll das Löschen von Zeilen unnötig gemacht werden, damit die Historie einsehbar bleibt.

### Kurse

Im Kurse-Sheet stehen die Nummer , der Name und das Datum des Kurses, einschließlich von-bis, die Anzahl der Kursplätze, und die Anzahl der freien Kursplätze. Jeder Kurs (sofern ohne Notiz) erscheint im Formular mit freien Plätzen.

### Buchungen

Hierhin schreibt das Formular für die Buchung die Daten. Vom Skript werden noch weitere Felder für die Anmeldebestätigung und die Email-Verifikation hinzugefügt. Für die Abbuchung werden weitere Felder hinzukommen.

### Email-Verifikation

Hierhin schreibt das Formular für die Email-Verifikation die Daten. Wenn in diesem Sheet eine Email steht, gilt sie als verifiziert und das Ausfüllen des Verifikations-Formulars wird bei Folgebuchungen nicht wieder verlangt.

## Das Skript

### Technische Details

Das Skript ist in Typescript programmiert (Endung .ts). Durch "flask push" wird es von der Datei Code.ts zur Google Apps Skript-Datei Code.gs übersetzt. In der ersten Zeile von Code.gs steht deshalb

```
// Compiled using ts2gas 3.6.4 (TypeScript 4.1.3)
```

### Einstiegspunkte

Über Skripteditor/Bearbeiten/Trigger des aktuellen Projekts ist ein Trigger "Aus Tabelle - Bei Formularübermittlung" gesetzt, der die Funktion dispatch aufruft.

Über die "onOpen"-Funktion wird im Backend ein Menü-Eintrag ADFC-TK erzeugt, mit den Menüpunkten "Anmeldebestätigung senden" und "Update". Die Anmeldebestätigung wird aber normalerweise automatisch verschickt. Wenn das nicht geklappt haben sollte, kann man eine Zeile im Buchungen-Sheet wählen und sendet dann die Anmeldebestätigung für diesen Kurs.

Update kann man aufrufen, wenn man in Buchungen oder Kurse irgendwelche Änderungen durchgeführt hat, wie z.B. Stornierungen oder neue Kurse. Damit wird für alle Kurse die Anzahl der freien Plätze gleich der Anzahl verfügbarer Plätze minus die Anzahl gebuchter Plätze gesetzt, und das Formular upgedatet (geupdatet?). Der Aufrufer wird über jede Änderung informiert, um Überraschungen vorzubeugen. Vor allem wird auch vor Überbuchungen gewarnt (Zahl freier Plätze < 0), die durch Manipulation der Tabellen entstehen könnten.

### Mehrfachbuchungen

Im Formular kann man mehrere Kurse gleichzeitig buchen. Das Skript erstellt für jeden gebuchten Kurs eine eigene Zeile im Buchungen-Sheet.

### Sortieren

Die Tabellen können beliebig sortiert werden, um z.B. in Buchungen alle Buchungen für einen Kurs oder alle Buchungen einer Person zusammenhängend zu sehen. Die Zeile 1 mit den Headern ist fixiert und sollte das auch bleiben (Ansicht/Fixieren/1 Zeile).

### Erste Mail oder Anmeldebestätigung

Nach Abschicken des Formulars erhält der Teilnehmer unter der von ihm angegebenen Email-Adresse eine erste Mail. In dieser steht, was er gebucht hat, ob er vorgemerkt ist oder der Kurs ausgebucht war, und er wird ggfs. zur Verifikation der Email aufgefordert, falls die Email-Adresse noch unbekannt ist. Ist die Email-Adresse schon bekannt, erhält er gleich eine Mail mit Zu- oder Absage.

### Zweite Mail / Anmeldebestätigung

Hat der Kunde als erste email einen Aufruf zur Emailverifikation bekommen, und hat er das Formular dazu abgeschickt, bekommt er jetzt die Anmeldebestätigung. Diese enthält den oder die Kurs(e), und den Gesamt-Preis, mit der Androhung der baldigen Abbuchung. Der Gesamtpreis bestimmt sich aus seinen Angaben zu ADFC-Mitglied und Studierender, mal Anzahl der Kurse.

### Löschen

Bitte lieber nicht. Lieber in der ersten Spalte eine Notiz setzen.

### Änderungswünsche

Je eher, desto weniger Zeitdruck, desto lieber.

### Fake-IBAN

DE91100000000123456789 ist eine syntaktisch gültige IBAN für Testzwecke.

## Setup Visual Studio Code / Typescript

See https://developers.google.com/apps-script/guides/typescript.
Installieren: node.js, visual studio code.
Dann:

```
    npm install -g @google/clasp
    npm install -S @types/google-apps-script
    clasp clone <scriptid>  # im Projektverzeichnis, scriptid = id des skripts, daß der SkriptEditor aufmacht.
```

Danach existiert Datei .clasp.json mit der scriptid.
In VSCode das Projektverzeichnis öffnen. Terminal öffnen, clasp push -w starten.
