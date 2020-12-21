# Tabellen, Formulare und Skripten für die ADFC Mehrtagestouren (MTT)

## Links

Das Formular kann unter https://forms.gle/yJp12y4tKy8bqs4h6 ausgefüllt werden. Die Backend-Tabelle dazu kann unter
https://docs.google.com/spreadsheets/d/1LNopZRbkggBm4OtRRp-YtRBFKKv1z8no0xLEhk-K8mo/edit?usp=sharing
angesehen werden.

Beides muß später, wie schon bei den Anmeldungen zur RFS und den Saisonkarten, in die Ablage eines neu zu schaffenden Benutzers Mehrtagestouren@adfc-muenchen.de kopiert werden, damit dieser Benutzer als Absender in den verschickten Emails erscheint. Momentan steckt alles noch in meiner privaten Ablage.

Ich habe Thomas eine Schreibberechtigung für die Tabelle erteilt, alle anderen sollten nur lesen können. Bei Bedarf melden.

## Neuerungen

Das Formular hat keinen Limiter und kein Skript mehr. Die Email-Verifikation wird von der Backend-Tabelle aus gesteuert. Die Reisen im Formular speisen sich aus einem Sheet der Backend-Tabelle. Das Formular wird nach jeder Buchung auf den aktuellen Stand der freien Zimmer gebracht. Von der Tabelle aus kann das Formular auch über einen Update-Knopf aktualisiert werden, wenn z.B. Reisen hinzugefügt oder gelöscht werden.

## Die Backend-Tabelle

Die Backend-Tabelle hat die 3 Sheets Email-Verifikation, Buchungen, Reisen.

### Notizen

Für die Buchungen- und Reisentabelle gilt: Steht in der Spalte A einer Tabelle eine Notiz, wird diese Zeile ignoriert. Buchungen mit einer ungültigen IBAN bekommen z.B. eine Notiz "IBAN ungültig". Buchungen für eine Reise deren Anzahl freier Zimmer 0 ist, bekommen eine Notiz "Ausgebucht".
Reisen, die stattgefunden haben, können mit "Abgeschlossen" o.ä. notiert werden. Stornierte Reisen können als "Storniert" notiert werden, usw. Damit soll das Löschen von Zeilen unnötig gemacht werden, damit die Historie einsehbar bleibt.

### Reisen

Im Reisen-Sheet stehen der Name der Reise, einschließlich von-bis, der DZ- und EZ-Preis, die DZ- und EZ-Anzahl, der DZ- und EZ-Rest. Jede Reise (sofern ohne Notiz) erscheint im Formular mit Preisen und freien Zimmern. Der Name der Reise soll natürlich eindeutig sein.

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

Über die "onOpen"-Funktion wird im Backend ein Menü-Eintrag ADFC-MTT erzeugt, mit den Menüpunkten "Anmeldebestätigung senden" und "Update". Wie bei der RFS wird derzeit die Anmeldebestätigung NICHT automatisch verschickt. Vielmehr wählt man eine Zeile im Buchungen-Sheet und sendet dann die Anmeldebestätigung für diese Reise.

Update kann man aufrufen, wenn man in Buchungen oder Reisen irgendwelche Änderungen durchgeführt hat, wie z.B. Stornierungen oder neue Reisen. Damit wird für alle Reisen die Anzahl der freien Zimmer gleich der Anzahl verfügbarer Zimmer minus die Anzahl gebuchter Zimmer gesetzt, und das Formular upgedatet (geupdatet?). Der Aufrufer wird über jede Änderung informiert, um Überraschungen vorzubeugen. Vor allem wird auch vor Überbuchungen gewarnt (Zahl freier Zimmer < 0), die durch Manipulation der Tabellen entstehen könnten.

### Mehrfachbuchungen

Im Formular kann man mehrere Reisen gleichzeitig buchen. Das Skript erstellt für jede gebuchte Reise eine eigene Zeile im Buchungen-Sheet.

### Sortieren

Die Tabellen können beliebig sortiert werden, um z.B. in Buchungen alle Buchungen für eine Reise oder alle Buchungen einer Person zusammenhängend zu sehen. Die Zeile 1 mit den Headern ist fixiert und sollte das auch bleiben (Ansicht/Fixieren/1 Zeile).

### Erste Mail

Nach Abschicken des Formulars erhält der Teilnehmer unter der von ihm angegebenen Email-Adresse eine erste Mail. In dieser steht, was er gebucht hat, ob er vorgemerkt ist oder das Zimmer ausgebucht war, und er wird ggfs. zur Verifikation der Email aufgefordert, falls die Email-Adresse noch unbekannt ist.

### Zweite Mail / Anmeldebestätigung

Diese wird durch den Knopf "Anmeldebestätigung senden" verschickt, und enthält die Reise, EZ oder DZ, und den Preis (bei DZ 2 mal der DZ-Preis), mit der Androhung der baldigen Abbuchung.

### Löschen

Bitte lieber nicht. Lieber in der ersten Spalte eine Notiz setzen.

### Änderungswünsche

Je eher, desto weniger Zeitdruck, desto lieber.

### Fake-IBAN

DE91100000000123456789 ist eine syntaktisch gültige IBAN für Testzwecke.
