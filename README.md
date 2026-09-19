# Rechnungen splitten

Windows-Tool für Pagella Lichtenberg GmbH (ambulanter Pflegedienst): teilt eine
CuraSoft-Sammel-PDF (mehrere Rechnungen in einer Datei) automatisch in
Einzelrechnungen auf, benennt sie konsistent und bereitet sie für DATEV vor.

## Funktionsweise

`rechnungen_splitten.py` läuft entweder per Doppelklick (GUI mit Dateiauswahl)
oder als Kommandozeilentool (`py rechnungen_splitten.py <Sammel-PDF>`).

Ablauf pro Lauf:

1. **Seiten gruppieren** — jede Seite wird per Regex auf Rechnungsnummer,
   Rechnungsdatum, Gesamtbetrag, Leistungsempfänger, Abrechnungszeitraum und
   Kostenträger ("abgerechnet für: ...") untersucht. Seiten mit identischer
   Rechnungsnummer werden zusammengehalten (mehrseitige Rechnungen).
2. **Lückenprüfung** — die laufenden Nummern aller Rechnungen eines Laufs
   werden auf Lücken geprüft (Warnung, kein Abbruch).
3. **Einzel-PDF schreiben** — Dateiname:
   `{laufende Nummer}_{Rechnungsdatum}_{Monatscode aus Rechnungsnummer}.pdf`.
4. **ZUGFeRD/Factur-X einbetten** — pro Rechnung wird eine EN16931-konforme
   CII-XML direkt in die PDF eingebettet (Profil `en16931`, `AFRelationship:
   Alternative`), inkl. Käufer (Kostenträger oder Patient, per Debitorenliste
   abgeglichen), Zahlungsbedingungen, PDF/A-3-OutputIntent.
5. **Debitoren-Abgleich** — `Debitoren.xlsx` (DATEV-Debitorenstammdaten-Export,
   liegt neben der EXE, vom Kunden selbst pflegbar) wird bei jedem Lauf neu
   eingelesen. Bei eindeutigem Treffer wird die DATEV-Kurzbezeichnung statt
   des vollen Namens verwendet (PK-Variante bevorzugt, falls PK und KK
   existieren). Kein eindeutiger Treffer → voller Name bleibt, Warnung.
6. **Excel-Übersicht + Log** — pro Lauf, mit Gesamtsumme und rot markierten
   Nummernlücken.

`beleg_pruefen.py` ist ein separates Prüfwerkzeug (nur für den internen
Gebrauch, nicht Teil der Kundenlieferung): zeigt an, ob eine beliebige PDF
eingebettete ZUGFeRD-Daten enthält und welche Werte darin stehen.

## Bewusste Design-Entscheidungen und bekannte Grenzen

- **Sammelposition statt Einzelpositionen.** Die Rechnungsposition in der
  ZUGFeRD-XML ist immer eine einzige Sammelposition über den Gesamtbetrag,
  nicht die einzelnen Leistungszeilen von der Rechnung. Grund: Das
  Rechnungslayout ist je nach Typ (Privat/Kasse/SEPA) unterschiedlich
  aufgebaut: eine zuverlässige zeilenweise Extraktion wäre fehleranfällig,
  und ein falscher Positionsbetrag ist schlimmer als eine fehlende
  Detailaufschlüsselung.
- **Steuerfrei-Annahme.** Alle Rechnungen werden als steuerfrei nach § 4 Nr. 16
  UStG (0 %, Kategorie „E") behandelt — passend zu allen bisher gesehenen
  Pagella-Rechnungen. Bei einer steuerpflichtigen Rechnung wäre das falsch.
- **Käufername unstrukturiert.** Bei Kostenträger-Rechnungen enthält der
  Name-Text bewusst Name+Straße zusammen (z. B. „Techniker Krankenkasse
  Bramfelder Str 140"), außer die Debitorenliste liefert einen eindeutigen
  Treffer. Ein Versuch, Name und Straße per Regex zu trennen, lieferte bei
  echten Daten falsche Ergebnisse (Straßenteile im Namensfeld) und wurde
  wieder verworfen.
- **DATEV-Erkennung unbestätigt.** Die ZUGFeRD-Datei besteht die offizielle
  EN16931-Schematron-Prüfung (inkl. eines gefixten Bibliotheksfehlers, siehe
  unten), es ist aber **nicht bestätigt**, dass DATEV Unternehmen
  online/Belege online die eingebetteten Daten automatisch ausliest und die
  Belegdaten-Maske damit vorbefüllt. Das müsste mit jedem weiteren
  Testupload neu verifiziert werden.
- **Bibliotheks-Workaround (BR-E-05).** Die verwendete `factur-x`-Bibliothek
  prüft beim Erzeugen des Positions-Steuersatzes `if line_dict.get('BT-152')`
  — `Decimal("0")` ist in Python falsy, wodurch der Steuersatz 0 % bei einer
  steuerfreien Position (unser Standardfall) fälschlich weggelassen wurde
  (Verstoß gegen die verpflichtende EN16931-Regel BR-E-05). Wird in
  `cii_xml_zeilensteuersatz_reparieren()` per Nachbearbeitung der erzeugten
  XML behoben.
- **Kein ZIP/`document.xml`-Paket.** Ein früherer Ansatz (DATEV-XML-
  Schnittstelle mit separatem ZIP-Paket + Verwaltungsdatei) wurde verworfen:
  laut offizieller DATEV-Dokumentation akzeptiert der vom Kunden genutzte
  Kanal (`DATEV Upload online`) keine ZIP-Dateien, nur `DATEV Belegtransfer`
  (separates Desktop-Programm, hier nicht vorausgesetzt).

## Build

Voraussetzungen: `pypdf`, `openpyxl`, `factur-x`, `lxml`, `pikepdf` (alle per
`pip install`), PyInstaller.

```powershell
.\build.ps1
```

Baut beide EXE-Dateien (`Rechnungen_splitten.exe`, `Beleg_pruefen.exe`) neu und
legt sie direkt im Projektordner ab. Muss nach jeder Änderung an den `.py`-
Dateien erneut laufen — eine Codeänderung wirkt sich nicht automatisch auf
bestehende EXE-Dateien aus. Die `--add-data`-Einträge im Skript sind nötig,
weil die `factur-x`-Bibliothek ihre XSD-/Schematron-Dateien zur Laufzeit von
der Festplatte lädt — PyInstaller bündelt sie sonst nicht automatisch mit.

Die fertige EXE wird mit einem selbstsignierten Zertifikat signiert (siehe
unten), damit Windows SmartScreen sie nach einmaliger Freigabe bei der IT
nicht bei jedem Update erneut blockiert.

## Auslieferung an den Kunden

Der Ordner `fuer Kunde/` (nicht Teil dieses Repos, siehe `.gitignore`) enthält
nur, was der Kunde braucht: `Rechnungen_splitten.exe`, `Debitoren.xlsx`,
`Arbeitsanweisung.pdf`. `Beleg_pruefen.exe` und alle Testdateien bleiben beim
Entwickler.

**Code-Signing:** Ein selbstsigniertes Zertifikat (`CN=MAV Consulting -
Pagella Rechnungen splitten`, 10 Jahre gültig) wird zum Signieren der EXE
verwendet; das öffentliche Zertifikat (`PagellaRechnungenSplitten.cer`, nicht
im Repo) muss von der Kunden-IT einmalig per Gruppenrichtlinie in
„Vertrauenswürdige Stammzertifizierungsstellen" und „Vertrauenswürdige
Herausgeber" importiert werden. Danach werden künftige, mit demselben
Zertifikat signierte Versionen automatisch akzeptiert.

## Dateien in diesem Repo

| Datei | Zweck |
|---|---|
| `rechnungen_splitten.py` | Hauptprogramm (Split, ZUGFeRD, Debitoren-Abgleich, Excel/Log) |
| `beleg_pruefen.py` | Internes Prüfwerkzeug für eingebettete ZUGFeRD-Daten |
| `icon.ico` | Programm-Icon |
| `sRGB.icc` | Windows-Standard-sRGB-Farbprofil, für das PDF/A-3-OutputIntent |

Bewusst **nicht** im Repo (siehe `.gitignore`): Beispiel-/Test-PDFs mit echten
Patienten-/Versicherungsdaten, `Debitoren.xlsx` (echte DATEV-Kontodaten),
gebaute EXE-/ZIP-Dateien, Zertifikate/Schlüssel, Ausgabeordner.

## Offene Punkte

- Nächster Testupload bei DATEV muss zeigen, ob die ZUGFeRD-Daten jetzt
  automatisch erkannt werden (BR-E-05-Fix, volle vs. 5-stellige
  Rechnungsnummer je nach Kundenwunsch, Debitoren-Kurzbezeichnung).
- Einzelpositionen statt Sammelposition: bewusst zurückgestellt (siehe oben),
  könnte bei Bedarf mit Kontrollsumme gegen den Gesamtbetrag nachgerüstet
  werden.
- Name/Straße-Trennung beim Käufer: nur über die Debitorenliste gelöst, nicht
  generisch.
