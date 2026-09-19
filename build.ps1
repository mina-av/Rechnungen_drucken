# Baut Rechnungen_splitten.exe und Beleg_pruefen.exe neu aus dem aktuellen .py-Stand.
# Eine Aenderung an den .py-Dateien wirkt sich NICHT automatisch auf die bestehenden
# .exe-Dateien aus -- dieses Skript muss nach jeder Aenderung erneut laufen, bevor
# eine neue Version getestet oder an den Kunden geliefert wird.
#
# Voraussetzungen (einmalig): py -m pip install pypdf openpyxl factur-x lxml pikepdf pyinstaller
#
# Aufruf: .\build.ps1

$ErrorActionPreference = "Stop"

$PKGDIR = py -c "import facturx, os; print(os.path.dirname(facturx.__file__))"

py -m PyInstaller --onefile --windowed --icon icon.ico --name "Rechnungen_splitten" `
  --distpath . --noconfirm `
  --add-data "$PKGDIR\xsd_and_schematron;facturx/xsd_and_schematron" `
  --add-data "$PKGDIR\xmp;facturx/xmp" `
  --add-data "sRGB.icc;." `
  rechnungen_splitten.py

py -m PyInstaller --onefile --windowed --icon icon.ico --name "Beleg_pruefen" `
  --distpath . --noconfirm `
  --add-data "$PKGDIR\xsd_and_schematron;facturx/xsd_and_schematron" `
  --add-data "$PKGDIR\xmp;facturx/xmp" `
  --add-data "sRGB.icc;." `
  beleg_pruefen.py

Write-Host ""
Write-Host "Fertig: Rechnungen_splitten.exe und Beleg_pruefen.exe im Projektordner aktualisiert."
Write-Host "Naechste Schritte vor Auslieferung: Code-Signing und Kopie nach 'fuer Kunde/' (siehe README.md)."
