"""
Sammel-PDF (CuraSoft-Rechnungsdruck) in einzelne Rechnungen aufteilen.

- Gruppiert Seiten anhand der Rechnungsnummer (mehrseitige Rechnungen bleiben zusammen).
- Benennt jede Rechnung nach dem Schema Rechnungsnummer_Rechnungsdatum_Erstellungsdatum.
- Bettet die Buchungsdaten als ZUGFeRD/Factur-X-XML direkt in die Rechnungs-PDF ein
  (Basic-WL-Profil), sodass eine einzige Datei pro Rechnung entsteht, die DATEV
  automatisch ausliest.
- Prüft die Rechnungsnummern auf Lücken und warnt, verarbeitet aber trotzdem alles.
- Schreibt eine Excel-Übersicht und ein Log pro Lauf.
"""

import calendar
import os
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from facturx import generate_cii_xml, generate_from_file, get_xml_from_pdf
from lxml import etree
import pikepdf

ICC_PROFIL_PFAD = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "sRGB.icc"

# Ordner, in dem die EXE (bzw. das Skript) tatsaechlich liegt -- hier wird die
# vom Kunden gepflegte Debitorenliste erwartet, NICHT im PyInstaller-Temp-Ordner.
if getattr(sys, "frozen", False):
    EXE_ORDNER = Path(sys.executable).parent
else:
    EXE_ORDNER = Path(__file__).parent
DEBITOREN_PFAD = EXE_ORDNER / "Debitoren.xlsx"

DEBITOREN_STOPWOERTER = {
    "ag", "gmbh", "kg", "co", "e", "v", "ev", "und", "vvag", "bkk", "ikk", "pk", "kk",
}

# Feste Stammdaten des Rechnungsstellers (aendert sich nicht pro Rechnung).
VERKAEUFER = {
    "BT-27": "Pagella Lichtenberg GmbH",
    "BT-31": "DE322174884",
    "BT-35": "Küstriner Str. 51",
    "BT-37": "Berlin",
    "BT-38": "13055",
    "BT-40": "DE",
}

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')

RE_NUMMER = re.compile(r"Nr\.:\s*([^\n(]+)")
RE_DATUM = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s*\nZeitraum:")
RE_GESAMTBETRAG = re.compile(r"Gesamtbetrag\s*\n?\s*([\d.,-]+)")
RE_LEISTUNGSEMPFAENGER = re.compile(r"Leistungsempf.nger:\s*\n?\s*([^\n]+)")
RE_LAUFENDE_NUMMER = re.compile(r"(\d+)\s*$")
RE_MONATSCODE = re.compile(r"-(\d{2}\.\d{2})-")
RE_ZEITRAUM = re.compile(r"Zeitraum:\s*(\w+)\s+(\d{4})")
RE_ABGERECHNET_FUER = re.compile(
    r"abgerechnet f.r\s*:\s*\d*\s*\n?(.+?)\s+(\d{5})\s+(\S+)\s+\d{2}\.\d{2}\.\d{4}"
)
RE_ZAHLUNG_WOCHEN = re.compile(r"innerhalb von\s*(\d+)\s*Wochen\s*nach Erhalt der Rechnung", re.IGNORECASE)
RE_ZAHLUNG_LASTSCHRIFT = re.compile(r"wird innerhalb der[^.]*abgebucht", re.IGNORECASE)

MONATSNAMEN = {
    "januar": 1, "februar": 2, "märz": 3, "maerz": 3, "april": 4, "mai": 5,
    "juni": 6, "juli": 7, "august": 8, "september": 9, "oktober": 10,
    "november": 11, "dezember": 12,
}


class Rechnungsgruppe:
    def __init__(self, nummer, first_page):
        self.nummer = nummer
        self.first_page = first_page
        self.last_page = first_page
        self.kundenname = None
        self.rechnungsdatum = None
        self.gesamtbetrag = None
        self.betrag_gefunden = False
        self.zeitraum_monat = None
        self.zeitraum_jahr = None
        self.kostentraeger_name = None
        self.kostentraeger_plz = None
        self.kostentraeger_ort = None
        self.zahlungsfrist_wochen = None
        self.zahlung_lastschrift = False

    def rechnungsempfaenger(self):
        """Der tatsächliche Rechnungsempfänger: Kostenträger (z.B. Krankenkasse),
        falls auf der Rechnung "abgerechnet für" angegeben ist, sonst der Patient."""
        if self.kostentraeger_name:
            return self.kostentraeger_name
        return self.kundenname or "Unbekannt"


def euro_formatieren(betrag):
    """Formatiert einen Betrag im deutschen Format: Punkt als Tausendertrenner,
    Komma als Dezimaltrennzeichen, immer zwei Nachkommastellen (z.B. 3245.8 -> "3.245,80")."""
    return f"{betrag:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def sanitize_dateiname(text):
    return INVALID_FILENAME_CHARS.sub("-", text).strip()


def parse_datum_de(text):
    tag, monat, jahr = text.split(".")
    return f"{jahr}-{monat}-{tag}"


def parse_betrag(text):
    text = text.strip().replace(".", "").replace(",", ".")
    return float(text)


def laufende_nummer_extrahieren(nummer):
    m = RE_LAUFENDE_NUMMER.search(nummer)
    return m.group(1) if m else nummer


def monatscode_aus_nummer(nummer):
    """Extrahiert den MM.JJ-Code aus der Rechnungsnummer, z.B. "08.26" aus "P-08.26-38191"."""
    m = RE_MONATSCODE.search(nummer)
    return m.group(1) if m else date.today().strftime("%m.%y")



def kundenname_kuerzen(rohtext):
    for trenner in ("*", ";"):
        idx = rohtext.find(trenner)
        if idx != -1:
            rohtext = rohtext[:idx]
    return rohtext.strip()


def debitoren_normalisieren(text):
    text = text.lower()
    text = re.sub(r"[^a-z0-9äöüß]+", " ", text)
    return [w for w in text.split() if w not in DEBITOREN_STOPWOERTER]


def debitoren_liste_laden(pfad):
    """
    Laedt die vom Kunden gepflegte Debitorenliste (DATEV-Debitorenstammdaten-Export)
    und baut einen Nachschlage-Index: normalisierte Wortmenge des Organisationsnamens
    (ohne Rechtsform-/Kassentyp-Woerter wie AG, GmbH, PK, KK, BKK) -> Liste der dazu
    passenden Zeilen (Konto, Beschriftung, Kurzbezeichnung, Typ PK/KK).
    Gibt None zurueck, wenn die Datei nicht existiert (Feature dann einfach inaktiv).
    """
    if not pfad.exists():
        return None

    from openpyxl import load_workbook
    wb = load_workbook(pfad, read_only=True, data_only=True)
    ws = wb.active

    index = {}
    for row in ws.iter_rows(values_only=True):
        beschriftung = row[4] if len(row) > 4 else None
        kurzbezeichnung = row[14] if len(row) > 14 else None
        konto = row[2] if len(row) > 2 else None
        if not beschriftung or not isinstance(konto, (int, float)):
            continue
        besch_lower = str(beschriftung).lower().rstrip()
        if besch_lower.endswith("pk"):
            typ = "PK"
        elif besch_lower.endswith("kk"):
            typ = "KK"
        else:
            typ = None
        kern = tuple(sorted(debitoren_normalisieren(str(beschriftung))))
        if not kern:
            continue
        index.setdefault(kern, []).append({
            "typ": typ, "konto": konto,
            "beschriftung": beschriftung, "kurzbezeichnung": kurzbezeichnung,
        })
    wb.close()
    return index


def debitor_nachschlagen(index, kostentraeger_text):
    """
    Sucht zu einem extrahierten Kostentraeger-Text (z.B. "Techniker Krankenkasse
    Bramfelder Str 140") einen eindeutigen Eintrag in der Debitorenliste: Der
    Organisationsname des Eintrags muss vollstaendig als Wortmenge im Text
    enthalten sein. Gibt bei eindeutigem Treffer die Kurzbezeichnung zurueck
    (PK-Variante bevorzugt, falls beide existieren), sonst None.
    """
    if index is None:
        return None
    query_woerter = set(debitoren_normalisieren(kostentraeger_text))
    treffer = [eintraege for kern, eintraege in index.items() if set(kern).issubset(query_woerter)]
    if len(treffer) != 1:
        return None
    eintraege = treffer[0]
    pk_eintrag = next((e for e in eintraege if e["typ"] == "PK"), None)
    gewaehlt = pk_eintrag or eintraege[0]
    kurz = gewaehlt["kurzbezeichnung"]
    return str(kurz).strip() if kurz else None


def leistungsdatum_ermitteln(monatsname, jahr_text):
    monat = MONATSNAMEN.get(monatsname.strip().lower())
    if not monat:
        return None
    jahr = int(jahr_text)
    letzter_tag = calendar.monthrange(jahr, monat)[1]
    return f"{jahr:04d}-{monat:02d}-{letzter_tag:02d}"


def seiten_auslesen(reader):
    seiten = []
    for i, seite in enumerate(reader.pages):
        text = seite.extract_text()
        m_nummer = RE_NUMMER.search(text)
        m_datum = RE_DATUM.search(text)
        m_betrag = RE_GESAMTBETRAG.search(text)
        m_kunde = RE_LEISTUNGSEMPFAENGER.search(text)
        m_zeitraum = RE_ZEITRAUM.search(text)
        m_kostentraeger = RE_ABGERECHNET_FUER.search(text)
        m_zahlung_wochen = RE_ZAHLUNG_WOCHEN.search(text)
        m_lastschrift = RE_ZAHLUNG_LASTSCHRIFT.search(text)
        seiten.append({
            "index": i,
            "nummer": m_nummer.group(1).strip() if m_nummer else None,
            "datum": m_datum.group(1) if m_datum else None,
            "gesamtbetrag": m_betrag.group(1) if m_betrag else None,
            "kundenname": kundenname_kuerzen(m_kunde.group(1)) if m_kunde else None,
            "zeitraum_monat": m_zeitraum.group(1) if m_zeitraum else None,
            "zeitraum_jahr": m_zeitraum.group(2) if m_zeitraum else None,
            "kostentraeger_name": m_kostentraeger.group(1).strip() if m_kostentraeger else None,
            "kostentraeger_plz": m_kostentraeger.group(2) if m_kostentraeger else None,
            "kostentraeger_ort": m_kostentraeger.group(3) if m_kostentraeger else None,
            "zahlungsfrist_wochen": int(m_zahlung_wochen.group(1)) if m_zahlung_wochen else None,
            "zahlung_lastschrift": bool(m_lastschrift),
        })
    return seiten


def seiten_gruppieren(seiten, warnungen):
    gruppen = []
    aktuelle = None

    for seite in seiten:
        if seite["nummer"] is None:
            if aktuelle is None:
                warnungen.append(
                    f"Seite {seite['index'] + 1}: keine Rechnungsnummer erkannt und "
                    "keine vorherige Rechnung zum Anhängen vorhanden – Seite wird ignoriert."
                )
                continue
            warnungen.append(
                f"Seite {seite['index'] + 1}: keine eigene Kopfzeile erkannt, "
                f"wird als Fortsetzung von Rechnung {aktuelle.nummer} behandelt."
            )
            aktuelle.last_page = seite["index"]
        elif aktuelle is not None and seite["nummer"] == aktuelle.nummer:
            aktuelle.last_page = seite["index"]
        else:
            aktuelle = Rechnungsgruppe(seite["nummer"], seite["index"])
            gruppen.append(aktuelle)

        if seite["datum"]:
            aktuelle.rechnungsdatum = seite["datum"]
        if seite["kundenname"]:
            aktuelle.kundenname = seite["kundenname"]
        if seite["gesamtbetrag"]:
            aktuelle.gesamtbetrag = seite["gesamtbetrag"]
            aktuelle.betrag_gefunden = True
        if seite["zeitraum_monat"] and seite["zeitraum_jahr"]:
            aktuelle.zeitraum_monat = seite["zeitraum_monat"]
            aktuelle.zeitraum_jahr = seite["zeitraum_jahr"]
        if seite["kostentraeger_name"]:
            aktuelle.kostentraeger_name = seite["kostentraeger_name"]
            aktuelle.kostentraeger_plz = seite["kostentraeger_plz"]
            aktuelle.kostentraeger_ort = seite["kostentraeger_ort"]
        if seite["zahlungsfrist_wochen"]:
            aktuelle.zahlungsfrist_wochen = seite["zahlungsfrist_wochen"]
        if seite["zahlung_lastschrift"]:
            aktuelle.zahlung_lastschrift = True

    return gruppen


def luecken_pruefen(gruppen, warnungen):
    nummern = []
    for g in gruppen:
        m = RE_LAUFENDE_NUMMER.search(g.nummer)
        if m:
            nummern.append(int(m.group(1)))

    if len(nummern) < 2:
        return []

    nummern_sortiert = sorted(nummern)
    fehlende = []
    for a, b in zip(nummern_sortiert, nummern_sortiert[1:]):
        if b - a > 1:
            fehlende.extend(range(a + 1, b))

    if fehlende:
        warnungen.append(
            "Lücke(n) in den Rechnungsnummern gefunden, fehlende laufende Nummer(n): "
            + ", ".join(str(n) for n in fehlende)
        )
    return fehlende


def zugferd_daten_erzeugen(gruppe, rechnungsdatum, leistungsdatum_iso, betrag, laufende_nummer, debitor_index):
    """
    Baut das Datenwoerterbuch (EN16931-Business-Terms) fuer eine ZUGFeRD/Factur-X-
    Rechnung im Profil EN16931 (mit einer zusammenfassenden Rechnungsposition).
    Nimmt an, dass die Leistung nach Paragraf 4 Nr. 16 UStG steuerfrei ist
    (Umsatzsteuerkategorie "E", 0%) -- passend zu allen bisher gesehenen
    Pagella-Rechnungen. Bei einer steuerpflichtigen Rechnung waere diese
    Annahme falsch und muesste angepasst werden.

    Als Rechnungsempfaenger (Kaeufer) wird bevorzugt die Kurzbezeichnung aus der
    Debitorenliste des Kunden verwendet (eindeutiger Treffer vorausgesetzt),
    sonst der auf der Rechnung gefundene Kostentraeger (z.B. Krankenkasse) bzw.
    der Patient. Es wird bewusst nur eine zusammenfassende Rechnungsposition
    gebildet statt der einzelnen Leistungszeilen, da deren zuverlaessige
    Extraktion aus dem unterschiedlich aufgebauten Rechnungslayout zu
    fehleranfaellig waere -- ein falscher Positionsbetrag waere schlimmer als
    eine fehlende Detailaufschluesselung.
    """
    betrag_dec = Decimal(str(betrag))
    leistungsdatum_dt = date.fromisoformat(leistungsdatum_iso) if leistungsdatum_iso else rechnungsdatum

    empfaenger_voll = gruppe.rechnungsempfaenger()
    debitor_kurz = debitor_nachschlagen(debitor_index, empfaenger_voll) if gruppe.kostentraeger_name else None
    kaeufer_name = debitor_kurz or empfaenger_voll

    data = {
        "BT-1": laufende_nummer,
        "BT-2": rechnungsdatum,
        "BT-3": "380",
        "BT-5": "EUR",
        "BT-72": leistungsdatum_dt,
        "BT-44": kaeufer_name,
        "BT-55": "DE",
        "BG-23": [{
            "BT-116": betrag_dec, "BT-116-1": "EUR",
            "BT-117": Decimal("0.00"), "BT-117-1": "EUR",
            "BT-118": "E",
            "BT-120": "Steuerfrei nach Paragraf 4 Nr. 16 UStG",
        }],
        "BT-106": betrag_dec, "BT-106-1": "EUR",
        "BT-109": betrag_dec, "BT-109-1": "EUR",
        "BT-110": Decimal("0.00"), "BT-110-1": "EUR",
        "BT-112": betrag_dec, "BT-112-1": "EUR",
        "BT-115": betrag_dec, "BT-115-1": "EUR",
        "BG-25": [{
            "BT-126": "1",
            "BT-129": Decimal("1"),
            "BT-130": "C62",
            "BT-131": betrag_dec, "BT-131-1": "EUR",
            "BT-153": "Pflegeleistungen gemäß Rechnung",
            "BT-146": betrag_dec, "BT-146-1": "EUR",
            "BT-151": "E",
            "BT-152": Decimal("0"),
        }],
    }
    if gruppe.kostentraeger_plz and gruppe.kostentraeger_ort:
        data["BT-53"] = gruppe.kostentraeger_plz
        data["BT-52"] = gruppe.kostentraeger_ort

    if gruppe.zahlungsfrist_wochen:
        data["BT-9"] = rechnungsdatum + timedelta(weeks=gruppe.zahlungsfrist_wochen)
        data["BT-20"] = (
            f"Zahlbar innerhalb von {gruppe.zahlungsfrist_wochen} Wochen nach Erhalt der Rechnung."
        )
    elif gruppe.zahlung_lastschrift:
        data["BT-20"] = "Der Rechnungsbetrag wird per Lastschrift eingezogen."

    data.update(VERKAEUFER)
    return data


def pdf_outputintent_ergaenzen(pdf_pfad):
    """
    Ergaenzt ein PDF/A-OutputIntent (sRGB-Farbprofil) im Dokumentkatalog.
    Die facturx-Bibliothek behauptet in den XMP-Metadaten PDF/A-3-Konformitaet,
    ergaenzt aber selbst kein OutputIntent -- ohne das waere die Behauptung falsch,
    was strenge PDF/A-Pruefungen (evtl. auch bei DATEV) stoeren kann.
    """
    with pikepdf.open(pdf_pfad, allow_overwriting_input=True) as pdf:
        if "/OutputIntents" in pdf.Root:
            return
        icc_bytes = ICC_PROFIL_PFAD.read_bytes()
        icc_stream = pikepdf.Stream(pdf, icc_bytes)
        icc_stream["/N"] = 3
        icc_stream["/Alternate"] = pikepdf.Name("/DeviceRGB")
        output_intent = pdf.make_indirect(pikepdf.Dictionary(
            Type=pikepdf.Name("/OutputIntent"),
            S=pikepdf.Name("/GTS_PDFA1"),
            OutputConditionIdentifier="sRGB IEC61966-2.1",
            Info="sRGB IEC61966-2.1",
            DestOutputProfile=icc_stream,
        ))
        pdf.Root.OutputIntents = pikepdf.Array([output_intent])
        pdf.save(pdf_pfad)


def pdf_annotationen_entfernen(pdf_pfad):
    """Entfernt Anmerkungen (z.B. Stempel/Markierungen aus der CuraSoft-Quelldatei),
    die bei einer strengen PDF/A-Pruefung stoeren koennen."""
    reader = PdfReader(pdf_pfad)
    veraendert = False
    for seite in reader.pages:
        if "/Annots" in seite:
            del seite["/Annots"]
            veraendert = True
    if veraendert:
        writer = PdfWriter()
        for seite in reader.pages:
            writer.add_page(seite)
        with open(pdf_pfad, "wb") as f:
            writer.write(f)


RAM_NS = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100"


def cii_xml_zeilensteuersatz_reparieren(cii_xml_bytes):
    """
    Workaround fuer einen Bug in der facturx-Bibliothek: Sie prueft beim
    Erzeugen von RateApplicablePercent auf einer Rechnungsposition mit
    "if line_dict.get('BT-152')" -- Decimal("0") ist in Python aber falsy,
    wodurch das Feld bei 0% Steuersatz (unser Steuerfrei-Fall) faelschlich
    weggelassen wird. Das verletzt die verpflichtende EN16931-Regel BR-E-05.
    Ergaenzt das fehlende <ram:RateApplicablePercent>0</ram:RateApplicablePercent>
    nachtraeglich in jeder ApplicableTradeTax einer Rechnungsposition, die es
    noch nicht hat.
    """
    root = etree.fromstring(cii_xml_bytes)
    ns = {"ram": RAM_NS}
    for tax in root.xpath(
        "//ram:IncludedSupplyChainTradeLineItem/ram:SpecifiedLineTradeSettlement"
        "/ram:ApplicableTradeTax",
        namespaces=ns,
    ):
        if tax.find("ram:RateApplicablePercent", namespaces=ns) is None:
            rate = etree.SubElement(tax, f"{{{RAM_NS}}}RateApplicablePercent")
            rate.text = "0"
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def zugferd_pdf_erzeugen(pdf_pfad, gruppe, rechnungsdatum, leistungsdatum_iso, betrag, laufende_nummer, debitor_index):
    """Bettet die ZUGFeRD-XML direkt in die angegebene PDF-Datei ein (in-place)."""
    pdf_annotationen_entfernen(pdf_pfad)
    data = zugferd_daten_erzeugen(gruppe, rechnungsdatum, leistungsdatum_iso, betrag, laufende_nummer, debitor_index)
    cii_xml = generate_cii_xml(data, level="en16931")
    cii_xml = cii_xml_zeilensteuersatz_reparieren(cii_xml)
    generate_from_file(
        str(pdf_pfad), cii_xml, flavor="factur-x", level="en16931", check_xsd=True,
        afrelationship="alternative",
    )
    pdf_outputintent_ergaenzen(pdf_pfad)


CII_NS = {
    "rsm": "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100",
    "ram": "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100",
    "udt": "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100",
}


def zugferd_pruefen(pdf_pfad):
    """
    Prueft, ob eine PDF eingebettete ZUGFeRD/Factur-X-Daten enthaelt.
    Gibt ein dict zurueck: {"gefunden": bool, ...Kernwerte, falls gefunden}.
    """
    with open(pdf_pfad, "rb") as f:
        pdf_bytes = f.read()

    try:
        xml_name, xml_bytes = get_xml_from_pdf(pdf_bytes)
    except Exception:
        xml_name, xml_bytes = None, None

    if not xml_bytes:
        return {"gefunden": False}

    root = etree.fromstring(xml_bytes)

    def xpath_text(pfad_ausdruck):
        ergebnis = root.xpath(pfad_ausdruck, namespaces=CII_NS)
        return ergebnis[0].text if ergebnis else None

    return {
        "gefunden": True,
        "dateiname_anhang": xml_name,
        "rechnungsnummer": xpath_text("//rsm:ExchangedDocument/ram:ID"),
        "rechnungsdatum": xpath_text(
            "//rsm:ExchangedDocument/ram:IssueDateTime/udt:DateTimeString"
        ),
        "kaeufer": xpath_text(
            "//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:Name"
        ),
        "gesamtbetrag": xpath_text(
            "//ram:SpecifiedTradeSettlementHeaderMonetarySummation/ram:GrandTotalAmount"
        ),
    }


def excel_uebersicht_schreiben(pfad, rechnungen, fehlende_nummern):
    wb = Workbook()
    ws = wb.active
    ws.title = "Übersicht"

    header = ["Rechnungsnummer", "Rechnungsdatum", "Kundenname", "Gesamtbetrag", "Seitenzahl"]
    ws.append(header)
    for zelle in ws[1]:
        zelle.font = Font(bold=True)

    gesamtsumme = 0.0
    for r in rechnungen:
        ws.append([
            r["nummer"], r["rechnungsdatum_iso"], r["kundenname"] or "",
            r["betrag"], r["seitenzahl"],
        ])
        if r["betrag"] is not None:
            gesamtsumme += r["betrag"]

    ws.append([])
    summenzeile = ws.max_row + 1
    ws.append(["Anzahl Rechnungen", len(rechnungen), "", "Gesamtsumme", round(gesamtsumme, 2)])
    for zelle in ws[summenzeile]:
        zelle.font = Font(bold=True)

    if fehlende_nummern:
        ws.append([])
        rot = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append(["Fehlende laufende Rechnungsnummern:"])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        for n in fehlende_nummern:
            ws.append([n])
            ws.cell(row=ws.max_row, column=1).fill = rot

    for spalte, breite in zip("ABCDE", (20, 16, 30, 14, 12)):
        ws.column_dimensions[spalte].width = breite

    wb.save(pfad)


def verarbeite(pdf_pfad, ziel_ordner=None):
    """
    ziel_ordner: Ordner, in dem der Lauf-Unterordner direkt angelegt wird.
    Wenn nicht angegeben (z. B. CLI-Aufruf), wird automatisch ein
    "ausgabe"-Ordner neben der Sammel-PDF verwendet.
    """
    pdf_pfad = Path(pdf_pfad)
    if ziel_ordner:
        ziel_ordner = Path(ziel_ordner)
    else:
        ziel_ordner = pdf_pfad.parent / "ausgabe"

    heute_iso = date.today().isoformat()
    zeitstempel = datetime.now().strftime("%H-%M-%S")

    lauf_ordnername = sanitize_dateiname(f"{pdf_pfad.stem}_{heute_iso}_{zeitstempel}")
    ausgabe_ordner = ziel_ordner / lauf_ordnername
    laufzaehler = 2
    while ausgabe_ordner.exists():
        # Zwei Laeufe innerhalb derselben Sekunde (z.B. schnell aufeinanderfolgende
        # Klicks) duerfen niemals denselben Ordner treffen -- sonst ueberschreibt
        # der zweite Lauf lautlos die Rechnungen des ersten.
        ausgabe_ordner = ziel_ordner / f"{lauf_ordnername}_{laufzaehler}"
        laufzaehler += 1
    ausgabe_ordner.mkdir(parents=True)

    warnungen = []

    debitor_index = debitoren_liste_laden(DEBITOREN_PFAD)
    if debitor_index is None:
        warnungen.append(
            f"Debitorenliste nicht gefunden ({DEBITOREN_PFAD.name} neben dem Programm) "
            "- Kostenträger-Namen werden unverändert übernommen."
        )

    reader = PdfReader(pdf_pfad)
    seiten = seiten_auslesen(reader)
    gruppen = seiten_gruppieren(seiten, warnungen)
    fehlende_nummern = luecken_pruefen(gruppen, warnungen)

    rechnungen_fuer_excel = []

    for gruppe in gruppen:
        if not gruppe.rechnungsdatum:
            warnungen.append(f"Rechnung {gruppe.nummer}: kein Rechnungsdatum gefunden.")
            rechnungsdatum_iso = "unbekannt"
        else:
            rechnungsdatum_iso = parse_datum_de(gruppe.rechnungsdatum)

        if not gruppe.betrag_gefunden:
            warnungen.append(f"Rechnung {gruppe.nummer}: kein Gesamtbetrag gefunden.")
            betrag = None
        else:
            betrag = parse_betrag(gruppe.gesamtbetrag)

        laufende_nummer = laufende_nummer_extrahieren(gruppe.nummer)
        rechnungsdatum_anzeige = gruppe.rechnungsdatum or "unbekannt"
        erstellungsdatum_kurz = monatscode_aus_nummer(gruppe.nummer)
        basisname = sanitize_dateiname(
            f"{laufende_nummer}_{rechnungsdatum_anzeige}_{erstellungsdatum_kurz}"
        )

        if gruppe.zeitraum_monat and gruppe.zeitraum_jahr:
            leistungsdatum_iso = leistungsdatum_ermitteln(gruppe.zeitraum_monat, gruppe.zeitraum_jahr)
        else:
            leistungsdatum_iso = None
            warnungen.append(f"Rechnung {gruppe.nummer}: kein Zeitraum/Leistungsdatum gefunden.")

        # PDF-Auszug schreiben
        writer = PdfWriter()
        for seiten_index in range(gruppe.first_page, gruppe.last_page + 1):
            writer.add_page(reader.pages[seiten_index])
        pdf_zielpfad = ausgabe_ordner / f"{basisname}.pdf"
        with open(pdf_zielpfad, "wb") as f:
            writer.write(f)

        if gruppe.kostentraeger_name and debitor_index is not None:
            if not debitor_nachschlagen(debitor_index, gruppe.rechnungsempfaenger()):
                warnungen.append(
                    f"Rechnung {gruppe.nummer}: kein eindeutiger Debitor für "
                    f"\"{gruppe.kostentraeger_name}\" in der Debitorenliste gefunden - "
                    "voller Name wird verwendet, bitte manuell prüfen."
                )

        # ZUGFeRD-Daten direkt in die PDF einbetten
        if betrag is not None and rechnungsdatum_iso != "unbekannt":
            try:
                zugferd_pdf_erzeugen(
                    pdf_zielpfad, gruppe, date.fromisoformat(rechnungsdatum_iso),
                    leistungsdatum_iso, betrag, laufende_nummer, debitor_index,
                )
            except Exception as exc:  # noqa: BLE001 - PDF bleibt nutzbar, nur ohne eingebettete Daten
                warnungen.append(f"Rechnung {gruppe.nummer}: ZUGFeRD-Daten konnten nicht eingebettet werden ({exc}).")
        else:
            warnungen.append(f"Rechnung {gruppe.nummer}: ZUGFeRD-Daten nicht eingebettet (Rechnungsdatum oder Betrag fehlt).")

        rechnungen_fuer_excel.append({
            "nummer": gruppe.nummer,
            "rechnungsdatum_iso": rechnungsdatum_iso,
            "kundenname": gruppe.kundenname,
            "betrag": betrag,
            "seitenzahl": gruppe.last_page - gruppe.first_page + 1,
            "basisname": basisname,
        })


    gesamtsumme = round(sum(r["betrag"] for r in rechnungen_fuer_excel if r["betrag"] is not None), 2)

    excel_pfad = ausgabe_ordner / f"uebersicht_{heute_iso}.xlsx"
    excel_uebersicht_schreiben(excel_pfad, rechnungen_fuer_excel, fehlende_nummern)

    log_zeilen = [
        f"Lauf am {heute_iso}",
        f"Quelldatei: {pdf_pfad.name}",
        f"Seiten insgesamt: {len(seiten)}",
        f"Erkannte Rechnungen: {len(gruppen)}",
        "",
        "Warnungen:" if warnungen else "Keine Warnungen.",
    ] + [f"- {w}" for w in warnungen]
    log_pfad = ausgabe_ordner / f"log_{heute_iso}.txt"
    log_pfad.write_text("\n".join(log_zeilen), encoding="utf-8")

    print("\n".join(log_zeilen))
    print(f"\nExcel-Übersicht: {excel_pfad}")
    print(f"Ausgabeordner: {ausgabe_ordner}")

    return {
        "gruppen": len(gruppen),
        "warnungen": warnungen,
        "ausgabe_ordner": ausgabe_ordner,
        "excel_pfad": excel_pfad,
        "gesamtsumme": gesamtsumme,
    }


def gui_main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    FARBE_MARKE = "#1c3d5a"
    FARBE_HINTERGRUND = "#f4f6f8"
    FARBE_AKZENT = "#2563eb"
    FARBE_AKZENT_HOVER = "#1d4fc4"

    root = tk.Tk()
    root.title("Rechnungen aufteilen")
    root.configure(bg=FARBE_HINTERGRUND)

    stil = ttk.Style(root)
    stil.theme_use("clam")
    stil.configure("TFrame", background=FARBE_HINTERGRUND)
    stil.configure("TLabel", background=FARBE_HINTERGRUND, font=("Segoe UI", 9))
    stil.configure("Schritt.TLabel", background=FARBE_HINTERGRUND, foreground=FARBE_MARKE, font=("Segoe UI", 9, "bold"))
    stil.configure("Hinweis.TLabel", background=FARBE_HINTERGRUND, foreground="#5a6774", font=("Segoe UI", 8))
    stil.configure("Kopf.TFrame", background=FARBE_MARKE)
    stil.configure("Kopf.TLabel", background=FARBE_MARKE, foreground="white", font=("Segoe UI", 13, "bold"))
    stil.configure("Akzent.TButton", background=FARBE_AKZENT, foreground="white", font=("Segoe UI", 9, "bold"), padding=6)
    stil.map("Akzent.TButton",
             background=[("disabled", "#a9b4c2"), ("active", FARBE_AKZENT_HOVER)],
             foreground=[("disabled", "#eef1f5")])
    stil.configure("ErgebnisFeld.TLabel", background=FARBE_HINTERGRUND, font=("Segoe UI", 9, "bold"))
    stil.configure("ErgebnisWert.TLabel", background=FARBE_HINTERGRUND, font=("Segoe UI", 11, "bold"), foreground=FARBE_MARKE)

    # Zielordner wird beim PDF-Wechsel automatisch vorgeschlagen, solange die
    # Person das Feld nicht selbst manuell geaendert hat.
    ziel_ordner_manuell_geaendert = False
    ausgabe_ergebnis = {}

    aussenabstand = {"padx": 12, "pady": 6}

    kopf = ttk.Frame(root, style="Kopf.TFrame")
    kopf.grid(row=0, column=0, sticky="we")
    ttk.Label(kopf, text="Rechnungen aufteilen", style="Kopf.TLabel").pack(padx=14, pady=10, anchor="w")

    rahmen = ttk.Frame(root, padding=12)
    rahmen.grid(row=1, column=0, sticky="nsew")

    ttk.Label(rahmen, text="1. Sammel-PDF auswählen", style="Schritt.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
    ttk.Label(
        rahmen, text="Die CuraSoft-Sammeldatei mit mehreren Rechnungen, die aufgeteilt werden soll.",
        style="Hinweis.TLabel",
    ).grid(row=1, column=0, columnspan=2, sticky="w", padx=(0, 0))
    pdf_var = tk.StringVar()
    pdf_feld = ttk.Entry(rahmen, textvariable=pdf_var, width=55, state="readonly")
    pdf_feld.grid(row=2, column=0, sticky="we", **aussenabstand)

    ttk.Label(rahmen, text="2. Zielordner", style="Schritt.TLabel").grid(row=3, column=0, columnspan=2, sticky="w")
    ttk.Label(
        rahmen, text="Hier legt das Programm die Einzelrechnungen, die Excel-Übersicht und das Log ab.",
        style="Hinweis.TLabel",
    ).grid(row=4, column=0, columnspan=2, sticky="w")
    ziel_var = tk.StringVar()
    ziel_feld = ttk.Entry(rahmen, textvariable=ziel_var, width=55)
    ziel_feld.grid(row=5, column=0, sticky="we", **aussenabstand)

    start_knopf = ttk.Button(rahmen, text="Rechnungen aufteilen", state="disabled", style="Akzent.TButton")
    start_knopf.grid(row=6, column=0, columnspan=2, pady=(4, 10))

    trenner = ttk.Separator(rahmen, orient="horizontal")
    trenner.grid(row=7, column=0, columnspan=2, sticky="we", pady=(0, 6))

    status_var = tk.StringVar(value="Noch keine Rechnungen aufgeteilt.")
    ttk.Label(rahmen, textvariable=status_var, style="Hinweis.TLabel").grid(row=8, column=0, columnspan=2, sticky="w")

    ergebnis_werte = ttk.Frame(rahmen)
    ergebnis_werte.grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))
    ttk.Label(ergebnis_werte, text="Anzahl Rechnungen:", style="ErgebnisFeld.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
    anzahl_var = tk.StringVar(value="–")
    ttk.Label(ergebnis_werte, textvariable=anzahl_var, style="ErgebnisWert.TLabel").grid(row=0, column=1, sticky="w")
    ttk.Label(ergebnis_werte, text="Gesamtsumme:", style="ErgebnisFeld.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(2, 0))
    summe_var = tk.StringVar(value="–")
    ttk.Label(ergebnis_werte, textvariable=summe_var, style="ErgebnisWert.TLabel").grid(row=1, column=1, sticky="w", pady=(2, 0))

    ttk.Label(rahmen, text="Anmerkungen:", style="ErgebnisFeld.TLabel").grid(row=10, column=0, columnspan=2, sticky="w", pady=(8, 2))
    anmerkungen_text = tk.Text(
        rahmen, width=64, height=7, state="disabled", wrap="word",
        bg="white", relief="solid", borderwidth=1, highlightthickness=0,
    )
    anmerkungen_text.grid(row=11, column=0, columnspan=2, sticky="we", padx=12)
    anmerkungen_text.tag_configure("ok", foreground="#1a7a1a")
    anmerkungen_text.tag_configure("warnung", foreground="#b35900")

    knopf_leiste = ttk.Frame(rahmen)
    knopf_leiste.grid(row=12, column=0, columnspan=2, pady=(8, 0))
    ordner_knopf = ttk.Button(knopf_leiste, text="Ausgabeordner öffnen", state="disabled")
    ordner_knopf.grid(row=0, column=0, padx=4)
    excel_knopf = ttk.Button(knopf_leiste, text="Excel öffnen", state="disabled")
    excel_knopf.grid(row=0, column=1, padx=4)
    loeschen_knopf = ttk.Button(knopf_leiste, text="Sammel-PDF löschen", state="disabled")
    loeschen_knopf.grid(row=0, column=2, padx=4)

    def anmerkungen_anzeigen(zeilen_mit_tags):
        anmerkungen_text.config(state="normal")
        anmerkungen_text.delete("1.0", "end")
        for text, tag in zeilen_mit_tags:
            anmerkungen_text.insert("end", text, tag)
        anmerkungen_text.config(state="disabled")

    def datei_oeffnen(pfad):
        try:
            os.startfile(pfad)  # noqa: S606 - gewuenschtes Verhalten fuer Endnutzer
        except OSError as exc:
            messagebox.showerror("Fehler", f"Konnte nicht geöffnet werden:\n{exc}")

    def sammel_pdf_loeschen():
        pfad = ausgabe_ergebnis.get("quelle")
        if not pfad:
            return
        bestaetigt = messagebox.askyesno(
            "Sammel-PDF löschen",
            f"Soll die Ausgangsdatei wirklich unwiderruflich gelöscht werden?\n\n{pfad}",
            icon="warning",
        )
        if not bestaetigt:
            return
        try:
            Path(pfad).unlink()
        except OSError as exc:
            messagebox.showerror("Fehler", f"Datei konnte nicht gelöscht werden:\n{exc}")
            return
        loeschen_knopf.config(state="disabled")
        status_var.set("Ausgangsdatei gelöscht.")

    def pdf_waehlen():
        nonlocal ziel_ordner_manuell_geaendert
        pfad = filedialog.askopenfilename(
            title="Sammel-PDF auswählen",
            filetypes=[("PDF-Dateien", "*.pdf")],
        )
        if not pfad:
            return
        loeschen_knopf.config(state="disabled")
        pdf_var.set(pfad)
        if not ziel_ordner_manuell_geaendert:
            ziel_var.set(str(Path(pfad).parent / "ausgabe"))
        start_knopf.config(state="normal")

    def ziel_waehlen():
        nonlocal ziel_ordner_manuell_geaendert
        start = ziel_var.get() or str(Path(pdf_var.get()).parent)
        pfad = filedialog.askdirectory(
            title="In welchem Ordner sollen die Ergebnisse abgelegt werden?",
            initialdir=start,
        )
        if pfad:
            ziel_var.set(pfad)
            ziel_ordner_manuell_geaendert = True

    def ziel_manuell_editiert(_event):
        nonlocal ziel_ordner_manuell_geaendert
        ziel_ordner_manuell_geaendert = True

    def aufteilen_starten():
        pdf_pfad = pdf_var.get()
        ziel_ordner = ziel_var.get()
        if not pdf_pfad or not ziel_ordner:
            return

        start_knopf.config(state="disabled")
        ordner_knopf.config(state="disabled")
        excel_knopf.config(state="disabled")
        loeschen_knopf.config(state="disabled")
        status_var.set("Verarbeite ...")
        anzahl_var.set("–")
        summe_var.set("–")
        anmerkungen_anzeigen([("Bitte warten ...", None)])
        root.update()

        try:
            ergebnis = verarbeite(pdf_pfad, ziel_ordner=ziel_ordner)
        except Exception as exc:  # noqa: BLE001 - Endnutzer soll Fehlertext sehen, kein Absturz
            status_var.set("Verarbeitung fehlgeschlagen.")
            anmerkungen_anzeigen([(str(exc), "warnung")])
            start_knopf.config(state="normal")
            return

        ausgabe_ergebnis["ordner"] = ergebnis["ausgabe_ordner"]
        ausgabe_ergebnis["excel"] = ergebnis["excel_pfad"]
        ausgabe_ergebnis["quelle"] = pdf_pfad

        status_var.set("Fertig.")
        anzahl_var.set(str(ergebnis["gruppen"]))
        summe_var.set(euro_formatieren(ergebnis["gesamtsumme"]) + " EUR")

        if ergebnis["warnungen"]:
            zeilen = [(f"{len(ergebnis['warnungen'])} Anmerkung(en):\n", "warnung")]
            zeilen.extend((f"  • {w}\n", "warnung") for w in ergebnis["warnungen"])
        else:
            zeilen = [("Keine Anmerkungen.", "ok")]
        anmerkungen_anzeigen(zeilen)

        start_knopf.config(state="normal")
        ordner_knopf.config(state="normal")
        excel_knopf.config(state="normal")
        loeschen_knopf.config(state="normal")

    start_knopf.config(command=aufteilen_starten)
    ordner_knopf.config(command=lambda: datei_oeffnen(ausgabe_ergebnis["ordner"]))
    excel_knopf.config(command=lambda: datei_oeffnen(ausgabe_ergebnis["excel"]))
    loeschen_knopf.config(command=sammel_pdf_loeschen)

    ttk.Button(rahmen, text="Durchsuchen", command=pdf_waehlen).grid(row=2, column=1, padx=(0, 12))
    ttk.Button(rahmen, text="Durchsuchen", command=ziel_waehlen).grid(row=5, column=1, padx=(0, 12))
    ziel_feld.bind("<Key>", ziel_manuell_editiert)

    root.mainloop()


def main():
    if len(sys.argv) == 3 and sys.argv[1] == "--pruefen":
        ergebnis = zugferd_pruefen(sys.argv[2])
        if not ergebnis["gefunden"]:
            print("Keine ZUGFeRD-Daten in dieser PDF gefunden.")
        else:
            print("ZUGFeRD-Daten gefunden:")
            print(f"  Rechnungsnummer: {ergebnis['rechnungsnummer']}")
            print(f"  Rechnungsdatum:  {ergebnis['rechnungsdatum']}")
            print(f"  Käufer:          {ergebnis['kaeufer']}")
            print(f"  Gesamtbetrag:    {ergebnis['gesamtbetrag']} EUR")
    elif len(sys.argv) == 2:
        verarbeite(sys.argv[1])
    elif len(sys.argv) == 1:
        gui_main()
    else:
        print("Aufruf: py rechnungen_splitten.py [<Sammel-PDF>] | --pruefen <PDF>")
        sys.exit(1)


if __name__ == "__main__":
    main()
