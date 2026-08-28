"""Eigenstaendiges Pruefwerkzeug: zeigt an, ob eine PDF eingebettete ZUGFeRD-Daten enthaelt."""

import tkinter as tk
from tkinter import filedialog, messagebox

from rechnungen_splitten import zugferd_pruefen


def main():
    root = tk.Tk()
    root.withdraw()

    pdf_pfad = filedialog.askopenfilename(
        title="Zu prüfende PDF auswählen",
        filetypes=[("PDF-Dateien", "*.pdf")],
    )
    if not pdf_pfad:
        return

    try:
        ergebnis = zugferd_pruefen(pdf_pfad)
    except Exception as exc:  # noqa: BLE001 - Endnutzer soll Fehlertext sehen, kein Absturz
        messagebox.showerror("Fehler", f"Prüfung fehlgeschlagen:\n{exc}")
        return

    if not ergebnis["gefunden"]:
        messagebox.showwarning(
            "Keine ZUGFeRD-Daten",
            "In dieser PDF wurden keine eingebetteten ZUGFeRD-Daten gefunden.",
        )
        return

    meldung = (
        "ZUGFeRD-Daten gefunden:\n\n"
        f"Rechnungsnummer: {ergebnis['rechnungsnummer']}\n"
        f"Rechnungsdatum:  {ergebnis['rechnungsdatum']}\n"
        f"Käufer:          {ergebnis['kaeufer']}\n"
        f"Gesamtbetrag:    {ergebnis['gesamtbetrag']} EUR"
    )
    messagebox.showinfo("ZUGFeRD-Prüfung", meldung)


if __name__ == "__main__":
    main()
