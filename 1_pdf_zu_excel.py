import re
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Pos. (Zahl oder "o. A.") | Bestell-Nr. | Bezeichnung | Preisgruppe
ZEILE = re.compile(r"^(\d{1,3}|o\.\s?A\.)\s+(\d{5,8})\s+(.+?)\s+(\d{1,4})$")
KOPF = ["Pos.", "Bestell-Nr.", "Bezeichnung", "Preisgruppe", "Datei", "Seite"]

zeilen, gesehen = [], set()

for pdf_pfad in sorted(Path(".").glob("*.pdf")):
    with pdfplumber.open(pdf_pfad) as pdf:
        for seite_nr, seite in enumerate(pdf.pages, start=1):
            for zeile in (seite.extract_text() or "").splitlines():
                treffer = ZEILE.match(zeile.strip())
                if not treffer:
                    continue  # leere Zeilen, Kopfzeilen, Überschriften
                pos, nr, text, preis = treffer.groups()
                eintrag = (pos.replace(" ", " "), nr, text.strip(), int(preis))
                schluessel = (pdf_pfad.name, seite_nr) + eintrag
                if schluessel in gesehen:
                    continue
                gesehen.add(schluessel)
                zeilen.append(list(eintrag) + [pdf_pfad.name, seite_nr])

wb = Workbook()
ws = wb.active
ws.title = "Bestandsliste"
ws.append(KOPF)
for z in zeilen:
    ws.append(z)

for zelle in ws[1]:
    zelle.font = Font(bold=True)
for i, breite in enumerate([8, 14, 50, 13, 30, 7], start=1):
    ws.column_dimensions[get_column_letter(i)].width = breite
ws.freeze_panes = "A2"

wb.save("bestandsliste_roh.xlsx")
print(f"{len(zeilen)} Zeilen gespeichert in bestandsliste_roh.xlsx")
