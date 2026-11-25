from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

MAX_PRO_FOLIE = 15
SPALTEN = ["Pos.", "Bestell-Nr.", "Bezeichnung", "Preisgruppe"]
BREITEN = [Inches(1.2), Inches(2.2), Inches(7.0), Inches(1.9)]  # = 12,3"

# Daten einlesen (leere Zeilen überspringen), nur die ersten 4 Spalten
ws = load_workbook("bestandsliste_fertig.xlsx", data_only=True).active
zeilen = [
    ["" if c is None else str(c) for c in r[:4]]
    for r in ws.iter_rows(min_row=2, values_only=True)
    if any(c not in (None, "") for c in r[:4])
]

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)  # 16:9
leer = prs.slide_layouts[6]

seiten = [zeilen[i:i + MAX_PRO_FOLIE] for i in range(0, len(zeilen), MAX_PRO_FOLIE)]


def zelle_setzen(zelle, text, fett=False, farbe=None, fuellung=None, mitte=False):
    zelle.text = text
    p = zelle.text_frame.paragraphs[0]
    p.font.size = Pt(13)
    p.font.bold = fett
    if mitte:
        p.alignment = PP_ALIGN.CENTER
    if farbe:
        p.font.color.rgb = farbe
    if fuellung:
        zelle.fill.solid()
        zelle.fill.fore_color.rgb = fuellung


for nr, block in enumerate(seiten, start=1):
    folie = prs.slides.add_slide(leer)

    titel = folie.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.6))
    titel.text_frame.text = f"Ersatzteilliste (Seite {nr}/{len(seiten)})"
    titel.text_frame.paragraphs[0].font.size = Pt(24)
    titel.text_frame.paragraphs[0].font.bold = True

    hoehe = Inches(0.38) * (len(block) + 1)
    tabelle = folie.shapes.add_table(
        len(block) + 1, len(SPALTEN), Inches(0.5), Inches(1.1), sum(BREITEN), hoehe
    ).table

    for i, b in enumerate(BREITEN):
        tabelle.columns[i].width = b

    for c, name in enumerate(SPALTEN):  # Kopfzeile auf jeder Folie
        zelle_setzen(tabelle.cell(0, c), name, fett=True,
                     farbe=RGBColor(255, 255, 255), fuellung=RGBColor(31, 78, 121),
                     mitte=(c != 2))

    for r, zeile in enumerate(block, start=1):
        grau = RGBColor(242, 242, 242) if r % 2 == 0 else RGBColor(255, 255, 255)
        for c, wert in enumerate(zeile):
            zelle_setzen(tabelle.cell(r, c), wert, farbe=RGBColor(0, 0, 0),
                         fuellung=grau, mitte=(c != 2))

prs.save("ersatzteilliste.pptx")
print(f"{len(zeilen)} Artikel auf {len(seiten)} Folien gespeichert.")
