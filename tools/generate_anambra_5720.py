import csv, io, zipfile, urllib.request
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

URL = "https://raw.githubusercontent.com/saidiadegoke/nigeria-inec-geo/3cf617721aea519eef960681d3d86a4fad4325da/data/polling-units.csv"
raw = urllib.request.urlopen(URL, timeout=120).read()
rows = list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig"))))
rows = [r for r in rows if (r.get("state_code") or "").zfill(2) == "04" or (r.get("state_name") or "").strip().upper() == "ANAMBRA"]

assert len(rows) == 5720, f"Expected 5720 Anambra PUs, got {len(rows)}"
wards = {(r.get("lga_code",""), r.get("ward_code","")) for r in rows}
lgas = {r.get("lga_code","") for r in rows}
assert len(wards) == 326, f"Expected 326 wards, got {len(wards)}"
assert len(lgas) == 21, f"Expected 21 LGAs, got {len(lgas)}"

rows.sort(key=lambda r: (r.get("lga_code",""), r.get("ward_code",""), r.get("code","")))
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = sec.page_height, sec.page_width
sec.top_margin = Inches(.45); sec.bottom_margin = Inches(.45)
sec.left_margin = Inches(.35); sec.right_margin = Inches(.35)

p = doc.add_paragraph()
p.alignment = 1
run = p.add_run("ANAMBRA STATE — COMPLETE POLLING UNIT DIRECTORY")
run.bold = True; run.font.size = Pt(15)
p = doc.add_paragraph()
p.alignment = 1
r = p.add_run("5,720 Polling Units • 326 Wards • 21 Local Government Areas")
r.bold = True; r.font.size = Pt(10)
p = doc.add_paragraph("Source basis: INEC Polling Unit Locator-derived dataset (nigeria-inec-geo), which states that its data were scraped directly from the INEC Continuous Voter Registration Polling Unit portal. Location values are used where INEC records them separately; where blank, an estimated location is constructed from the official PU name + Ward + LGA. Current Status identifies locator listing/status, not election-day deployment.")
p.runs[0].font.size = Pt(8)

headers = ["S/N","LGA Code","LGA","Ward Code","Ward","Polling Unit Code","Polling Unit","Estimated or Specific Location / Address","Current Status"]
table = doc.add_table(rows=1, cols=len(headers))
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = "Table Grid"
table.autofit = False
widths = [0.38,0.48,0.9,0.62,1.0,0.85,2.15,3.65,1.25]
for i,h in enumerate(headers):
    c=table.rows[0].cells[i]; c.text=h; c.width=Inches(widths[i])
    for rr in c.paragraphs[0].runs: rr.bold=True; rr.font.size=Pt(7)
    c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

for i,r in enumerate(rows,1):
    lga = (r.get("lga_name") or "").strip()
    ward = (r.get("ward_name") or "").strip()
    pu = (r.get("name") or "").strip()
    loc = (r.get("location") or "").strip()
    status = "Listed in INEC locator-derived dataset"
    if not loc:
        loc = f"Estimated: {pu}, {ward}, {lga}, Anambra State"
        status = "Listed; location estimated from INEC PU name"
    vals = [
        str(i), (r.get("lga_code") or "").zfill(2), lga,
        (r.get("ward_code") or "").zfill(2), ward,
        (r.get("code") or "").zfill(3), pu, loc, status
    ]
    cells=table.add_row().cells
    for j,v in enumerate(vals):
        cells[j].text=v; cells[j].width=Inches(widths[j]); cells[j].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for p in cells[j].paragraphs:
            for rr in p.runs: rr.font.size=Pt(6.5)

# Repeat header row
trPr = table.rows[0]._tr.get_or_add_trPr()
tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), "true"); trPr.append(tblHeader)

# Footer
for section in doc.sections:
    footer = section.footer.paragraphs[0]
    footer.alignment = 1
    footer.add_run("Anambra State Polling Unit Directory • INEC Locator-derived • Prepared 25 September 2026").font.size = Pt(7)

doc.save("Anambra_5720_Polling_Units_INEC_Locator_Directory.docx")
print("Generated 5,720 PUs / 326 wards / 21 LGAs")
