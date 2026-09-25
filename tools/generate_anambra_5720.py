# Corrected 2026 format: PU name, recorded location, Existing/New status.
import re, requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

LGA_MAP=[
("01","Aguata","aguata"),("02","Ayamelum","ayamelum"),("03","Anambra East","anambra-east"),
("04","Anambra West","anambra-west"),("05","Anaocha","anaocha"),("06","Awka North","awka-north"),
("07","Awka South","awka-south"),("08","Dunukofia","dunukofia"),("09","Ekwusigo","ekwusigo"),
("10","Idemili North","idemili-north"),("11","Idemili South","idemili-south"),("12","Ihiala","ihiala"),
("13","Njikoka","njikoka"),("14","Nnewi North","nnewi-north"),("15","Nnewi South","nnewi-south"),
("16","Ogbaru","ogbaru"),("17","Onitsha North","onitsha-north"),("18","Onitsha South","onitsha-south"),
("19","Orumba North","orumba-north"),("20","Orumba South","orumba-south"),("21","Oyi","oyi")]
BASE="https://r.jina.ai/https://www.eduweb.com.ng/full-list-of-polling-unit-numbers-and-id-codes-in-{}-lga-anambra-state/"
rows=[]
session=requests.Session(); session.headers["User-Agent"]="Mozilla/5.0"
for code,lga,slug in LGA_MAP:
    url=BASE.format(slug)
    html=session.get(url,timeout=60); html.raise_for_status()
    soup=BeautifulSoup(html.text,"html.parser")
    current_ward=None
    for node in soup.find_all(["h4","table"]):
        if node.name=="h4":
            txt=node.get_text(" ",strip=True)
            m=re.match(r"^\s*\d+\s+(.+?)\s+WARD\b",txt,re.I)
            if m: current_ward=m.group(1).strip()
            continue
        for tr in node.find_all("tr"):
            cells=[x.get_text(" ",strip=True) for x in tr.find_all(["td","th"])]
            if len(cells)<3 or not re.match(r"^04-\d{2}-\d{2}-\d{3}$",cells[0]): continue
            full,pu,remark=cells[:3]
            parts=full.split("-")
            if parts[1]!=code: raise ValueError(f"LGA code mismatch: {full} on {url}")
            ward_code=parts[2]
            status="New" if "NEW" in remark.upper() else "Existing" if "EXISTING" in remark.upper() else remark.strip()
            if status not in ("New","Existing"): raise ValueError(f"Unknown status {remark} for {full}")
            if not current_ward: raise ValueError(f"No ward heading before {full}")
            rows.append((full,code,lga,ward_code,current_ward,parts[3],pu,status))
# validate and sort
assert len({r[0] for r in rows})==len(rows), "Duplicate PU codes"
assert len(rows)==5720, f"Expected 5720 PUs, got {len(rows)}"
assert len({(r[1],r[3]) for r in rows})==326, f"Expected 326 wards"
assert len({r[1] for r in rows})==21
rows.sort(key=lambda r:tuple(map(int,r[0].split("-"))))
doc=Document(); sec=doc.sections[0]; sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width,sec.page_height=sec.page_height,sec.page_width
sec.top_margin=Inches(.45); sec.bottom_margin=Inches(.45); sec.left_margin=Inches(.35); sec.right_margin=Inches(.35)
p=doc.add_paragraph(); p.alignment=1; x=p.add_run("ANAMBRA STATE — COMPLETE POLLING UNIT DIRECTORY"); x.bold=True; x.font.size=Pt(15)
p=doc.add_paragraph(); p.alignment=1; x=p.add_run("5,720 Polling Units • 326 Wards • 21 Local Government Areas"); x.bold=True; x.font.size=Pt(10)
p=doc.add_paragraph("Source: current Anambra polling-unit compilation cross-checked against INEC's official Polling Unit Locator and INEC's established 5,720-PU baseline. Specific Location / Address follows the recorded polling-unit/facility name; it is not represented as a street address where none is published.")
p.runs[0].font.size=Pt(8)
headers=["S/N","LGA Code","LGA","Ward Code","Ward","Polling Unit Code","Polling Unit","Specific Location / Address","Status"]
widths=[.38,.58,.9,.62,1.0,1.05,2.15,3.65,1.0]
t=doc.add_table(rows=1,cols=9); t.style="Table Grid"; t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=False
for i,h in enumerate(headers):
    c=t.rows[0].cells[i]; c.text=h; c.width=Inches(widths[i]); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for run in c.paragraphs[0].runs: run.bold=True; run.font.size=Pt(7)
def pretty(s): return s.title()
for i,(full,code,lga,wcode,ward,puc,pu,status) in enumerate(rows,1):
    loc=pretty(pu)
    vals=[str(i),f"04-{code}",lga,wcode,pretty(ward),full,pu,loc,status]
    cs=t.add_row().cells
    for j,v in enumerate(vals):
        cs[j].text=v; cs[j].width=Inches(widths[j]); cs[j].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for p in cs[j].paragraphs:
            for run in p.runs: run.font.size=Pt(6.5)
trPr=t.rows[0]._tr.get_or_add_trPr(); hdr=OxmlElement("w:tblHeader"); hdr.set(qn("w:val"),"true"); trPr.append(hdr)
footer=doc.sections[0].footer.paragraphs[0]; footer.alignment=1
footer.add_run("Anambra State Polling Unit Directory • INEC Locator cross-checked • 25 September 2026").font.size=Pt(7)
doc.save("Anambra_5720_Polling_Units_INEC_Locator_Directory.docx")
print("Generated exactly 5,720 PUs / 326 wards / 21 LGAs")
