import csv,io,urllib.request
from docx import Document
from docx.shared import Inches,Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
URL="https://raw.githubusercontent.com/saidiadegoke/nigeria-inec-geo/3cf617721aea519eef960681d3d86a4fad4325da/data/polling-units.csv"
rows=list(csv.DictReader(io.StringIO(urllib.request.urlopen(URL,timeout=120).read().decode("utf-8-sig"))))
rows=[r for r in rows if (r.get("state_code") or "").zfill(2)=="04" or (r.get("state_name") or "").strip().upper()=="ANAMBRA"]
assert len(rows)==5720
assert len({(r.get("lga_code",""),r.get("ward_code","")) for r in rows})==326
assert len({r.get("lga_code","") for r in rows})==21
rows.sort(key=lambda r:(r.get("lga_code",""),r.get("ward_code",""),r.get("code","")))
doc=Document(); sec=doc.sections[0]; sec.orientation=WD_ORIENT.LANDSCAPE; sec.page_width,sec.page_height=sec.page_height,sec.page_width
sec.top_margin=Inches(.45);sec.bottom_margin=Inches(.45);sec.left_margin=Inches(.35);sec.right_margin=Inches(.35)
p=doc.add_paragraph();p.alignment=1;x=p.add_run("ANAMBRA STATE — COMPLETE POLLING UNIT DIRECTORY");x.bold=True;x.font.size=Pt(15)
p=doc.add_paragraph();p.alignment=1;x=p.add_run("5,720 Polling Units • 326 Wards • 21 Local Government Areas");x.bold=True;x.font.size=Pt(10)
p=doc.add_paragraph("Source: INEC Polling Unit Locator-derived data. Specific Location / Address uses the INEC locator location field where supplied. Where INEC does not separately provide a location value, the polling-unit name is used as the location reference, without presenting it as a separate official street address.");p.runs[0].font.size=Pt(8)
headers=["S/N","LGA Code","LGA","Ward Code","Ward","Polling Unit Code","Polling Unit","Specific Location / Address","Status"];widths=[.38,.58,.9,.62,1,1.05,2.15,3.65,1]
t=doc.add_table(rows=1,cols=9);t.style="Table Grid";t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
for i,h in enumerate(headers):
 c=t.rows[0].cells[i];c.text=h;c.width=Inches(widths[i]);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
 for run in c.paragraphs[0].runs: run.bold=True;run.font.size=Pt(7)
for i,r in enumerate(rows,1):
 lga=(r.get("lga_name") or "").strip();ward=(r.get("ward_name") or "").strip();pu=(r.get("name") or "").strip();loc=(r.get("location") or "").strip()
 if not loc: loc=pu
 vals=[str(i),f"04-{(r.get('lga_code') or '').zfill(2)}",lga,(r.get("ward_code") or "").zfill(2),ward,f"04-{(r.get('lga_code') or '').zfill(2)}-{(r.get('ward_code') or '').zfill(2)}-{(r.get('code') or '').zfill(3)}",pu,loc,"Existing"]
 cs=t.add_row().cells
 for j,v in enumerate(vals):
  cs[j].text=v;cs[j].width=Inches(widths[j]);cs[j].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
  for p in cs[j].paragraphs:
   for run in p.runs: run.font.size=Pt(6.5)
trPr=t.rows[0]._tr.get_or_add_trPr();hdr=OxmlElement("w:tblHeader");hdr.set(qn("w:val"),"true");trPr.append(hdr)
footer=doc.sections[0].footer.paragraphs[0];footer.alignment=1;footer.add_run("Anambra State Polling Unit Directory • INEC Locator-derived • 25 September 2026").font.size=Pt(7)
doc.save("Anambra_5720_Polling_Units_INEC_Locator_Directory.docx")
print("Generated 5,720 PUs / 326 wards / 21 LGAs")
