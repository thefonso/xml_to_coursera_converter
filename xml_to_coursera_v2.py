"""Fill the Coursera template's 'FOR IMPORT' tab from a Xyleme WebCourse XML export.
Usage: python xml_to_coursera.py course.xml template.xlsx output.xlsx
Requires: pip install openpyxl
"""
import sys, copy, datetime as dt
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

OFFERING_TYPE = "Private"     # Private / Public / Your Org
ITEM_TYPE = "Reading"
# --- Page content options -------------------------------------------------
# Every page's body is always written to pages/NN_title.html and (plain text)
# into the Description column. To make the content show up INSIDE the lesson on
# import, host the HTML files at a public URL and set BASE_URL; each item then
# becomes an "Ungraded Plugin" (iframe) pointing at its page.
BASE_URL = ""   # e.g. "https://example.org/ai-cs/"  (must end with /)
OUT_DIR = "pages"
LESSON_NAME = "Orientation"
FIRST_ITEM_ROW, TEMPLATE_ITEM_ROWS = 40, 5   # rows 40-44 in the template
LESSON_END_ROW = 45                           # the "Add as many rows" line

import os, re, html as _h
SKIP = {"Title","DesignData","Figure","Properties","Instructions","LOM"}
def to_html(el):
    out=[]
    for ch in el:
        t=ch.tag
        if t in SKIP: continue
        if t=="RichText":
            x=text(ch)
            if x: out.append(f"<p>{_h.escape(x)}</p>")
        elif t=="List":
            pre=ch.find("ListPreamble")
            if pre is not None and text(pre): out.append(f"<p>{_h.escape(text(pre))}</p>")
            items=[text(i) for i in ch.iter("ItemPara") if text(i)]
            out.append("<ul>"+"".join(f"<li>{_h.escape(i)}</li>" for i in items)+"</ul>")
        elif t=="Tab":
            out.append(f"<h3>{_h.escape(text(ch.find('Label')))}</h3>")
            c=ch.find("Content")
            if c is not None: out.append(to_html(c))
        else:
            out.append(to_html(ch))
    return "\n".join(o for o in out if o)

def text(el):
    return " ".join("".join(el.itertext()).split()) if el is not None else ""

xml_path, tpl_path, out_path = sys.argv[1:4]
root = ET.parse(xml_path).getroot()
course_title = text(root.find("Title"))
container = root.find("Container")
module_title = text(container.find("Title"))
minutes = int(text(container.find(".//TypicalLearningTime")) or 0)
pages = container.findall("Page")

# Course description = first RichText on page 1; module description = first
# paragraph on the page titled "Summary..."; objectives = bullets on that page.
course_desc = text(pages[0].find(".//RichText"))
summary = next((p for p in pages if "summary" in text(p.find(".//Title")).lower()), None)
module_desc = text(summary.find(".//ParaBlock/RichText")) if summary is not None else ""
objectives = [text(i) for i in summary.iter("ItemPara")] if summary is not None else []

wb = openpyxl.load_workbook(tpl_path)
ws = wb["FOR IMPORT"]

# Remove everything after the first lesson block (extra placeholder lesson/modules)
ws.delete_rows(LESSON_END_ROW + 1, ws.max_row)

# Resize item rows to match number of pages
n = len(pages)
if n < TEMPLATE_ITEM_ROWS:
    ws.delete_rows(FIRST_ITEM_ROW + n, TEMPLATE_ITEM_ROWS - n)
elif n > TEMPLATE_ITEM_ROWS:
    extra = n - TEMPLATE_ITEM_ROWS
    ws.insert_rows(FIRST_ITEM_ROW + TEMPLATE_ITEM_ROWS, extra)
    for r in range(FIRST_ITEM_ROW + TEMPLATE_ITEM_ROWS, FIRST_ITEM_ROW + n):
        for c in range(1, ws.max_column + 1):
            src, dst = ws.cell(FIRST_ITEM_ROW, c), ws.cell(r, c)
            dst._style = copy.copy(src._style)
    # carry dropdowns/checkboxes over to the new rows
    for dv in list(ws.data_validations.dataValidation):
        for rng in list(dv.sqref.ranges):
            if rng.min_row <= FIRST_ITEM_ROW <= rng.max_row:
                new = copy.copy(dv)
                new.sqref = f"{rng.coord.split(':')[0][0]}{FIRST_ITEM_ROW + TEMPLATE_ITEM_ROWS}:" \
                            f"{rng.coord.split(':')[0][0]}{FIRST_ITEM_ROW + n - 1}"
                ws.add_data_validation(new)

dur = dt.timedelta(minutes=minutes)
ws["B16"], ws["B17"], ws["B18"] = course_title, course_desc, dur
ws["B23"] = OFFERING_TYPE
ws["B26"], ws["B27"], ws["B28"], ws["B29"] = 1, module_title, module_desc, dur
for i in range(32, 35):                      # learning objective rows A32:A34
    ws.cell(i, 1).value = objectives[i - 32] if i - 32 < len(objectives) else None
ws["B36"], ws["B37"] = 1, LESSON_NAME

per_item = dt.timedelta(minutes=minutes / n) if n else dt.timedelta(0)
for i, page in enumerate(pages):
    r = FIRST_ITEM_ROW + i
    ws.cell(r, 1).value = ITEM_TYPE
    ws.cell(r, 2).value = text(page.find(".//Title"))
    for c in range(5, 14):                   # clear video-only placeholders E:M
        ws.cell(r, c).value = None
    title = text(page.find(".//Title"))
    body = to_html(page.find("Topic"))
    fname = f"{i+1:02d}_" + re.sub(r"[^a-z0-9]+","-",title.lower()).strip("-") + ".html"
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(os.path.join(OUT_DIR, fname), "w", encoding="utf-8") as f:
        f.write(f"<!doctype html><meta charset='utf-8'><title>{_h.escape(title)}</title>"
                f"<body style='font-family:sans-serif;max-width:46em;margin:2em auto'>"
                f"<h2>{_h.escape(title)}</h2>\n{body}</body>")
    ws.cell(r, 3).value = re.sub(r"<[^>]+>", " ", body.replace("</p>","\n").replace("</li>","\n")).strip()
    ws.cell(r, 3).value = _h.unescape("\n".join(l.strip() for l in ws.cell(r,3).value.splitlines() if l.strip()))
    if BASE_URL:
        ws.cell(r, 1).value = "Ungraded Plugin"
        ws.cell(r, 5).value = BASE_URL + fname
    ws.cell(r, 4).value = per_item

wb.save(out_path)
print(f"Wrote {out_path}: {n} items, {len(objectives)} objectives found")
