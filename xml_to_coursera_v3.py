#!/usr/bin/env python3
"""
Fill the Coursera Course Template ('FOR IMPORT' tab) from a Xyleme WebCourse XML export.
Handles light courses (pages directly under one container) and robust ones (lessons
nested in containers, multiple modules, videos, knowledge checks).

Usage:
    python xml_to_coursera_v3.py course.xml template.xlsx output.xlsx
    python xml_to_coursera_v3.py course.xml template.xlsx output.xlsx --base-url https://host/path/

Requires: pip install openpyxl

How the XML is mapped
---------------------
  WebCourse > Container (course container)
    - has Pages only                      -> 1 module, 1 lesson
    - has child Containers (no deeper)    -> 1 module, one lesson per child Container
    - has child Containers with children  -> one module per child, one lesson per grandchild
  Each <Page> becomes one item:
    page with a <Movie>        -> Video
    page with <QuestionBlock>  -> KNOWLEDGE_CHECK_TYPE (default "Assignment")
    anything else              -> Reading

Every page's body is also saved as an HTML file in <output>_pages/ (see --base-url).
"""
import argparse, copy, datetime as dt, html as _h, os, re, sys
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.cell_range import MultiCellRange

# ----------------------------- settings you may edit -----------------------------
OFFERING_TYPE = "Private"          # Private / Public / Your Org
READING_TYPE = "Reading"
VIDEO_TYPE = "Video"
KNOWLEDGE_CHECK_TYPE = "Assignment"  # try "Reading" if the import rejects Assignment rows
MERGE_KNOWLEDGE_CHECKS = False     # True: all question pages in a lesson -> ONE item
FILL_DESCRIPTION = True            # write the page's plain text into the Description column
DEFAULT_ITEM_MINUTES = 2           # used when the XML has no learning time for a lesson
# ----------------------------------------------------------------------------------

SKIP = {"Title", "DesignData", "Properties", "Instructions", "LOM", "Sequencing", "Icon",
        "InlineDesignInstruction", "WebActivityBehavior", "Renditions", "MediaObject"}
warnings = []


# ------------------------------- XML text helpers ---------------------------------
def clean_text(el):
    """Text of el, dropping authoring notes (<InlineDesignInstruction>)."""
    parts = []
    def rec(e):
        if e.text:
            parts.append(e.text)
        for c in e:
            if c.tag != "InlineDesignInstruction":
                rec(c)
            if c.tail:
                parts.append(c.tail)
    rec(el)
    return "".join(parts)

def one_line(el):
    return " ".join(clean_text(el).split()) if el is not None else ""

def paras(el):
    return [" ".join(p.split()) for p in re.split(r"\n\s*\n", clean_text(el)) if p.strip()]

def esc(s):
    return _h.escape(s, quote=False)


# ------------------------------ XML -> HTML rendering -----------------------------
def render_list(lst):
    marker = lst.get("ListMarker", "Bullet")
    tag = "ul" if marker in ("Bullet", "None", "") else "ol"
    attr = ' type="a"' if marker == "LowercaseAlpha" else ""
    out = []
    pre = lst.find("ListPreamble")
    if pre is not None:
        out += [f"<p>{esc(p)}</p>" for p in paras(pre)]
    lis = []
    for ib in lst.findall("ItemBlock"):
        for item in ib.findall("Item"):
            text = " ".join(one_line(ip) for ip in item.findall("ItemPara"))
            subs = "".join(render_list(s) for s in item.findall("SubList"))
            lis.append(f"<li>{esc(text)}{subs}</li>")
    out.append(f"<{tag}{attr}>{''.join(lis)}</{tag}>")
    return "".join(out)

def render_question(qb):
    out = []
    title = one_line(qb.find("Title"))
    out.append(f"<h3>{esc(title) if title else 'Question'}</h3>")
    handled = False
    for q in qb:
        if q.tag != "MultipleChoice":
            continue
        handled = True
        stem = q.find("QuestionStem")
        out += [f"<p>{esc(p)}</p>" for p in (paras(stem) if stem is not None else [])]
        lis = []
        for o in q.findall("Option"):
            mark = " <strong>(correct)</strong>" if o.get("IsCorrect") == "true" else ""
            lis.append(f"<li>{esc(one_line(o.find('Distractor')))}{mark}</li>")
        out.append(f"<ul>{''.join(lis)}</ul>")
        fb = q.find("Feedback")
        if fb is not None:
            pos, neg = fb.find("RichPositiveFeedback"), fb.find("RichNegativeFeedback")
            if pos is not None and one_line(pos):
                out.append(f"<p><em>Feedback (correct):</em> {esc(one_line(pos))}</p>")
            if neg is not None and one_line(neg):
                out.append(f"<p><em>Feedback (incorrect):</em> {esc(one_line(neg))}</p>")
    if not handled:   # unknown question type: keep the text rather than lose it
        out += [f"<p>{esc(p)}</p>" for p in paras(qb)]
        warnings.append(f"Question type other than MultipleChoice in '{title or 'question'}' - check output.")
    return "\n".join(out)

def render(el, is_transition=False):
    out = []
    for ch in el:
        t = ch.tag
        if t in SKIP:
            continue
        if t == "RichText":
            out += [f"<p>{esc(p)}</p>" for p in paras(ch)]
        elif t == "List":
            out.append(render_list(ch))
        elif t == "Tab":
            out.append(f"<h3>{esc(one_line(ch.find('Label')))}</h3>")
            c = ch.find("Content")
            if c is not None:
                out.append(render(c))
        elif t == "Movie":
            out.append(f"<p><strong>Video:</strong> {esc(one_line(ch.find('Title')))}</p>")
            src = one_line(ch.find("LaunchMovie"))
            cc = one_line(ch.find("ClosedCaptionFile"))
            out.append(f"<p><small>Source file: {esc(src)}<br>Captions: {esc(cc)}</small></p>")
        elif t == "Figure":
            if ch.get("borderType") == "BackgroundImage" or is_transition:
                continue
            for web in ch.iter("Web"):
                alt = (web.get("altText") or "").strip()
                if alt:
                    out.append(f"<p><em>[Image description: {esc(' '.join(alt.split()))}]</em></p>")
        elif t == "QuestionBlock":
            out.append(render_question(ch))
        else:
            out.append(render(ch, is_transition))
    return "\n".join(o for o in out if o)

def html_to_plain(body):
    s = re.sub(r"</(p|li|h3|ul|ol)>|<br>", "\n", body)
    s = _h.unescape(re.sub(r"<[^>]+>", "", s))
    return "\n".join(l.strip() for l in s.splitlines() if l.strip())


# --------------------------------- XML structure ----------------------------------
def minutes_of(container):
    el = container.find("DesignData/LOM/Educational/TypicalLearningTime")
    try:
        return float(el.text) if el is not None and el.text and el.text.strip() else None
    except ValueError:
        return None

def build_structure(course):
    """Return list of modules: dict(title, minutes, lessons=[dict(title, pages, minutes)])."""
    subs = course.findall("Container")
    title = one_line(course.find("Title"))
    def lesson(c, name=None):
        return dict(title=name or one_line(c.find("Title")), pages=list(c.iter("Page")), minutes=minutes_of(c))
    if not subs:
        return [dict(title=title, minutes=minutes_of(course),
                     lessons=[dict(title=title, pages=course.findall("Page"), minutes=minutes_of(course))])]
    if not any(s.findall("Container") for s in subs):
        lessons = [lesson(s) for s in subs]
        if course.findall("Page"):
            lessons.insert(0, dict(title="Introduction", pages=course.findall("Page"), minutes=None))
        return [dict(title=title, minutes=minutes_of(course), lessons=lessons)]
    if course.findall("Page"):
        warnings.append("Pages directly under the course container were not placed (course has modules).")
    modules = []
    for m in subs:
        gs = m.findall("Container")
        lessons = [lesson(g) for g in gs]
        if m.findall("Page"):
            lessons.insert(0, dict(title="Introduction", pages=m.findall("Page"), minutes=None))
        if not gs:
            lessons = [lesson(m)]
        modules.append(dict(title=one_line(m.find("Title")), minutes=minutes_of(m), lessons=lessons))
    return modules

def classify(page):
    """-> (kind, base name)  kind in video / quiz / reading"""
    topic = page.find("Topic")
    title = one_line(topic.find("Title")) if topic is not None else ""
    movie = next(page.iter("Movie"), None)
    if movie is not None:
        return "video", one_line(movie.find("Title")) or title
    if next(page.iter("QuestionBlock"), None) is not None:
        return "quiz", title
    return "reading", title

def goal_sentences(pages):
    out = []
    for p in pages[:1]:
        for rt in p.iter("RichText"):
            for para in paras(rt):
                if para.lower().startswith("by the end of"):
                    out.append(para)
    return out

def to_objective(sentence):
    m = re.match(r"(?i)by the end of (this|the) \w+, you will (be able to )?(.*)", sentence)
    txt = m.group(3) if m else sentence
    return txt[:1].upper() + txt[1:]


# ------------------------------ build item list -----------------------------------
def build_items(lesson, outdir, lesson_no):
    """Create item dicts + write HTML files for one lesson."""
    pages = lesson["pages"]
    first_title = one_line(pages[0].find("Topic/Title")) if pages else ""
    items, qn = [], 0
    for i, page in enumerate(pages):
        kind, base = classify(page)
        topic = page.find("Topic")
        title = one_line(topic.find("Title")) if topic is not None else base
        is_trans = (i == 0 and title == lesson["title"]) or (i == 0 and "Transition" in
                    [one_line(a.find("Name")) for a in page.iter("Attribute")])
        body = render(topic, is_trans) if topic is not None else ""
        name, note = base or f"Page {i+1}", ""
        if is_trans:
            name = f"Overview: {name}"
        if kind == "quiz":
            qn += 1
            qb = next(page.iter("QuestionBlock"))
            stem = next((one_line(q.find("QuestionStem")) for q in qb.iter("MultipleChoice")), "")
            qtitle = one_line(qb.find("Title"))
            name = f"{base} {qn}: {qtitle or stem}"[:80].rstrip()
            note = "Quiz question - build in Coursera quiz editor (see description / HTML file)"
        if kind == "video":
            mv = next(page.iter("Movie"))
            note = f"Source: {one_line(mv.find('LaunchMovie'))} | Captions: {one_line(mv.find('ClosedCaptionFile'))}"
        if len(name) < 5:   # Coursera import silently drops very short item names
            name = f"{name} - {lesson['title']}"[:80]
            warnings.append(f"Short page title '{base}' renamed to '{name}'.")
        items.append(dict(kind=kind, name=name, body=body, note=note, title=title))

    if MERGE_KNOWLEDGE_CHECKS:
        qs = [it for it in items if it["kind"] == "quiz"]
        if len(qs) > 1:
            merged = dict(kind="quiz", name=f"Knowledge check ({len(qs)} questions)",
                          body="\n<hr>\n".join(q["body"] for q in qs), note=qs[0]["note"], title="Knowledge check")
            pos = items.index(qs[0])
            items = [it for it in items if it["kind"] != "quiz"]
            items.insert(pos, merged)

    for n, it in enumerate(items, 1):
        slug = re.sub(r"[^a-z0-9]+", "-", it["name"].lower()).strip("-")[:50]
        it["file"] = f"{lesson_no:02d}-{n:02d}_{slug}.html"
        os.makedirs(outdir, exist_ok=True)
        with open(os.path.join(outdir, it["file"]), "w", encoding="utf-8") as f:
            f.write("<!doctype html><meta charset='utf-8'><title>%s</title>"
                    "<body style='font-family:sans-serif;max-width:46em;margin:2em auto;line-height:1.5'>"
                    "<h2>%s</h2>\n%s</body>" % (esc(it["name"]), esc(it["name"]), it["body"]))
    return items


# ------------------------------- write workbook -----------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xml"); ap.add_argument("template"); ap.add_argument("output")
    ap.add_argument("--base-url", default="", help="Public URL folder (ending /) where the HTML pages will be hosted; "
                    "items then import as Ungraded Plugins pointing at them")
    ap.add_argument("--out-dir", default="", help="Folder for page HTML files (default: <output>_pages)")
    ap.add_argument("--offering-type", default=OFFERING_TYPE)
    a = ap.parse_args()
    base_url = a.base_url
    if base_url and not base_url.endswith("/"):
        base_url += "/"
    outdir = a.out_dir or os.path.splitext(a.output)[0] + "_pages"

    root = ET.parse(a.xml).getroot()
    course = root.find("Container")
    if course is None:
        sys.exit("No <Container> found in the XML - is this a Xyleme WebCourse export?")
    course_title = one_line(root.find("Title")) or one_line(course.find("Title"))
    modules = build_structure(course)

    wb = openpyxl.load_workbook(a.template)
    ws = wb["FOR IMPORT"]
    MAXC = ws.max_column

    # --- snapshot prototype rows (value, style, note) before clearing -----------------
    def snap(r):
        cells = []
        for c in range(1, MAXC + 1):
            cell = ws.cell(r, c)
            cm = cell.comment
            cells.append((c, cell.value, copy.copy(cell._style),
                          Comment(cm.text, cm.author, height=cm.height, width=cm.width) if cm else None))
        return cells, ws.row_dimensions[r].height
    P = {r: snap(r) for r in (26, 27, 28, 29, 30, 31, 32, 35, 36, 37, 38, 39, 40, 45, 46)}
    dv_type = next(d for d in ws.data_validations.dataValidation if "A40" in str(d.sqref))
    dv_video = next(d for d in ws.data_validations.dataValidation if "H40" in str(d.sqref))
    dv_other = [d for d in ws.data_validations.dataValidation if d is not dv_type and d is not dv_video]

    # --- clear everything from row 26 down -----------------------------------------
    ws.delete_rows(26, ws.max_row)
    for k in [k for k in ws.row_dimensions.keys() if k >= 26]:
        del ws.row_dimensions[k]
    old_cf, ws.conditional_formatting = ws.conditional_formatting, ConditionalFormattingList()
    for cf in old_cf:                         # keep highlighting rules for rows above 26
        if all(r.max_row < 26 for r in cf.sqref.ranges):
            for rule in cf.rules:
                ws.conditional_formatting.add(str(cf.sqref), rule)
    ws.data_validations.dataValidation = dv_other

    cur = [26]
    def put(proto_row, values=None):
        cells, h = P[proto_row]
        r = cur[0]
        for c, val, style, cm in cells:
            cell = ws.cell(r, c)
            cell._style = copy.copy(style)
            v = val
            if values is not None and c in values:
                v = values[c]
            elif values is not None and proto_row in (40,):
                v = None                        # blank out placeholder text in item rows
            cell.value = v
            if cm:
                cell.comment = Comment(cm.text, cm.author, height=cm.height, width=cm.width)
        if h:
            ws.row_dimensions[r].height = h
        cur[0] += 1
        return r

    # --- course-level fields -----------------------------------------------------
    all_goals, all_objs, total_secs = [], [], 0
    item_ranges, summary = [], []
    lesson_no = 0
    for mi, mod in enumerate(modules, 1):
        # build items first so we know durations
        for ls in mod["lessons"]:
            lesson_no += 1
            ls["items"] = build_items(ls, outdir, lesson_no)
            n = max(len(ls["items"]), 1)
            mins = ls["minutes"] if ls["minutes"] is not None else DEFAULT_ITEM_MINUTES * n
            if ls["minutes"] is None:
                warnings.append(f"No learning time in XML for lesson '{ls['title']}' - used {DEFAULT_ITEM_MINUTES} min/item.")
            ls["secs"] = int(round(mins * 60))
        mod_secs = sum(l["secs"] for l in mod["lessons"])
        total_secs += mod_secs
        goals = [g for l in mod["lessons"] for g in goal_sentences(l["pages"])]
        # a "Summary" page, if the course has one, supplies the module description + bullets
        summ = next((p for l in mod["lessons"] for p in l["pages"]
                     if "summary" in one_line(p.find("Topic/Title")).lower()), None)
        desc = ""
        objs = []
        if summ is not None:
            first = next(summ.iter("RichText"), None)
            desc = paras(first)[0] if first is not None and paras(first) else ""
            objs = [one_line(i) for i in summ.iter("ItemPara")]
        if not desc:
            desc = " ".join(goals)
        if not objs:
            objs = [to_objective(g) for g in goals]
        mod["desc"], mod["objs"] = desc, objs
        all_goals.append(desc)

        # --- module block ---
        r0 = put(26, {2: mi})
        put(27, {2: mod["title"]})
        put(28, {2: desc})
        put(29, {2: dt.timedelta(seconds=mod_secs)})
        put(30)
        put(31)
        for ob in (objs or [None]):
            put(32, {1: ob})
        put(35)

        for li, ls in enumerate(mod["lessons"], 1):
            put(36, {2: li})
            put(37, {2: ls["title"]})
            put(38)
            put(39)
            first_item = cur[0]
            per = int(round(ls["secs"] / max(len(ls["items"]), 1)))
            for it in ls["items"]:
                typ = {"video": VIDEO_TYPE, "quiz": KNOWLEDGE_CHECK_TYPE, "reading": READING_TYPE}[it["kind"]]
                link = None
                if base_url and it["kind"] == "reading":   # videos/quizzes keep their own type
                    typ, link = "Ungraded Plugin", base_url + it["file"]
                desc_txt = html_to_plain(it["body"]) if FILL_DESCRIPTION else None
                if desc_txt and len(desc_txt) > 32000:
                    desc_txt = desc_txt[:32000]
                put(40, {1: typ, 2: it["name"], 3: desc_txt, 4: dt.timedelta(seconds=per),
                         5: link, 6: it["note"] or None})
            item_ranges.append((first_item, cur[0] - 1))
            put(45)
            put(46)
            summary.append((mod["title"], ls["title"], len(ls["items"]),
                            sum(1 for i in ls["items"] if i["kind"] == "video"),
                            sum(1 for i in ls["items"] if i["kind"] == "quiz")))

    ws["B16"] = course_title
    ws["B17"] = " ".join(dict.fromkeys(all_goals)) or course_title
    ws["B18"] = dt.timedelta(seconds=total_secs)
    ws["B23"] = a.offering_type

    # --- re-create the dropdowns on the new item rows -----------------------------------
    for dv, col in ((dv_type, "A"), (dv_video, "H")):
        dv.sqref = MultiCellRange(" ".join(f"{col}{s}:{col}{e}" for s, e in item_ranges))
        ws.add_data_validation(dv)

    wb.save(a.output)

    print(f"\nWrote {a.output}")
    print(f"Page files: {outdir}  ({'linked via --base-url' if base_url else 'not linked; use --base-url to link'})")
    print(f"\n{'Module':40} {'Lesson':40} items video quiz")
    for m, l, n, v, q in summary:
        print(f"{m[:39]:40} {l[:39]:40} {n:5} {v:5} {q:4}")
    if warnings:
        print("\nWarnings:")
        for w in dict.fromkeys(warnings):
            print(" -", w)

if __name__ == "__main__":
    main()