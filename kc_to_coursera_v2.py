#!/usr/bin/env python3
"""
Port the "Knowledge check" questions from a Xyleme WebCourse XML file into the
Coursera Assignment Import template (.docx).

Usage:
    python kc_to_coursera.py course.xml my_template.docx output.docx
    python kc_to_coursera.py course.xml my_template.docx output.docx --partial-credit --no-shuffle

Arguments (all three are required, in this order):
    course.xml         the Xyleme WebCourse XML file to read
    my_template.docx   the Coursera Assignment Import template to fill in
    output.docx        the file to create (overwritten if it exists)

Options:
    --partial-credit   add "partial credit" to multiple-correct (checkbox) questions
    --no-shuffle       never add "shuffle" to a question

Requires:  pip install python-docx lxml
"""
import argparse
import copy
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from docx import Document
from lxml import etree

START_MARK = "importable content starts here"
END_MARK = "end of importable content"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/commentsExtensible"


# --------------------------------------------------------------------------
# XML -> question objects
# --------------------------------------------------------------------------
def local(tag):
    return tag.split("}", 1)[-1]


def flat_text(el):
    """All text under el, ignoring <InlineDesignInstruction> authoring notes."""
    parts = []

    def walk(e):
        if local(e.tag) == "InlineDesignInstruction":
            return
        if e.text:
            parts.append(e.text)
        for c in e:
            walk(c)
            if c.tail:
                parts.append(c.tail)

    walk(el)
    return re.sub(r"\s+", " ", "".join(parts)).strip()


FORMAT_TAGS = {"Bold": "b", "Strong": "b", "Italic": "i", "Underline": "u"}


def segments(el):
    """Like flat_text, but returns [(text, bold, italic, underline), ...] so
    Coursera's rich-text support can keep <Bold>/<Italic>/<Underline>."""
    segs = []

    def walk(e, fmt):
        if local(e.tag) == "InlineDesignInstruction":
            return
        fmt = fmt | ({FORMAT_TAGS[local(e.tag)]} if local(e.tag) in FORMAT_TAGS else set())
        if e.text:
            segs.append([e.text, fmt])
        for c in e:
            walk(c, fmt)
            if c.tail:
                segs.append([c.tail, fmt])

    walk(el, set())
    # collapse whitespace across segment boundaries, trim ends, drop empties
    out, prev_space = [], True
    for text, fmt in segs:
        text = re.sub(r"\s+", " ", text)
        if prev_space and text.startswith(" "):
            text = text[1:]
        if not text:
            continue
        prev_space = text.endswith(" ")
        out.append([text, fmt])
    if out:
        out[-1][0] = out[-1][0].rstrip()
        if not out[-1][0]:
            out.pop()
    return [(t, "b" in f, "i" in f, "u" in f) for t, f in out]


def rich_texts(el):
    """One segment list per non-empty <RichText> under el."""
    return [sg for sg in (segments(rt) for rt in el.iter("RichText")) if sg]


def join_segs(lists):
    out = []
    for i, sg in enumerate(lists):
        if i:
            out.append((" ", False, False, False))
        out.extend(sg)
    return out


def parent_map(root):
    return {c: p for p in root.iter() for c in p}


def parse_xml(path, warn):
    """Return a list of question dicts from every 'Knowledge check' topic."""
    root = ET.parse(path).getroot()
    pmap = parent_map(root)
    questions = []

    for qb in root.iter("QuestionBlock"):
        # Find the enclosing Topic and its title.
        node, topic = qb, None
        while node in pmap:
            node = pmap[node]
            if local(node.tag) == "Topic":
                topic = node
                break
        title = ""
        if topic is not None:
            t = topic.find("Title")
            title = flat_text(t) if t is not None else ""
        if not title.lower().startswith("knowledge check"):
            warn(f"skipped a QuestionBlock inside topic '{title or '?'}' (not a 'Knowledge check')")
            continue

        handled = False
        for q in qb:
            kind = local(q.tag)
            if kind == "MultipleChoice":
                parsed = parse_multiple_choice(q, warn)
                if parsed:
                    questions.append(parsed)
                handled = True
            elif kind in ("Title", "WebActivityBehavior"):
                continue
            else:
                warn(f"UNSUPPORTED question type <{kind}> in '{title}' - NOT ported, add it by hand")
                handled = True
        if not handled:
            warn(f"empty QuestionBlock in '{title}'")
    return questions


def plain(sg):
    return "".join(t for t, *_ in sg)


def parse_multiple_choice(mc, warn):
    stem = mc.find("QuestionStem")
    prompt = rich_texts(stem) if stem is not None else []
    if not prompt:
        warn("a multiple-choice question has no stem text - skipped")
        return None

    pos_fb = neg_fb = []
    fb = mc.find("Feedback")
    if fb is not None:
        p, n = fb.find("RichPositiveFeedback"), fb.find("RichNegativeFeedback")
        pos_fb = segments(p) if p is not None else []
        neg_fb = segments(n) if n is not None else []

    options = []
    for opt in mc.findall("Option"):
        dist = opt.find("Distractor")
        text = join_segs(rich_texts(dist)) if dist is not None else []
        correct = opt.get("IsCorrect", "false").lower() == "true"
        own = opt.find("Feedback")  # per-option feedback, if a file has it
        own_fb = join_segs(rich_texts(own)) if own is not None else []
        options.append({"text": text, "correct": correct,
                        "feedback": own_fb or (pos_fb if correct else neg_fb)})

    if len(options) < 2:
        warn(f"question '{plain(prompt[0])[:50]}...' has fewer than 2 options - skipped")
        return None
    n_correct = sum(o["correct"] for o in options)
    if n_correct == 0:
        warn(f"question '{plain(prompt[0])[:50]}...' has NO correct option marked - ported anyway, fix by hand")

    return {
        "prompt": prompt,  # list of segment lists, one per paragraph
        "options": options,
        "checkbox": n_correct > 1,
        # Source preserves option order only when PreserveChoiceOrder="true"
        "shuffle": mc.get("PreserveChoiceOrder", "false").lower() != "true",
    }


# --------------------------------------------------------------------------
# Write into the Word template
# --------------------------------------------------------------------------
PARTIAL_CREDIT = False  # set by --partial-credit
NO_SHUFFLE = False      # set by --no-shuffle


def heading_for(n, q):
    props = ["checkbox" if q["checkbox"] else "multiple choice"]
    if q["shuffle"] and not NO_SHUFFLE:
        props.append("shuffle")
    if q["checkbox"] and PARTIAL_CREDIT:
        props.append("partial credit")
    return f"Question {n} - {', '.join(props)}"


def build_docx(questions, template, out_path):
    doc = Document(str(template))
    body = doc.element.body
    paras = list(doc.paragraphs)
    start = next((p for p in paras if START_MARK in p.text.lower()), None)
    end = next((p for p in paras if END_MARK in p.text.lower()), None)
    if start is None or end is None:
        raise SystemExit("Template is missing the 'Importable content starts here' / "
                         "'End of importable content' lines.")

    # 1. delete the example content between the two markers
    removed_ids = set()
    node = start._p.getnext()
    while node is not None and node is not end._p:
        nxt = node.getnext()
        for ref in node.iter(f"{{{W}}}commentReference"):
            removed_ids.add(ref.get(f"{{{W}}}id"))
        body.remove(node)
        node = nxt

    # 2. insert generated questions before the end marker
    def add(content="", style=None, prefix=""):
        """content: plain str or a segment list; prefix is always unformatted."""
        p = end.insert_paragraph_before("", style=style)
        if prefix:
            p.add_run(prefix)
        if isinstance(content, str):
            if content:
                p.add_run(content)
        else:
            for text, b, i, u in content:
                r = p.add_run(text)
                r.bold = b or None
                r.italic = i or None
                r.underline = u or None
        return p

    add()
    for i, q in enumerate(questions, 1):
        add(heading_for(i, q), "Heading 4")
        for line in q["prompt"]:
            add(line)
        add()
        for j, o in enumerate(q["options"]):
            letter = chr(ord("A") + j)
            add(o["text"], prefix=f"{'*' if o['correct'] else ''}{letter}: ")
            if o["feedback"]:
                add(o["feedback"], prefix="Feedback: ")
            add()
    doc.save(str(out_path))
    prune_comments(out_path, removed_ids)


def prune_comments(docx_path, removed_ids):
    """Drop the template's example margin comments that were deleted with the examples."""
    if not removed_ids:
        return
    tmp = Path(tempfile.mkdtemp()) / "x.docx"
    with zipfile.ZipFile(docx_path) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}
        infos = {n: zin.getinfo(n) for n in names}

    para_ids = set()
    if "word/comments.xml" in data:
        root = etree.fromstring(data["word/comments.xml"])
        for c in root.findall(f"{{{W}}}comment"):
            if c.get(f"{{{W}}}id") in removed_ids:
                for p in c.iter(f"{{{W}}}p"):
                    pid = p.get(f"{{{W14}}}paraId")
                    if pid:
                        para_ids.add(pid)
                root.remove(c)
        data["word/comments.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def prune(part, ns, tag):
        if part in data:
            r = etree.fromstring(data[part])
            for e in r.findall(f"{{{ns}}}{tag}"):
                if e.get(f"{{{ns}}}paraId") in para_ids:
                    r.remove(e)
            data[part] = etree.tostring(r, xml_declaration=True, encoding="UTF-8", standalone=True)

    prune("word/commentsExtended.xml", W15, "commentEx")
    prune("word/commentsIds.xml", W16CID, "commentId")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    shutil.move(str(tmp), docx_path)


# --------------------------------------------------------------------------
def convert(xml_path, out_path, template):
    msgs = []
    warn = lambda m: msgs.append(m)
    questions = parse_xml(xml_path, warn)
    for m in msgs:
        print(f"  WARNING: {m}")
    if not questions:
        print(f"  no knowledge-check questions found in {xml_path.name}; nothing written")
        return False
    build_docx(questions, template, out_path)
    print(f"  {len(questions)} question(s) -> {out_path}")
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xml", help="Xyleme WebCourse XML file")
    ap.add_argument("template", help="Coursera Assignment Import template (.docx)")
    ap.add_argument("output", help="output .docx to create")
    ap.add_argument("--partial-credit", action="store_true",
                    help="add 'partial credit' to multiple-correct (checkbox) questions")
    ap.add_argument("--no-shuffle", action="store_true", help="never add 'shuffle'")
    a = ap.parse_args()
    global PARTIAL_CREDIT, NO_SHUFFLE
    PARTIAL_CREDIT, NO_SHUFFLE = a.partial_credit, a.no_shuffle

    xml, template, out = Path(a.xml), Path(a.template), Path(a.output)
    for label, p in (("XML file", xml), ("Template", template)):
        if not p.is_file():
            sys.exit(f"{label} not found: {p}")
    if out.suffix.lower() != ".docx":
        sys.exit(f"Output must be a .docx file: {out}")
    if out.resolve() in (xml.resolve(), template.resolve()):
        sys.exit("Output path must differ from the XML and template paths")

    print(xml.name)
    try:
        ok = convert(xml, out, template)
    except ET.ParseError as e:
        sys.exit(f"  ERROR: could not parse XML ({e})")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
