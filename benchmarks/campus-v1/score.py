"""Score the synthetic Campus benchmark using public requirements and fixed labels. Stdlib only."""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path, PurePosixPath
import posixpath
import re
from urllib.parse import unquote
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

ROOT = Path(__file__).resolve().parent
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PK = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
EXCEPTIONS = (OSError, ValueError, BadZipFile, KeyError, ET.ParseError)


def read_json(file):
    return json.loads(file.read_text(encoding="utf-8-sig"))


def texts(node):
    return "".join(t.text or "" for t in node.iter(W + "t"))


def fields(tree):
    """Return names of fields with a simple declaration or a balanced complex begin/end pair."""
    names = []
    stack = []
    for e in tree.iter():
        if e.tag == W + "fldSimple":
            names.append(e.get(W + "instr", "").strip().split(maxsplit=1)[:1])
        elif e.tag == W + "fldChar":
            kind = e.get(W + "fldCharType")
            if kind == "begin":
                stack.append([[], False])
            elif kind == "separate" and stack:
                stack[-1][1] = True
            elif kind == "end" and stack:
                instructions, _separated = stack.pop()
                names.append("".join(instructions).strip().split(maxsplit=1)[:1])
        elif e.tag == W + "instrText" and stack and not stack[-1][1]:
            stack[-1][0].append(e.text or "")
    return {name[0].upper() for name in names if name}


class Document:
    def __init__(self, file):
        with ZipFile(file) as zin:
            info = zin.infolist()
            if len(info) > 2000 or sum(x.file_size for x in info) > 64 * 1024 * 1024:
                raise ValueError("Generated package exceeds the small-fixture scoring limits.")
            if len({i.filename for i in info}) != len(info) or zin.testzip() is not None:
                raise ValueError("Invalid ZIP members")
            self.parts = {x.filename: zin.read(x) for x in info}
        self.xml = {n: ET.fromstring(v) for n, v in self.parts.items() if n.endswith((".xml", ".rels"))}
        self.main = self.xml["word/document.xml"]
        if self.main.tag != W + "document":
            raise ValueError("Unexpected Word document root")
        self.body = self.main.find(W + "body")
        if self.body is None:
            raise ValueError("Missing document body")
        self.paras = list(self.body.iter(W + "p"))
        self.sections = list(self.body.iter(W + "sectPr"))
        self.rels = {}
        for e in self.xml.get("word/_rels/document.xml.rels", ET.Element("empty")).findall(PK + "Relationship"):
            if e.get("TargetMode") == "External":
                continue
            target = unquote(e.attrib["Target"])
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("word", target))
            self.rels[e.attrib["Id"]] = (e.get("Type", "").rsplit("/", 1)[-1], resolved)
        self.styles = {e.get(W + "styleId"): e for e in self.xml.get("word/styles.xml", ET.Element("empty")).findall(W + "style")}

    def section_parts(self, kind):
        result = []
        previous = None
        for section in self.sections:
            for ref in section.findall(W + kind + "Reference"):
                if ref.get(W + "type") == "default":
                    typ, name = self.rels.get(ref.get(R + "id"), (None, None))
                    previous = self.xml.get(name) if typ == kind else None
            result.append(previous)
        return result

    def geometry(self):
        for s in self.sections:
            size, margin = s.find(W + "pgSz"), s.find(W + "pgMar")
            if size is None or margin is None:
                return False
            if size.get(W + "orient", "portrait") != "portrait" or [size.get(W + a) for a in ("w", "h")] != ["12240", "15840"]:
                return False
            if [margin.get(W + a) for a in ("top", "right", "bottom", "left")] != ["1740", "1320", "1500", "1540"]:
                return False
        return len(self.sections) == 2

    def numbering(self):
        configs = [s.find(W + "pgNumType") for s in self.sections]
        return len(configs) == 2 and all(e is not None and e.get(W + "fmt", "decimal") == fmt and e.get(W + "start") == "1"
                                          for e, fmt in zip(configs, ["upperRoman", "decimal"]))

    def page_fields(self):
        footers = self.section_parts("footer")
        return len(footers) == 2 and all(f is not None and "PAGE" in fields(f) for f in footers)

    def campus_styles(self):
        return {key for key, e in self.styles.items() if
                key == "BodyCampus" or (e.find(W + "name") is not None and e.find(W + "name").get(W + "val") == "Campus body")}

    def style_size(self, key):
        seen = set()
        while key in self.styles and key not in seen:
            seen.add(key)
            style = self.styles[key]
            size = style.find(f"{W}rPr/{W}sz")
            if size is not None:
                return size.get(W + "val")
            base = style.find(W + "basedOn")
            key = base.get(W + "val") if base is not None else None
        return None

    def direct(self):
        anchors = [p for p in self.paras if texts(p) == "Anchor paragraph."]
        if len(anchors) != 1:
            return False
        style = anchors[0].find(f"{W}pPr/{W}pStyle")
        return (style is not None and style.get(W + "val") in self.campus_styles()
                and not any(e.tag in (W + "sz", W + "szCs") for e in anchors[0].iter()))

    def media(self):
        return {n: hashlib.sha256(v).hexdigest() for n, v in self.parts.items() if n.startswith("word/media/")}

    def media_instance(self):
        media = self.media()
        return any(self.rels.get(e.get(R + "embed"), (None, None))[0] == "image"
                   and self.rels[e.get(R + "embed")][1] in media for e in self.body.iter(A + "blip"))

    def identity_free(self):
        token = "example-q17"
        for tree in self.xml.values():
            for elem in tree.iter():
                if any(token in value.casefold() for value in elem.attrib.values()) or token in (elem.text or "").casefold():
                    return False
            if any(token in texts(p).casefold() for p in tree.iter(W + "p")):
                return False
        return True

    def independent_labels(self, baseline):
        valid = {
            "PAGE": self.geometry(), "NUMBER": self.numbering() and self.page_fields(),
            "STYLE": bool(self.campus_styles()) and all(self.style_size(s) == "26" for s in self.campus_styles()),
            "DIRECT": self.direct(), "MEDIA": self.media() == baseline.media(),
            "IDENTITY": self.identity_free(), "TOC": "TOC" in fields(self.body),
        }
        return {k for k, v in valid.items() if not v}

    def generation(self, baseline):
        bad = self.independent_labels(baseline)
        strings = [texts(p) for p in self.paras]
        ct = self.xml.get("[Content_Types].xml", ET.Element("empty"))
        package_links = self.xml.get("_rels/.rels", ET.Element("empty")).findall(PK + "Relationship")
        office_document = any(e.get("Type", "").rsplit("/", 1)[-1] == "officeDocument"
                              and e.get("TargetMode", "Internal") == "Internal"
                              and posixpath.normpath(e.get("Target", "").lstrip("/")) == "word/document.xml"
                              for e in package_links)
        return {
            "valid_docx": office_document and any(e.get("PartName") == "/word/document.xml" and e.get("ContentType") == DOCX_TYPE for e in ct.findall(CT + "Override")),
            "title_replaced": "Forms of Evidence" in strings and not any("{{title_en}}" in x for x in strings),
            "author_masked": any(re.match(r"^\s*Author\s*[:：]", x, re.I) and re.search(r"anonymous|anonymi[sz]ed|redacted|withheld|匿名|隐去", x, re.I) for x in strings)
                             and not any("{{author}}" in x for x in strings),
            "no_private_identity": self.identity_free(),
            "original_text": {"1  Materials", "Anchor paragraph.", "Unchanged ordinary paragraph."}.issubset(strings),
            "page_geometry": self.geometry(), "section_numbering": self.numbering(), "page_fields": self.page_fields(),
            "toc_field": "TOC" in fields(self.body), "body_style": "STYLE" not in bad and self.direct(),
            "embedded_media": self.media() == baseline.media() and self.media_instance(),
            "header_text": len(self.sections) == 2 and all(h is not None and "Campus archive" in texts(h) for h in self.section_parts("header")),
        }


def audit_scores(data, design):
    labels = set(design["labels"])
    rows = data.get("cases", {}) if isinstance(data, dict) else {}
    if not isinstance(rows, dict):
        rows = {}
    totals, by_label = Counter(), {k: Counter() for k in design["labels"]}
    extra = sorted(set(rows) - set(design["cases"]))
    missed, spurious, invalid = [], [], []
    for file, truth in design["cases"].items():
        found = rows.get(file)
        well_formed = isinstance(found, list) and all(isinstance(v, str) and v in labels for v in found)
        if not well_formed:
            invalid.append(file)
        for label in design["labels"]:
            expected = label in truth
            # Missing/invalid rows cannot earn TNs by abstaining.
            reported = label in found if well_formed else not expected
            outcome = ("TP" if reported else "FN") if expected else ("FP" if reported else "TN")
            totals[outcome] += 1
            by_label[label][outcome] += 1
            if outcome in ("FN", "FP"):
                (missed if expected else spurious).append([file, label])
    def metrics(counts):
        tp, fp, fn, tn = (counts[k] for k in ("TP", "FP", "FN", "TN"))
        return {"TP": tp, "FP": fp, "FN": fn, "TN": tn, "micro_F1": 2*tp/(2*tp+fp+fn) if 2*tp+fp+fn else 0,
                "errors": fp+fn, "accuracy": (tp+tn)/(tp+tn+fp+fn)}
    return {**metrics(totals), "by_label": {k: metrics(v) for k,v in by_label.items()}, "missed": missed,
            "spurious": spurious, "invalid_rows": invalid, "extra_rows": extra, "audit_complete": not invalid and not extra}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, help="Folder containing the solver's result.json and generated.docx")
    parser.add_argument("--self-check", action="store_true", help="Verify frozen bytes and cross-check labels via an independent reader")
    args = parser.parse_args()
    design = read_json(ROOT / "design.json")
    inputs = ROOT / "inputs"
    for name, wanted in read_json(ROOT / "input-sha256.json").items():
        if PurePosixPath(name).name != name or hashlib.sha256((inputs / name).read_bytes()).hexdigest() != wanted:
            raise ValueError("Input mismatch: "+name)
    base = Document(inputs / "school.docx")
    if args.self_check:
        for name, gold in design["cases"].items():
            observed = Document(inputs / (name + ".docx")).independent_labels(base)
            if observed != set(gold):
                raise ValueError(f"{name}: independent interpretation mismatch")
        print("SELF_CHECK=PASS 8 documents, 56 labels, 14 positive labels, 2 clean controls")
        if args.output is None:
            return
    if args.output is None:
        parser.error("--output or --self-check required")
    try:
        data = read_json(args.output / "result.json")
    except (OSError, ValueError):
        data = {}
    report = {"audit": audit_scores(data, design)}
    try:
        output_doc = Document(args.output / "generated.docx")
        generated = output_doc.generation(base)
        generation_status = "scored"
    except EXCEPTIONS:
        generated = {k: False for k in design["generation_items"]}
        generation_status = "unreadable" if (args.output / "generated.docx").exists() else "missing"
    if set(generated) != set(design["generation_items"]):
        raise AssertionError("Generation criteria mismatch")
    report["generation"] = {"status": generation_status, "score": sum(generated.values()), "total": len(generated), "items": generated}
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
