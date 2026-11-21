"""Score only publicly specified fields and preserved template semantics."""
import argparse
import hashlib
import json
import posixpath
import re
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import BadZipFile, ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
OFF = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
DOC = "word/document.xml"
MASK = r"(?i:anonymous(?: (?:author|student|candidate))?|【匿名评审隐去】|匿名)"
SLOTS = re.compile(r"\{\{(author|student_id|title_en|major)\}\}")


def package(file):
    with ZipFile(file) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())):
            raise ValueError("Unreadable ZIP")
        items={n:z.read(n) for n in z.namelist()}
    xml={n:ET.fromstring(data) for n,data in items.items() if n.endswith((".xml",".rels"))}
    return items,xml


def props(e):
    return (e.tag,tuple(sorted(e.attrib.items())),(e.text or "").strip(),tuple(props(c) for c in e))


def relationship_values(root, owner):
    result={}
    if root is None:
        return result
    for e in root.findall(PKG+"Relationship"):
        target=e.get("Target","")
        mode=e.get("TargetMode","Internal")
        if mode=="Internal" and target:
            target=posixpath.normpath(posixpath.join(posixpath.dirname(owner),target)).lstrip("/")
        result[e.get("Id")]=(e.get("Type"),target,mode)
    return result


def paras(root):
    return ["".join(t.text or "" for t in p.iter(W+"t")) for p in root.iter(W+"p")]


def expected(p,entry):
    parts=[]; offset=0
    for m in SLOTS.finditer(p):
        parts.append(re.escape(p[offset:m.start()]))
        parts.append(MASK if m[1] in ("author","student_id") else re.escape(entry[m[1]]))
        offset=m.end()
    parts.append(re.escape(p[offset:]))
    return "".join(parts)


def shape(e,rels):
    # Remove ordinary textual runs (they can be merged at serialization). Meaningful
    # paragraph-level text, run properties and all non-text nodes are scored separately.
    if e.tag == W+"r" and all(c.tag in {W+"rPr",W+"t"} for c in e):
        return None
    children=tuple(x for c in e if (x:=shape(c,rels)) is not None)
    attrs=tuple(sorted((k,rels.get(v,v) if k.startswith(R) else v) for k,v in e.attrib.items()))
    return (e.tag,attrs,(e.text or "").strip(),children)


def runs(p):
    """Consecutive text with the same effective run properties can be coalesced."""
    previous=None; signatures=[]
    for r in p.iter(W+"r"):
        text="".join(t.text or "" for t in r.findall(W+"t"))
        if not text:
            continue
        rp=r.find(W+"rPr")
        signature=props(rp) if rp is not None else ()
        if signature!=previous:
            signatures.append(signature)
        previous=signature
    return signatures


def clean(value):
    return unicodedata.normalize("NFKC",value).casefold()


def inspect(original,dest,row,all_identities=()):
    base,bxml=package(original)
    keys=("docx","filled","private","preserved")
    try:
        data,xml=package(dest)
        types=xml.get("[Content_Types].xml")
        overrides=[] if types is None else [e.get("ContentType") for e in types.findall(CT+"Override") if e.get("PartName")=="/word/document.xml"]
        defaults=[] if types is None else [e.get("ContentType") for e in types.findall(CT+"Default") if e.get("Extension","").casefold()=="xml"]
        rels=relationship_values(xml.get("_rels/.rels"),"")
        office=[(k,v) for k,v in rels.items() if v[0]==OFF]
        body=xml.get(DOC)
        docx=(overrides or defaults)==[DOCX] and len(office)==1 and bool(office[0][0]) and office[0][1]==(OFF,DOC,"Internal") and body is not None and body.tag==W+"document" and body.find(W+"body") is not None and not any("vbaproject" in n.lower() for n in data)
        filled=True; preserved=True
        stories=[n for n in base if n==DOC or re.fullmatch(r"word/(?:header|footer)\d+\.xml",n)]
        for n in stories:
            actual=xml.get(n)
            if actual is None:
                filled &= not any(SLOTS.search(p) for p in paras(bxml[n]))
                preserved=False
                continue
            original_paras=paras(bxml[n]); new_paras=paras(actual)
            for i,old in enumerate(original_paras):
                if SLOTS.search(old):
                    filled &= i<len(new_paras) and re.fullmatch(expected(old,row),new_paras[i]) is not None
                else:
                    preserved &= i<len(new_paras) and old==new_paras[i]
            relpath=posixpath.join(posixpath.dirname(n),"_rels",posixpath.basename(n)+".rels")
            source_rels=relationship_values(bxml.get(relpath),n)
            dest_rels=relationship_values(xml.get(relpath),n)
            preserved &= len(original_paras)==len(new_paras) and shape(bxml[n],source_rels)==shape(actual,dest_rels)
            preserved &= [runs(p) for p in bxml[n].iter(W+"p")]==[runs(p) for p in actual.iter(W+"p")]
        for n in (x for x in base if x.startswith("word/") and x not in stories):
            if n.endswith(".rels"):
                owner=n.replace("/_rels/","/").removesuffix(".rels")
                preserved &= n in xml and sorted(relationship_values(bxml[n],owner).values())==sorted(relationship_values(xml.get(n),owner).values())
            elif n.endswith(".xml"):
                preserved &= n in xml and props(bxml[n])==props(xml[n])
            else:
                preserved &= data.get(n)==base[n]
        preserved &= {n for n in data if n.startswith("word/media/")}=={n for n in base if n.startswith("word/media/")}
        values=[v for root in xml.values() for node in root.iter() for v in (node.text or "",*node.attrib.values())]
        values.extend(t for root in xml.values() for t in paras(root))
        tokens=[row["author"],row["student_id"],*all_identities,"{{author}}","{{student_id}}"]
        private=all(clean(token) not in clean(value) for value in values for token in tokens)
        return dict(zip(keys,(docx,filled,private,preserved)))
    except (BadZipFile,ET.ParseError,FileNotFoundError,KeyError,ValueError,OSError):
        return dict.fromkeys(keys,False)


def main():
    p=argparse.ArgumentParser()
    p.add_argument("inputs",type=Path)
    p.add_argument("output",type=Path)
    a=p.parse_args()
    rows=json.loads((a.inputs/"intake.json").read_text(encoding="utf-8"))
    all_identities=[row[k] for row in rows for k in ("author","student_id")]
    results={}
    for r in rows:
        for k in ("template","output"):
            if Path(r[k]).name!=r[k]:
                raise ValueError("Non-local filename")
        f=a.output/r["output"]
        checks=inspect(a.inputs/r["template"],f,r,all_identities)
        results[r["output"]]={"checks":checks,"sha256":hashlib.sha256(f.read_bytes()).hexdigest() if f.is_file() else None}
    score=sum(sum(row["checks"].values()) for row in results.values())
    print(json.dumps({"score":score,"total":4*len(rows),"documents":results},indent=2))


if __name__=="__main__":
    main()
