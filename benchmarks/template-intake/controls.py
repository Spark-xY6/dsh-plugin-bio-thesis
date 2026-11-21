"""Independent mutations to validate the judge before any contestants see the task."""
import argparse
import json
import re
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from judge import inspect


def main():
    p=argparse.ArgumentParser()
    p.add_argument("inputs",type=Path)
    args=p.parse_args()
    rows=json.loads((args.inputs/"intake.json").read_text(encoding="utf-8"))
    controls=0
    with TemporaryDirectory() as d:
        for row in rows:
            original=args.inputs/row["template"]
            with ZipFile(original) as z:
                base={n:z.read(n) for n in z.namelist()}
            base["[Content_Types].xml"]=base["[Content_Types].xml"].replace(b"template.main+xml",b"document.main+xml")
            for n in ("word/document.xml","word/header1.xml"):
                s=base[n].decode()
                for a,b in (("{{title_","en}}"),("{{student_","id}}"),("{{au","thor}}")):
                    s=re.sub(re.escape(a)+r'</w:t></w:r><w:r><w:t xml:space="preserve">'+re.escape(b),lambda _:a+b,s)
                for key in ("author","student_id","major","title_en"):
                    s=s.replace("{{"+key+"}}","Anonymous" if key in ("author","student_id") else row[key])
                base[n]=s.encode()
            variants=[("positive",dict(base),(1,1,1,1))]
            lower=dict(base); lower["word/document.xml"]=lower["word/document.xml"].replace(row["title_en"].encode(),row["title_en"].lower().encode())
            variants.append(("exact-case",lower,(1,0,1,1)))
            missing=dict(base); del missing["[Content_Types].xml"]
            variants.append(("missing-manifest",missing,(0,1,1,1)))
            rel=dict(base); relname="word/_rels/document.xml.rels"
            rel[relname]=re.sub(rb'(<Relationship Id=.*?/>)((?:<Relationship Id=.*?/>)*)</Relationships>',rb'\2\1</Relationships>',rel[relname]).replace(b'rPic"',b'rNewImage"')
            rel["word/document.xml"]=rel["word/document.xml"].replace(b'rPic"',b'rNewImage"')
            variants.append(("relationship-equivalence",rel,(1,1,1,1)))
            leaked=dict(base); leaked["docProps/core.xml"]=leaked["docProps/core.xml"].replace(b'Campus examples',row["author"].encode())
            variants.append(("metadata-identity",leaked,(1,1,0,1)))
            damage=dict(base); damage["word/media/mark.png"]+=b"x"
            variants.append(("media-change",damage,(1,1,1,0)))
            for name,items,flags in variants:
                file=Path(d)/(row["output"]+"-"+name+".docx")
                with ZipFile(file,"x") as z:
                    for member,content in items.items():
                        z.writestr(member,content)
                got=tuple(inspect(original,file,row).values())
                if got!=flags:
                    raise ValueError(f"{name} {row['output']} expected {flags}, observed {got}")
                controls+=1
    print(f"{controls} independent positive, negative, equivalence and partial-credit controls passed")


if __name__=="__main__":
    main()
