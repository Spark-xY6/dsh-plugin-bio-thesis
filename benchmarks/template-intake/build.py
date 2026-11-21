"""Deterministically derive three synthetic school packages from a synthetic base."""
import argparse
import hashlib
import json
import struct
import zlib
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
from xml.sax.saxutils import escape


def png(rgb):
    def chunk(tag, data):
        return struct.pack(">I",len(data))+tag+data+struct.pack(">I",zlib.crc32(tag+data))
    return b"\x89PNG\r\n\x1a\n"+chunk(b"IHDR",struct.pack(">IIBBBBB",1,1,8,2,0,0,0))+chunk(b"IDAT",zlib.compress(bytes([0,*rgb])))+chunk(b"IEND",b"")


def text(s):
    return '<w:r><w:t xml:space="preserve">'+escape(s)+'</w:t></w:r>'


def par(*runs):
    return "<w:p>"+''.join(text(s) for s in runs)+"</w:p>"


def table(*cells):
    row="".join("<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr>"+p+"</w:tc>" for p in cells)
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>'+('<w:gridCol w:w="2400"/>'*len(cells))+'</w:tblGrid><w:tr>'+row+'</w:tr></w:tbl>'


CASES=[
    ("alpine","docx","Tracing Orchard Shadows","Landscape Studies","AX-SAMPLE-61","MOCK-111","Alpine working papers","Sampling light across the valley.",(12240,15840,1660,1390,1570,1600,"portrait"),"AlpineBody",23,(24,111,60)),
    ("bay","dotx","Tidal Calendars and Craft","Material History","BZ-SAMPLE-82","MOCK-222","Bay collection","A working record of tidal markers.",(11906,16838,1430,1470,1540,1720,"portrait"),"BayDraft",29,(35,97,192)),
    ("ridge","dotx","A Map of Small Bridges","Spatial Arts","CW-SAMPLE-73","MOCK-333","Ridge series","Bridge fragments form a connected archive.",(15840,12240,1380,1590,1610,1510,"landscape"),"RidgeText",31,(180,92,31))]


def main():
    p=argparse.ArgumentParser()
    p.add_argument("base",type=Path)
    p.add_argument("out",type=Path)
    args=p.parse_args()
    if hashlib.sha256(args.base.read_bytes()).hexdigest() != "23fb2c6ad266cfdae554f5d5c0922a7911f64cef0890bc7599d95824511b5b8b":
        raise ValueError("Unexpected source template digest")
    args.out.mkdir(parents=True,exist_ok=False)
    with ZipFile(args.base) as z:
        original={n:z.read(n) for n in z.namelist()}
    index=[]
    for ident,ext,title,major,author,number,header,words,geom,sty,size,color in CASES:
        w,h,top,right,bottom,left,orient=geom
        items=original.copy()
        doc=items["word/document.xml"].decode()
        pos=doc.index("<w:body>")+len("<w:body>")
        end=doc.index("<w:p><w:r><w:drawing>",pos)
        title_p=par("Working title: ","{{title_","en}}")
        author_p=par("Candidate: ","{{au","thor}}")
        number_p=par("Registration: ","{{student_","id}}")
        major_p=par("Programme: ","{{major}}")
        cover={"alpine":title_p+table(author_p,number_p)+major_p,
               "bay":table(title_p,major_p)+author_p+number_p,
               "ridge":author_p+table(number_p,major_p)+title_p}[ident]
        doc=doc[:pos]+cover+doc[end:]
        doc=doc.replace('w:w="12240" w:h="15840"',f'w:w="{w}" w:h="{h}" w:orient="{orient}"')
        doc=doc.replace('w:top="1740" w:right="1320" w:bottom="1500" w:left="1540"',f'w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}"')
        doc=doc.replace("BodyCampus",sty).replace("Unchanged ordinary paragraph.",words)
        items["word/document.xml"]=doc.encode()
        items["word/styles.xml"]=items["word/styles.xml"].replace(b"BodyCampus",sty.encode()).replace(b"Campus body",(ident.title()+" body").encode()).replace(b'w:val="26"',f'w:val="{size}"'.encode())
        items["word/header1.xml"]=items["word/header1.xml"].replace(b"Campus archive",header.encode())
        if ident!="alpine":
            items["word/header1.xml"]=items["word/header1.xml"].replace(b"</w:hdr>",(par("Series title: ","{{title_en}}")+"</w:hdr>").encode())
        items["word/media/mark.png"]=png(color)
        if ext=="docx":
            items["[Content_Types].xml"]=items["[Content_Types].xml"].replace(b"wordprocessingml.template.main+xml",b"wordprocessingml.document.main+xml")
        name=f"{ident}.{ext}"
        with ZipFile(args.out/name,"x",compression=ZIP_DEFLATED) as z:
            for n,data in items.items():
                i=ZipInfo(n,(2020,1,1,0,0,0)); i.compress_type=ZIP_DEFLATED
                z.writestr(i,data)
        index.append({"template":name,"output":ident+".docx","title_en":title,"major":major,"author":author,"student_id":number})
    (args.out/"intake.json").write_text(json.dumps(index,indent=2,ensure_ascii=False)+"\n",encoding="utf-8")


if __name__=="__main__":
    main()
