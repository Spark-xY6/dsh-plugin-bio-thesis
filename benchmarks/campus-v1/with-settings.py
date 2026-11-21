"""Make clearly distinguished copies for a checker that requires word/settings.xml."""
from pathlib import Path
from sys import argv
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

source, output = map(Path, argv[1:])
output.mkdir(exist_ok=True, parents=True)
docrel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
settings_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
for path in sorted(source.glob("A??.docx")):
    target = output / path.name
    if target.resolve() == path.resolve() or target.exists():
        raise ValueError("Expected a fresh, separate output")
    with ZipFile(path) as src, ZipFile(target, "x", compression=ZIP_DEFLATED) as dst:
        if "word/settings.xml" in src.namelist():
            raise ValueError("Unexpected settings in baseline")
        for member in src.infolist():
            data = src.read(member)
            if member.filename in ("[Content_Types].xml", "word/_rels/document.xml.rels"):
                root = ET.fromstring(data)
                ns = root.tag.split("}")[0] + "}"
                if member.filename == "[Content_Types].xml":
                    ET.SubElement(root, ns+"Override", PartName="/word/settings.xml", ContentType=settings_type)
                else:
                    if any(e.get("Id") == "rIdSettingsCompatibility" for e in root):
                        raise ValueError("Unexpected existing ID")
                    ET.SubElement(root, ns+"Relationship", Id="rIdSettingsCompatibility", Type=docrel, Target="settings.xml")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            dst.writestr(member, data)
        dst.writestr("word/settings.xml", '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
