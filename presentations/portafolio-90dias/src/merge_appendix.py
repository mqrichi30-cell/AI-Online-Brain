"""Anexa la diapositiva original (tabla completa) al final del mazo nuevo."""
import re
import shutil
import zipfile

BASE = "new.pptx"          # mazo generado
SRC = "original.pptx"      # mazo original
OUT = "Portafolio_90dias_v3_legible.pptx"

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

src = zipfile.ZipFile(SRC)
base = zipfile.ZipFile(BASE)

appendix_xml = src.read("ppt/slides/slide1.xml").decode("utf-8")

# --- marca de anexo: un shape extra en la zona libre de abajo a la derecha
label = (
    '<p:sp><p:nvSpPr><p:cNvPr id="900" name="Appendix Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
    '<p:spPr><a:xfrm><a:off x="7955280" y="5806440"/><a:ext cx="3840480" cy="228600"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>'
    '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>'
    '<a:lstStyle/><a:p><a:pPr algn="r" indent="0" marL="0"><a:buNone/></a:pPr>'
    '<a:r><a:rPr lang="en-US" sz="1400" b="1" spc="120" dirty="0">'
    '<a:solidFill><a:srgbClr val="9FB4D2"/></a:solidFill>'
    '<a:latin typeface="Calibri" pitchFamily="34" charset="0"/>'
    '<a:ea typeface="Calibri" pitchFamily="34" charset="-122"/>'
    '<a:cs typeface="Calibri" pitchFamily="34" charset="-120"/></a:rPr>'
    "<a:t>APPENDIX · FULL TABLE</a:t></a:r>"
    '<a:endParaRPr lang="en-US" sz="1400" dirty="0"/></a:p></p:txBody></p:sp>'
)
assert appendix_xml.count("</p:spTree>") == 1
appendix_xml = appendix_xml.replace("</p:spTree>", label + "</p:spTree>")

# --- numero de la nueva diapositiva
existing = [
    int(m.group(1))
    for n in base.namelist()
    for m in [re.fullmatch(r"ppt/slides/slide(\d+)\.xml", n)]
    if m
]
new_no = max(existing) + 1
new_part = f"ppt/slides/slide{new_no}.xml"

# el fondo del anexo es la misma imagen que ya empaqueta el mazo nuevo
bg_target = "../media/image-1-1.png"
assert "ppt/media/image-1-1.png" in base.namelist()

appendix_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{REL_NS}/image" Target="{bg_target}"/>'
    f'<Relationship Id="rId2" Type="{REL_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    "</Relationships>"
)

# --- content types
ct = base.read("[Content_Types].xml").decode("utf-8")
override = (
    f'<Override PartName="/{new_part}" '
    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
)
assert override not in ct
ct = ct.replace("</Types>", override + "</Types>")

# --- relacion desde presentation.xml
prels = base.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
used = {int(m) for m in re.findall(r'Id="rId(\d+)"', prels)}
new_rid = f"rId{max(used) + 1}"
prels = prels.replace(
    "</Relationships>",
    f'<Relationship Id="{new_rid}" Type="{REL_NS}/slide" Target="slides/slide{new_no}.xml"/>'
    "</Relationships>",
)

# --- registro en la lista de diapositivas
pres = base.read("ppt/presentation.xml").decode("utf-8")
ids = [int(m) for m in re.findall(r'<p:sldId id="(\d+)"', pres)]
new_sld_id = max(ids) + 1
pres = pres.replace(
    "</p:sldIdLst>",
    f'<p:sldId id="{new_sld_id}" r:id="{new_rid}"/></p:sldIdLst>',
)

replacements = {
    "[Content_Types].xml": ct.encode("utf-8"),
    "ppt/_rels/presentation.xml.rels": prels.encode("utf-8"),
    "ppt/presentation.xml": pres.encode("utf-8"),
}

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as out:
    for item in base.infolist():
        if item.filename.endswith("/"):
            continue
        out.writestr(item.filename, replacements.get(item.filename, base.read(item.filename)))
    out.writestr(new_part, appendix_xml.encode("utf-8"))
    out.writestr(f"ppt/slides/_rels/slide{new_no}.xml.rels", appendix_rels.encode("utf-8"))

print(f"escrito {OUT} · anexo como slide{new_no}")
