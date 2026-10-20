#!/usr/bin/env python3
"""Build a synthetic test deck full of OBVIOUSLY FAKE identifiers (stdlib only).

    python3 tests/make_fixture.py [output.pptx]   # default: tests/fixtures/synthetic_case.pptx

Nothing here is real patient data. "Jane Testpatient", MRN 00012345, DOB 01/02/1950,
phone 555-0100 and *.example.com addresses are all invented.

Covers: a name split across text runs, speaker notes, a comment + comment author,
alt text, a hidden slide, docProps creator/company, a table, a footer (slide and
layout), a hyperlink, a thumbnail, a PNG with a text chunk, a JPEG clinical
photo carrying fake EXIF (camera, owner, date, serial number, GPS location), a
chart whose cached labels and embedded workbook hold a name, and an embedded
Word document (OLE, inside mc:AlternateContent). build(path, macro=True) makes a
macro-enabled (.pptm-style) variant with a fake vbaProject.bin.
"""

import os
import struct
import sys
import zipfile
import zlib
from xml.sax.saxutils import escape

# Every string the tests require to be gone from the scrubbed deck (checked case-insensitively).
FAKE_IDENTIFIERS = [
    "testpatient", "00012345", "01/02/1950", "555-0100", "fakename", "commenty",
    "hiddenslidesecret", "123 fake street", "92-year-old", "fake hospital", "testville",
    "springfield", "62704", "portal.example.com", "fakecam", "998877", "2024:01:02",
]

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
MC = ('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" mc:Ignorable="p14"')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT = "application/vnd.openxmlformats-officedocument.presentationml."


# ---------------------------------------------------------------- images
def _png(width, height, text_chunks=()):
    def chunk(tag, body):
        return struct.pack(">I", len(body)) + tag + body + struct.pack(">I", zlib.crc32(tag + body) & 0xFFFFFFFF)
    raw = b"".join(b"\x00" + bytes([(x * 4) % 256 for x in range(width)]) for _ in range(height))
    out = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 0, 0, 0, 0))
    for key, val in text_chunks:
        out += chunk(b"tEXt", key.encode("latin-1") + b"\x00" + val.encode("latin-1"))
    return out + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def _ifd(offset, entries):
    """entries: (tag, type, count, value_bytes). Returns IFD bytes placed at `offset` (big-endian)."""
    n = len(entries)
    data_start = offset + 2 + 12 * n + 4
    head, data = struct.pack(">H", n), b""
    for tag, typ, count, val in sorted(entries):
        if len(val) <= 4:
            head += struct.pack(">HHI", tag, typ, count) + val.ljust(4, b"\x00")
        else:
            head += struct.pack(">HHII", tag, typ, count, data_start + len(data))
            data += val + (b"\x00" if len(val) % 2 else b"")
    return head + struct.pack(">I", 0) + data


def _ascii(s):
    b = s.encode("ascii") + b"\x00"
    return (2, len(b), b)


def _rational3(a, b, c):
    return (5, 3, struct.pack(">IIIIII", a, 1, b, 1, c, 100))


def fake_exif():
    """A TIFF/EXIF block with fake camera, owner, dates, serial number and GPS position."""
    def ifd0(exif_off, gps_off):
        return [
            (0x010F,) + _ascii("FakeCam"), (0x0110,) + _ascii("FakeCam Derm 3000"),
            (0x0132,) + _ascii("2024:01:02 10:00:00"), (0x013B,) + _ascii("Fakename Photographer"),
            (0x8769, 4, 1, struct.pack(">I", exif_off)), (0x8825, 4, 1, struct.pack(">I", gps_off)),
        ]
    exif_entries = [(0x9003,) + _ascii("2024:01:02 10:00:00"), (0xA431,) + _ascii("SN-FAKE-998877")]
    gps_entries = [
        (0, 1, 4, bytes([2, 3, 0, 0])), (1,) + _ascii("N"), (2,) + _rational3(40, 26, 4612),
        (3,) + _ascii("W"), (4,) + _rational3(79, 58, 5512),
    ]
    first = _ifd(8, ifd0(0, 0))
    exif_off = 8 + len(first)
    exif = _ifd(exif_off, exif_entries)
    gps_off = exif_off + len(exif)
    gps = _ifd(gps_off, gps_entries)
    return b"MM\x00\x2a" + struct.pack(">I", 8) + _ifd(8, ifd0(exif_off, gps_off)) + exif + gps


def tiny_jpeg(exif=None, comment=None):
    """A valid 8x8 mid-grey baseline JPEG, optionally with APP1 EXIF and a COM segment."""
    def seg(marker, body):
        return b"\xff" + bytes([marker]) + struct.pack(">H", len(body) + 2) + body
    out = b"\xff\xd8" + seg(0xE0, b"JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00")
    if exif:
        out += seg(0xE1, b"Exif\x00\x00" + exif)
    if comment:
        out += seg(0xFE, comment.encode("latin-1"))
    out += seg(0xDB, b"\x00" + b"\x01" * 64)
    out += seg(0xC0, b"\x08" + struct.pack(">HH", 8, 8) + b"\x01" + b"\x01\x11\x00")
    out += seg(0xC4, b"\x00" + b"\x01" + b"\x00" * 15 + b"\x00")   # DC table: 1 code, symbol 0
    out += seg(0xC4, b"\x10" + b"\x01" + b"\x00" * 15 + b"\x00")   # AC table: 1 code, EOB
    out += seg(0xDA, b"\x01" + b"\x01\x00" + b"\x00\x3f\x00")
    return out + b"\x3f" + b"\xff\xd9"   # DC diff 0 ('0'), EOB ('0'), padded with 1s


# ---------------------------------------------------------------- XML helpers
def para(*runs):
    """runs: str or (text, rid) for a hyperlink run."""
    xml = "<a:p>"
    for r in runs:
        if isinstance(r, tuple):
            xml += '<a:r><a:rPr lang="en-US"><a:hlinkClick r:id="%s"/></a:rPr><a:t>%s</a:t></a:r>' % (r[1], escape(r[0]))
        else:
            xml += '<a:r><a:rPr lang="en-US" dirty="0"/><a:t>%s</a:t></a:r>' % escape(r)
    return xml + "</a:p>"


def xfrm(x, y, cx, cy, tag="a:xfrm"):
    return '<%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></%s>' % (tag, x, y, cx, cy, tag)


def sp(sid, name, paras, ph=None, box=(457200, 274638, 8229600, 1143000), descr=None):
    ph_xml = "<p:nvPr>%s</p:nvPr>" % (ph or "")
    d = ' descr="%s"' % escape(descr, {'"': "&quot;"}) if descr else ""
    body = '<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody>' % "".join(paras) if paras is not None else ""
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"%s/><p:cNvSpPr/>%s</p:nvSpPr>'
            '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>%s</p:sp>'
            % (sid, name, d, ph_xml, xfrm(*box), body))


def pic(sid, name, rid, descr, box):
    return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s" descr="%s"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/>'
            '</p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch>'
            '</p:blipFill><p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
            % (sid, name, escape(descr, {'"': "&quot;"}), rid, xfrm(*box)))


def table(sid, rows):
    cols = len(rows[0])
    grid = "".join('<a:gridCol w="%d"/>' % (8229600 // cols) for _ in range(cols))
    trs = ""
    for row in rows:
        tcs = "".join('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>%s</a:txBody><a:tcPr/></a:tc>' % para(c) for c in row)
        trs += '<a:tr h="370840">%s</a:tr>' % tcs
    return ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="Table %d"/><p:cNvGraphicFramePr>'
            '<a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>%s'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>'
            '<a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>%s</a:tblGrid>%s</a:tbl></a:graphicData></a:graphic>'
            '</p:graphicFrame>' % (sid, sid - 1, xfrm(457200, 1600200, 8229600, 1483360, "p:xfrm"), grid, trs))


GRP = ('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>'
       '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>'
       '</a:xfrm></p:grpSpPr>')
FOOTER_BOX = (3124200, 6356350, 2895600, 365125)
CLRMAP = ('<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
          'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>')


def slide(shapes, hidden=False, mc=False):
    return (DECL + '<p:sld %s%s%s><p:cSld><p:spTree>%s%s</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
            % (NS, (" " + MC) if mc else "", ' show="0"' if hidden else "", GRP, "".join(shapes)))


def rels(items):
    body = "".join('<Relationship Id="%s" Type="%s" Target="%s"%s/>' % (
        rid, typ, escape(tgt, {'"': "&quot;"}), ' TargetMode="External"' if ext else "")
        for rid, typ, tgt, ext in [(i + (False,))[:4] for i in items])
    return DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>' % body


def theme():
    def clr(n, v):
        return "<a:%s><a:srgbClr val=\"%s\"/></a:%s>" % (n, v, n)
    clrs = ('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
            + clr("dk2", "1F497D") + clr("lt2", "EEECE1") + clr("accent1", "4F81BD") + clr("accent2", "C0504D")
            + clr("accent3", "9BBB59") + clr("accent4", "8064A2") + clr("accent5", "4BACC6") + clr("accent6", "F79646")
            + clr("hlink", "0000FF") + clr("folHlink", "800080"))
    font = '<a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>'
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    return (DECL + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
            '<a:themeElements><a:clrScheme name="Office">%s</a:clrScheme><a:fontScheme name="Office"><a:majorFont>%s'
            '</a:majorFont><a:minorFont>%s</a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst>%s'
            '</a:fillStyleLst><a:lnStyleLst>%s</a:lnStyleLst><a:effectStyleLst>%s</a:effectStyleLst><a:bgFillStyleLst>%s'
            '</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
            % (clrs, font, font, fill * 3, ('<a:ln w="9525">%s</a:ln>' % fill) * 3,
               "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3, fill * 3))


def tiny_zip(files):
    import io
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            z.writestr(name, text)
    return buf.getvalue()


C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def chart_xml():
    def pt(i, v):
        return '<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v)
    return DECL + (
        '<c:chartSpace xmlns:c="%s" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:layout/>'
        '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/>'
        '<c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/>%s</c:strCache>'
        '</c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>%s%s'
        '</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>'
        '<c:formatCode>General</c:formatCode><c:ptCount val="2"/>%s%s</c:numCache></c:numRef></c:val></c:ser>'
        '<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/>'
        '</c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/>'
        '</c:valAx></c:plotArea></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'
        '</c:chartSpace>' % (C_NS, pt(0, "Lesion size (mm)"), pt(0, "Jane Testpatient"), pt(1, "Comparison"),
                             pt(0, "3"), pt(1, "5")))


def chart_frame(sid, rid):
    return ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="Chart %d"/><p:cNvGraphicFramePr/><p:nvPr/>'
            '</p:nvGraphicFramePr>%s<a:graphic><a:graphicData uri="%s"><c:chart xmlns:c="%s" r:id="%s"/>'
            '</a:graphicData></a:graphic></p:graphicFrame>'
            % (sid, sid - 1, xfrm(457200, 1600200, 4000000, 3000000, "p:xfrm"), C_NS, C_NS, rid))


def ole_frame(sid, rid):
    frame = ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="Object %d"/><p:cNvGraphicFramePr>'
             '<a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>%s'
             '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">'
             '<p:oleObj name="Document" r:id="%s" imgW="914400" imgH="914400" progId="Word.Document.12"><p:embed/>'
             '</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>'
             % (sid, sid - 1, xfrm(4800000, 1600200, 3000000, 3000000, "p:xfrm"), rid))
    return ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            '<mc:Choice Requires="p14">%s</mc:Choice><mc:Fallback>%s</mc:Fallback></mc:AlternateContent>' % (frame, frame))


# ---------------------------------------------------------------- deck
def build(path, macro=False):
    footer_ph = '<p:ph type="ftr" sz="quarter" idx="11"/>'
    title_ph = '<p:ph type="title"/>'
    body_ph = '<p:ph idx="1"/>'
    parts = {}

    parts["[Content_Types].xml"] = DECL + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
        '<Default Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"/>'
        '%(bin)s'
        '<Override PartName="/ppt/presentation.xml" ContentType="%(main)s"/>'
        '<Override PartName="/ppt/slides/slide5.xml" ContentType="%(ct)sslide+xml"/>'
        '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="%(ct)sslideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="%(ct)sslideLayout+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" ContentType="%(ct)sslide+xml"/>'
        '<Override PartName="/ppt/slides/slide2.xml" ContentType="%(ct)sslide+xml"/>'
        '<Override PartName="/ppt/slides/slide3.xml" ContentType="%(ct)sslide+xml"/>'
        '<Override PartName="/ppt/slides/slide4.xml" ContentType="%(ct)sslide+xml"/>'
        '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="%(ct)snotesMaster+xml"/>'
        '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="%(ct)snotesSlide+xml"/>'
        '<Override PartName="/ppt/notesSlides/notesSlide2.xml" ContentType="%(ct)snotesSlide+xml"/>'
        '<Override PartName="/ppt/comments/comment1.xml" ContentType="%(ct)scomments+xml"/>'
        '<Override PartName="/ppt/commentAuthors.xml" ContentType="%(ct)scommentAuthors+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/ppt/presProps.xml" ContentType="%(ct)spresProps+xml"/>'
        '<Override PartName="/ppt/viewProps.xml" ContentType="%(ct)sviewProps+xml"/>'
        '<Override PartName="/ppt/tableStyles.xml" ContentType="%(ct)stableStyles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>' % {
            "ct": CT,
            "main": ("application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml" if macro
                     else CT + "presentation.main+xml"),
            "bin": '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' if macro else ""})

    parts["_rels/.rels"] = rels([
        ("rId1", RT + "officeDocument", "ppt/presentation.xml"),
        ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
        ("rId3", RT + "extended-properties", "docProps/app.xml"),
        ("rId4", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail", "docProps/thumbnail.jpeg"),
    ])
    parts["docProps/core.xml"] = DECL + (
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Case of Jane Testpatient</dc:title><dc:creator>Fakename Author</dc:creator>'
        '<cp:keywords>MRN 00012345</cp:keywords><cp:lastModifiedBy>Fakename Modifier</cp:lastModifiedBy>'
        '<cp:revision>3</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T10:00:00Z</dcterms:created>'
        '</cp:coreProperties>')
    parts["docProps/app.xml"] = DECL + (
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Company>Fake Hospital of Testville</Company><PresentationFormat>On-screen Show (4:3)</PresentationFormat>'
        '<Slides>4</Slides><Notes>2</Notes><TotalTime>5</TotalTime><HiddenSlides>1</HiddenSlides>'
        '<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Calibri</vt:lpstr>'
        '<vt:lpstr>Case of Jane Testpatient</vt:lpstr></vt:vector></TitlesOfParts>'
        '<Application>Microsoft Office PowerPoint</Application><AppVersion>16.0000</AppVersion></Properties>')
    parts["docProps/thumbnail.jpeg"] = tiny_jpeg(comment="slide 1 render - Jane Testpatient")

    parts["ppt/presentation.xml"] = DECL + (
        '<p:presentation %s %s saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        '<p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rId10"/><p:sldId id="257" r:id="rId11"/>'
        '<p:sldId id="258" r:id="rId12"/><p:sldId id="259" r:id="rId13"/><p:sldId id="260" r:id="rId14"/></p:sldIdLst>'
        '<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/><p:notesSz cx="6858000" cy="9144000"/>'
        '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst>'
        '<p14:section name="Default Section" id="{6E5D4C3B-2A19-4F80-8E7D-6C5B4A392817}"><p14:sldIdLst>'
        '<p14:sldId id="256"/><p14:sldId id="257"/><p14:sldId id="258"/><p14:sldId id="259"/><p14:sldId id="260"/>'
        '</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>' % (NS, MC))
    parts["ppt/_rels/presentation.xml.rels"] = rels([
        ("rId1", RT + "slideMaster", "slideMasters/slideMaster1.xml"),
        ("rId2", RT + "notesMaster", "notesMasters/notesMaster1.xml"),
        ("rId3", RT + "theme", "theme/theme1.xml"),
        ("rId4", RT + "presProps", "presProps.xml"),
        ("rId5", RT + "viewProps", "viewProps.xml"),
        ("rId6", RT + "tableStyles", "tableStyles.xml"),
        ("rId7", RT + "commentAuthors", "commentAuthors.xml"),
        ("rId10", RT + "slide", "slides/slide1.xml"),
        ("rId11", RT + "slide", "slides/slide2.xml"),
        ("rId12", RT + "slide", "slides/slide3.xml"),
        ("rId13", RT + "slide", "slides/slide4.xml"),
        ("rId14", RT + "slide", "slides/slide5.xml"),
    ] + ([("rId20", "http://schemas.microsoft.com/office/2006/relationships/vbaProject", "vbaProject.bin")]
         if macro else []))
    if macro:
        parts["ppt/vbaProject.bin"] = b"FAKE VBA PROJECT - not a real macro"
    parts["ppt/presProps.xml"] = DECL + "<p:presentationPr %s/>" % NS
    parts["ppt/viewProps.xml"] = DECL + '<p:viewPr %s><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>' % NS
    parts["ppt/tableStyles.xml"] = DECL + ('<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                                           'def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>')
    parts["ppt/theme/theme1.xml"] = theme()
    parts["ppt/theme/theme2.xml"] = theme()

    parts["ppt/slideMasters/slideMaster1.xml"] = DECL + (
        '<p:sldMaster %s><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>%s%s%s'
        '</p:spTree></p:cSld>%s<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'
        % (NS, GRP, sp(2, "Title Placeholder 1", [para("Click to edit Master title style")], title_ph),
           sp(3, "Footer Placeholder 2", [para("footer")], footer_ph, FOOTER_BOX), CLRMAP))
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = rels([
        ("rId1", RT + "slideLayout", "../slideLayouts/slideLayout1.xml"),
        ("rId2", RT + "theme", "../theme/theme1.xml"),
    ])
    parts["ppt/slideLayouts/slideLayout1.xml"] = DECL + (
        '<p:sldLayout %s type="obj" preserve="1"><p:cSld name="Title and Content"><p:spTree>%s%s%s%s</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
        % (NS, GRP, sp(2, "Title 1", [para("Click to edit Master title style")], title_ph),
           sp(3, "Content Placeholder 2", [para("Click to edit Master text styles")], body_ph,
              (457200, 1600200, 8229600, 4525963)),
           sp(4, "Footer Placeholder 3", [para("Jane Testpatient | Fake Hospital of Testville")], footer_ph, FOOTER_BOX)))
    parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = rels([("rId1", RT + "slideMaster", "../slideMasters/slideMaster1.xml")])

    layout_rel = ("rId1", RT + "slideLayout", "../slideLayouts/slideLayout1.xml")

    # Slide 1: title + footer
    parts["ppt/slides/slide1.xml"] = slide([
        sp(2, "Title 1", [para("Grand Rounds: An Unusual Rash")], title_ph),
        sp(3, "Content Placeholder 2", [para("Presented by Dr Example Presenter")], body_ph, (457200, 1600200, 8229600, 2000000)),
        sp(4, "Footer Placeholder 3", [para("Jane Testpatient | MRN 00012345")], footer_ph, FOOTER_BOX),
    ])
    parts["ppt/slides/_rels/slide1.xml.rels"] = rels([layout_rel])

    # Slide 2: case details, run-split name, hyperlink, picture with alt text + PNG text chunk, notes, comment
    parts["ppt/slides/slide2.xml"] = slide([
        sp(2, "Title 1", [para("Case presentation")], title_ph),
        sp(3, "Content Placeholder 2", [
            para("Ja", "ne Test", "patient", " is a 92-year-old seen in clinic on 01/02/2024."),
            para("Patient: Jane Testpatient, DOB 01/02/1950, MRN: 00012345"),
            para("Phone 555-0100, email jane.testpatient@example.com"),
            para("Home: 123 Fake Street, Springfield, IL 62704"),
            para("Outside records: ", ("open chart", "rId3")),
        ], body_ph, (457200, 1600200, 5000000, 4525963)),
        pic(5, "Picture 4", "rId2", "Photo of Jane Testpatient forearm lesion and tattoo", (5600000, 1600200, 3000000, 2250000)),
    ], mc=True)
    parts["ppt/slides/_rels/slide2.xml.rels"] = rels([
        layout_rel,
        ("rId2", RT + "image", "../media/image1.png"),
        ("rId3", RT + "hyperlink", "https://portal.example.com/chart?mrn=00012345", True),
        ("rId4", RT + "notesSlide", "../notesSlides/notesSlide1.xml"),
        ("rId5", RT + "comments", "../comments/comment1.xml"),
    ])
    parts["ppt/media/image1.png"] = _png(64, 48, [("Author", "Jane Testpatient"), ("Comment", "MRN 00012345")])

    # Slide 3: table + clinical photo JPEG with fake EXIF GPS
    parts["ppt/slides/slide3.xml"] = slide([
        sp(2, "Title 1", [para("Labs and follow-up")], title_ph),
        table(4, [["Name", "DOB", "MRN"], ["Jane Testpatient", "01/02/1950", "00012345"]]),
        pic(6, "Picture 5", "rId2", "Dermoscopy, left forearm", (5600000, 3500000, 3000000, 2250000)),
    ])
    parts["ppt/slides/_rels/slide3.xml.rels"] = rels([layout_rel, ("rId2", RT + "image", "../media/image2.jpeg")])
    parts["ppt/media/image2.jpeg"] = tiny_jpeg(exif=fake_exif())

    # Slide 4: hidden backup slide with its own notes
    parts["ppt/slides/slide4.xml"] = slide([
        sp(2, "Title 1", [para("Backup")], title_ph),
        sp(3, "Content Placeholder 2", [para("HIDDENSLIDESECRET Jane Testpatient backup, MRN 00012345")], body_ph,
           (457200, 1600200, 8229600, 2000000)),
    ], hidden=True)
    parts["ppt/slides/_rels/slide4.xml.rels"] = rels([layout_rel, ("rId2", RT + "notesSlide", "../notesSlides/notesSlide2.xml")])

    # Slide 5: chart with embedded workbook + embedded Word document (OLE)
    parts["ppt/slides/slide5.xml"] = slide([
        sp(2, "Title 1", [para("Response over time")], title_ph),
        chart_frame(4, "rId2"),
        ole_frame(6, "rId3"),
    ], mc=True)
    parts["ppt/slides/_rels/slide5.xml.rels"] = rels([
        layout_rel,
        ("rId2", RT + "chart", "../charts/chart1.xml"),
        ("rId3", RT + "package", "../embeddings/Microsoft_Word_Document.docx"),
    ])
    parts["ppt/charts/chart1.xml"] = chart_xml()
    parts["ppt/charts/_rels/chart1.xml.rels"] = rels([
        ("rId1", RT + "package", "../embeddings/Microsoft_Excel_Worksheet.xlsx")])
    parts["ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"] = tiny_zip({
        "xl/sharedStrings.xml": '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>'
                                'Jane Testpatient</t></si><si><t>MRN 00012345</t></si></sst>'})
    parts["ppt/embeddings/Microsoft_Word_Document.docx"] = tiny_zip({
        "word/document.xml": '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                             '<w:body><w:p><w:r><w:t>Referral letter for Jane Testpatient, DOB 01/02/1950</w:t>'
                             '</w:r></w:p></w:body></w:document>'})

    # Notes
    parts["ppt/notesMasters/notesMaster1.xml"] = DECL + (
        '<p:notesMaster %s><p:cSld><p:spTree>%s</p:spTree></p:cSld>%s</p:notesMaster>' % (NS, GRP, CLRMAP))
    parts["ppt/notesMasters/_rels/notesMaster1.xml.rels"] = rels([("rId1", RT + "theme", "../theme/theme2.xml")])
    for i, (slide_no, text) in enumerate([(2, "Remember Jane Testpatient called from 555-0100 on Jan 5."),
                                          (4, "HIDDENSLIDESECRET notes for the backup slide")], 1):
        parts["ppt/notesSlides/notesSlide%d.xml" % i] = DECL + (
            '<p:notes %s><p:cSld><p:spTree>%s%s%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
            % (NS, GRP, sp(2, "Slide Image Placeholder 1", None, '<p:ph type="sldImg"/>', (1143000, 685800, 4572000, 3429000)),
               sp(3, "Notes Placeholder 2", [para(text)], '<p:ph type="body" idx="1"/>', (685800, 4343400, 5486400, 4114800))))
        parts["ppt/notesSlides/_rels/notesSlide%d.xml.rels" % i] = rels([
            ("rId1", RT + "notesMaster", "../notesMasters/notesMaster1.xml"),
            ("rId2", RT + "slide", "../slides/slide%d.xml" % slide_no),
        ])

    # Comments
    parts["ppt/commentAuthors.xml"] = DECL + (
        '<p:cmAuthorLst %s><p:cmAuthor id="0" name="Commenty Reviewer" initials="CR" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>' % NS)
    parts["ppt/comments/comment1.xml"] = DECL + (
        '<p:cmLst %s><p:cm authorId="0" dt="2024-01-02T10:00:00.000" idx="1"><p:pos x="10" y="10"/>'
        '<p:text>Is it OK to show Jane Testpatient MRN 00012345 here?</p:text></p:cm></p:cmLst>' % NS)

    order = ["[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml", "docProps/thumbnail.jpeg"]
    order += sorted(k for k in parts if k not in order)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with zipfile.ZipFile(path, "w") as z:
        for name in order:
            data = parts[name]
            if isinstance(data, str):
                data = data.encode("utf-8")
            z.writestr(zipfile.ZipInfo(name, (2024, 1, 2, 10, 0, 0)), data,
                       compress_type=zipfile.ZIP_STORED if name.endswith((".png", ".jpeg", ".xlsx", ".docx"))
                       else zipfile.ZIP_DEFLATED)
    return path


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "fixtures", "synthetic_case.pptx")
    print(build(out, macro=out.lower().endswith(".pptm")))
