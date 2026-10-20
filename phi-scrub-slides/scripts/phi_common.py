"""Shared helpers for phi-scrub-slides (PowerPoint and Google Slides exported as .pptx).

Python standard library only. A .pptx file is a zip of OOXML parts; this module
reads that package, walks every place text can hide, runs pattern detectors for
the HIPAA Safe Harbor identifiers that can be pattern-matched, and masks values
so nothing identifying is printed unless the user asks for it.
"""

import io
import posixpath
import re
import struct
import zipfile
import zlib
from urllib.parse import unquote
from xml.sax.saxutils import escape as _xml_escape
import xml.etree.ElementTree as ET

VERSION = "0.1.0"

# --------------------------------------------------------------------------
# Namespaces and small XML helpers
# --------------------------------------------------------------------------
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

A_P = "{%s}p" % A
A_T = "{%s}t" % A
A_TBL = "{%s}tbl" % A
A_TR = "{%s}tr" % A
A_TC = "{%s}tc" % A

PHI_WARNING = (
    "THIS FILE CONTAINS PHI (protected health information) copied from the deck. "
    "Keep it on this computer, do not email, upload or commit it, and delete it "
    "when you are done."
)


def local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) and "}" in tag else tag


def nsuri(tag):
    return tag[1:].split("}", 1)[0] if isinstance(tag, str) and tag.startswith("{") else ""


XML_EXTS = (".xml", ".rels", ".vml", ".svg")


def is_xml_name(name):
    return name.lower().endswith(XML_EXTS)


def rels_path_for(part):
    if part == "":
        return "_rels/.rels"
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def source_for_rels(rels_path):
    d, b = posixpath.split(rels_path)
    if not b.endswith(".rels") or posixpath.basename(d) != "_rels":
        return None
    parent = posixpath.dirname(d)
    name = b[:-5]
    if name == "":
        return ""
    return posixpath.join(parent, name) if parent else name


def resolve_target(source, target):
    target = unquote(target.split("#", 1)[0])
    if target.startswith("/"):
        path = target[1:]
    else:
        path = posixpath.join(posixpath.dirname(source), target)
    path = posixpath.normpath(path)
    return "" if path == "." else path.lstrip("/")


def relative_target(source, part):
    return posixpath.relpath(part, posixpath.dirname(source) or ".")


class XmlDoc:
    """An ElementTree root that serializes back with its original namespace
    declarations (OOXML relies on prefixes in mc:Ignorable / mc:Choice that
    ElementTree would otherwise drop)."""

    STD_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

    def __init__(self, data):
        m = re.match(rb"\s*(<\?xml[^>]*\?>)", data)
        decl = m.group(1).decode("ascii", "replace") if m else self.STD_DECL
        self.decl = decl if "utf-8" in decl.lower() else self.STD_DECL
        self.ns = []
        for _ev, item in ET.iterparse(io.BytesIO(data), events=("start-ns",)):
            if item not in self.ns:
                self.ns.append(item)
        self.root = ET.fromstring(data)

    def tobytes(self):
        for prefix, uri in self.ns:
            if prefix and not re.match(r"ns\d+$", prefix):
                try:
                    ET.register_namespace(prefix, uri)
                except ValueError:
                    pass
        default = None
        for prefix, uri in self.ns:
            if prefix == "":
                default = uri
                break
        try:
            body = ET.tostring(self.root, encoding="unicode", default_namespace=default)
        except ValueError:
            body = ET.tostring(self.root, encoding="unicode")
        end = body.index(">")
        head, rest = body[:end], body[end:]
        selfclose = head.endswith("/")
        if selfclose:
            head = head[:-1]
        present = dict(re.findall(r'\sxmlns(?::([\w.\-]+))?="([^"]*)"', head))
        extra = []
        for prefix, uri in self.ns:
            if prefix == "" or prefix in present:
                continue
            present[prefix] = uri
            extra.append(' xmlns:%s="%s"' % (prefix, _xml_escape(uri, {'"': "&quot;"})))
        head = head + "".join(extra) + ("/" if selfclose else "")
        return (self.decl + "\r\n" + head + rest).encode("utf-8")


class Rel:
    __slots__ = ("id", "type", "target", "mode", "resolved", "el", "source")

    def __init__(self, source, el):
        self.source = source
        self.el = el
        self.id = el.get("Id")
        self.type = el.get("Type", "")
        self.target = el.get("Target", "")
        self.mode = el.get("TargetMode", "Internal")
        self.resolved = None if self.mode == "External" else resolve_target(source, self.target)

    @property
    def external(self):
        return self.mode == "External"

    def is_type(self, *suffixes):
        return any(self.type.endswith("/" + s) for s in suffixes)


class Package:
    """In-memory view of a .pptx (never writes to the source file)."""

    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.infos = {}
            self.blobs = {}
            self.order = []
            for info in z.infolist():
                if info.is_dir():
                    continue
                self.infos[info.filename] = info
                self.blobs[info.filename] = z.read(info)
                self.order.append(info.filename)
        self._docs = {}
        self.parse_errors = {}
        self.dirty = set()
        self.deleted = set()
        self.added = []

    # -- parts --------------------------------------------------------------
    def exists(self, name):
        return name in self.blobs and name not in self.deleted

    def parts(self):
        return [n for n in self.order + self.added if self.exists(n)]

    def doc(self, name):
        if name not in self._docs:
            self._docs[name] = XmlDoc(self.blobs[name])
        return self._docs[name]

    def try_doc(self, name):
        try:
            return self.doc(name)
        except ET.ParseError as exc:
            self.parse_errors[name] = str(exc)
            return None

    def mark(self, name):
        self.dirty.add(name)

    def delete(self, name):
        if name in self.blobs:
            self.deleted.add(name)
            rp = rels_path_for(name)
            if rp in self.blobs:
                self.deleted.add(rp)

    def set_blob(self, name, data):
        self.blobs[name] = data
        self._docs.pop(name, None)
        self.dirty.discard(name)
        self.deleted.discard(name)
        if name not in self.order and name not in self.added:
            self.added.append(name)

    def get_bytes(self, name):
        if name in self.dirty and name in self._docs:
            return self._docs[name].tobytes()
        return self.blobs[name]

    # -- relationships -------------------------------------------------------
    def rels(self, source):
        rp = rels_path_for(source)
        if not self.exists(rp):
            return []
        d = self.try_doc(rp)
        if d is None:
            return []
        return [Rel(source, el) for el in list(d.root) if local(el.tag) == "Relationship"]

    def all_rels(self):
        out = []
        for name in self.parts():
            if name.endswith(".rels"):
                src = source_for_rels(name)
                if src is not None and (src == "" or self.exists(src)):
                    out.extend(self.rels(src))
        return out

    def remove_rel(self, rel):
        d = self.doc(rels_path_for(rel.source))
        if rel.el in list(d.root):
            d.root.remove(rel.el)
            self.mark(rels_path_for(rel.source))

    def main_part(self):
        for r in self.rels(""):
            if r.is_type("officeDocument"):
                return r.resolved
        return "ppt/presentation.xml"

    # -- content types ------------------------------------------------------
    def content_type(self, name):
        d = self.try_doc("[Content_Types].xml")
        if d is None:
            return ""
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        default = ""
        for el in d.root:
            if local(el.tag) == "Override" and el.get("PartName", "").lstrip("/") == name:
                return el.get("ContentType", "")
            if local(el.tag) == "Default" and el.get("Extension", "").lower() == ext:
                default = el.get("ContentType", "")
        return default

    # -- slides ---------------------------------------------------------------
    def slides(self):
        pres = self.main_part()
        if not self.exists(pres):
            return []
        rmap = {r.id: r for r in self.rels(pres)}
        root = self.doc(pres).root
        lst = root.find("{%s}sldIdLst" % P)
        out = []
        if lst is None:
            return out
        for i, sld in enumerate(list(lst), 1):
            rid = sld.get("{%s}id" % R)
            rel = rmap.get(rid)
            part = rel.resolved if rel else None
            hidden = False
            if part and self.exists(part):
                d = self.try_doc(part)
                hidden = d is not None and d.root.get("show") in ("0", "false")
            out.append({"num": i, "part": part, "hidden": hidden,
                        "sld_id": sld.get("id"), "rid": rid, "el": sld})
        return out

    # -- writing ---------------------------------------------------------------
    def save(self, out_path):
        names = ["[Content_Types].xml"] + [
            n for n in self.order + self.added
            if n != "[Content_Types].xml" and self.exists(n)
        ]
        with zipfile.ZipFile(out_path, "w") as zout:
            for n in names:
                src = self.infos.get(n)
                zi = zipfile.ZipInfo(n, date_time=src.date_time if src else (1980, 1, 1, 0, 0, 0))
                zi.compress_type = src.compress_type if src else zipfile.ZIP_DEFLATED
                if src:
                    zi.external_attr = src.external_attr
                zout.writestr(zi, self.get_bytes(n))


# --------------------------------------------------------------------------
# Locating parts (which slide owns what) and human-readable labels
# --------------------------------------------------------------------------
GENERIC_SHAPE = re.compile(
    r"^(Title|Subtitle|Text ?Box|TextBox|Content Placeholder|Picture|Table|Footer Placeholder|"
    r"Slide Number Placeholder|Date Placeholder|Text Placeholder|Notes Placeholder|"
    r"Slide Image Placeholder|Rectangle|Rounded Rectangle|Oval|Group|Chart|Diagram|Graphic|"
    r"Straight Connector|Straight Arrow Connector|Connector|Shape|Google Shape|Freeform|"
    r"Object|Image|Media|Video|Audio|Header Placeholder|Footer|Vertical Text Placeholder|"
    r"Subtitle Placeholder|Online Media|Ink|SmartArt|Arrow|Line)[\w ;:\-]*\d*$",
    re.I,
)


def safe_label(text):
    """Shape/author names are user-editable and may hold names: mask unless generic."""
    if not text:
        return ""
    return text if GENERIC_SHAPE.match(text) else mask(text)


def build_context(pkg):
    """Map part -> {'kind', 'label', 'slides'}."""
    ctx = {}
    slides = pkg.slides()
    slide_of = {}
    for s in slides:
        if s["part"]:
            slide_of[s["part"]] = s
    owners = {}

    def own(part, num, depth=0):
        if part is None or depth > 3:
            return
        owners.setdefault(part, set()).add(num)
        for r in pkg.rels(part):
            if r.external or r.resolved in slide_of or r.is_type("slideLayout", "slideMaster", "notesMaster"):
                continue
            if r.resolved in owners and num in owners[r.resolved]:
                continue
            own(r.resolved, num, depth + 1)

    for s in slides:
        if s["part"]:
            own(s["part"], s["num"])

    for name in pkg.parts():
        low = name.lower()
        nums = sorted(owners.get(name, []))
        where = ("slide " + ",".join(str(n) for n in nums)) if nums else ""
        if name in slide_of:
            s = slide_of[name]
            kind = "slide"
            label = "slide %d%s" % (s["num"], " (HIDDEN)" if s["hidden"] else "")
        elif "/notesslides/" in low:
            kind, label = "notes", "speaker notes of %s" % (where or name)
        elif "/comments/" in low or "/comments/modern" in low:
            kind, label = "comments", "comments on %s" % (where or name)
        elif low.endswith("commentauthors.xml") or low.endswith("/authors.xml"):
            kind, label = "comment_authors", "comment authors"
        elif "/charts/" in low:
            kind, label = "chart", "chart (%s) on %s" % (posixpath.basename(name), where or "?")
        elif "/diagrams/" in low:
            kind, label = "smartart", "SmartArt (%s) on %s" % (posixpath.basename(name), where or "?")
        elif "/slidelayouts/" in low:
            kind, label = "layout", "slide layout %s" % posixpath.basename(name)
        elif "/slidemasters/" in low:
            kind, label = "master", "slide master %s" % posixpath.basename(name)
        elif "/notesmasters/" in low:
            kind, label = "notes_master", "notes master"
        elif "/handoutmasters/" in low:
            kind, label = "handout_master", "handout master"
        elif low.startswith("docprops/"):
            kind, label = "metadata", "document properties (%s)" % posixpath.basename(name)
        elif low.startswith("customxml/"):
            kind, label = "custom_xml", "custom XML (%s)" % posixpath.basename(name)
        elif "/media/" in low:
            kind, label = "media", "%s%s" % (name, (" on " + where) if where else "")
        elif "/embeddings/" in low:
            kind, label = "embedding", "embedded file %s%s" % (name, (" on " + where) if where else "")
        elif "changesinfo" in low or "revisioninfo" in low:
            kind, label = "revision_history", "revision history (%s)" % posixpath.basename(name)
        elif "/tags/" in low:
            kind, label = "tags", "hidden tag data (%s)" % posixpath.basename(name)
        elif "/fonts/" in low:
            kind, label = "font", "embedded font %s" % posixpath.basename(name)
        elif "vbaproject" in low or "vbadata" in low:
            kind, label = "macro", "VBA macro project %s" % posixpath.basename(name)
        else:
            kind, label = "other", name
        ctx[name] = {"kind": kind, "label": label, "slides": nums}
    return ctx


# --------------------------------------------------------------------------
# Text inventory
# --------------------------------------------------------------------------
SHAPE_LOCALS = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart"}
NAME_ATTR_ELEMENTS = {"cNvPr", "cSld", "section", "custShow", "cmAuthor", "author"}
AUTHOR_ELEMENTS = {"cmAuthor", "author"}


def _find_cnvpr(el):
    for ch in el:
        if local(ch.tag).startswith("nv"):
            for g in ch:
                if local(g.tag) == "cNvPr":
                    return g
    return None


def _walk(el, path, out):
    if local(el.tag) in SHAPE_LOCALS:
        cnv = _find_cnvpr(el)
        if cnv is not None:
            path = path + ["shape #%s %r" % (cnv.get("id", "?"), safe_label(cnv.get("name", "")))]
    if el.tag == A_TBL:
        for ri, tr in enumerate(el.findall(A_TR), 1):
            for ci, tc in enumerate(tr.findall(A_TC), 1):
                for ch in tc:
                    _walk(ch, path + ["table r%dc%d" % (ri, ci)], out)
        return
    if el.tag == A_P:
        text = "".join(t.text or "" for t in el.iter(A_T))
        if text.strip():
            out.append((" · ".join(path), text))
        return
    for ch in el:
        _walk(ch, path, out)


def text_items(pkg, ctx=None):
    """Every piece of text in the package with its location.

    Returns dicts: {part, kind, location, field, text}.
    """
    ctx = ctx or build_context(pkg)
    items = []
    for name in pkg.parts():
        if name == "[Content_Types].xml":
            continue
        info = ctx.get(name, {"kind": "other", "label": name})
        if name.endswith(".rels"):
            src = source_for_rels(name)
            if src is None:
                continue
            for r in pkg.rels(src):
                if r.external:
                    items.append({"part": name, "kind": "hyperlink",
                                  "location": "link target from %s" % ctx.get(src, {}).get("label", src or "package"),
                                  "field": "Target", "text": r.target})
            continue
        if not is_xml_name(name):
            continue
        d = pkg.try_doc(name)
        if d is None:
            raw = pkg.blobs[name].decode("utf-8", "replace")
            raw = re.sub(r"<[^>]+>", " ", raw)
            items.append({"part": name, "kind": info["kind"], "location": info["label"] + " (unparsed)",
                          "field": "raw", "text": raw})
            continue
        paras = []
        _walk(d.root, [], paras)
        for loc, text in paras:
            items.append({"part": name, "kind": info["kind"],
                          "location": info["label"] + (" · " + loc if loc else ""),
                          "field": "paragraph", "text": text})
        top_of = {}
        for top in d.root:
            for el in top.iter():
                top_of[el] = local(top.tag)
        for el in d.root.iter():
            ln = local(el.tag)
            if el.tag != A_T and el.text and el.text.strip():
                items.append({"part": name, "kind": info["kind"],
                              "location": "%s · <%s>" % (info["label"], ln), "field": "element",
                              "top": top_of.get(el, ln), "text": el.text})
            for attr in ("descr", "title"):
                v = el.get(attr)
                if v and v.strip() and ln == "cNvPr":
                    items.append({"part": name, "kind": "alt_text",
                                  "location": "%s · shape #%s alt-text %s" % (info["label"], el.get("id", "?"), attr),
                                  "field": attr, "text": v})
            if ln in NAME_ATTR_ELEMENTS and el.get("name", "").strip() and not GENERIC_SHAPE.match(el.get("name")):
                items.append({"part": name, "kind": info["kind"],
                              "location": "%s · %s name" % (info["label"], ln), "field": "name",
                              "text": el.get("name")})
            if ln in AUTHOR_ELEMENTS:
                for attr in ("initials", "userId"):
                    if el.get(attr, "").strip():
                        items.append({"part": name, "kind": info["kind"],
                                      "location": "%s · author %s" % (info["label"], attr), "field": attr,
                                      "text": el.get(attr)})
            if ln == "tag" and el.get("val", "").strip():
                items.append({"part": name, "kind": info["kind"],
                              "location": "%s · tag value" % info["label"], "field": "val",
                              "text": el.get("val")})
    return items


# --------------------------------------------------------------------------
# Detectors (Safe Harbor identifiers that can be pattern-matched)
# --------------------------------------------------------------------------
_MON = (r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|"
        r"Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)(?![a-z])\.?")
_DAY = r"(?:0?[1-9]|[12]\d|3[01])"
_MO = r"(?:0?[1-9]|1[0-2])"
_YR = r"(?:19|20)\d{2}"
_ORD = r"(?:st|nd|rd|th)?"
_STREET_SUFFIX = (r"(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Lane|Ln|Drive|Dr|Court|Ct|Way|Place|"
                  r"Pl|Terrace|Ter|Circle|Cir|Highway|Hwy|Parkway|Pkwy|Square|Sq|Trail|Trl)")
_STATES = (r"(?:A[LKZR]|C[AOT]|D[EC]|FL|GA|HI|I[ADLN]|K[SY]|LA|M[ADEINOST]|N[CDEHJMVY]|O[HKR]|PA|"
           r"RI|S[CD]|T[NX]|UT|V[AT]|W[AIVY]|PR)")
_NOT_LABEL = r"(?!(?:MRN|DOB|Age|Sex|MR|Acct|ID|SSN|Phone|Tel|Fax|Email|Room|Bed|Date)\b)"
_NAME_TOKEN = r"[A-Z][A-Za-z'\u2019\-]+"


def _valid_numeric_date(m):
    a, b = re.split(r"[/.\-]", m.group(0))[:2]
    a, b = int(a), int(b)
    return 1 <= a <= 31 and 1 <= b <= 31 and (a <= 12 or b <= 12)


# (category, detector, pattern, group, confidence, validator)
_RAW_DETECTORS = [
    ("ssn", "ssn", r"(?<![\d-])\d{3}-\d{2}-\d{4}(?![\d-])", 0, "high", None),
    ("ssn", "ssn_label", r"(?i:\b(?:SSN|social security(?:\s+(?:number|no\.?|#))?))\s*[:#]?\s*(\d{9}|\d{3}[- ]\d{2}[- ]\d{4})\b", 1, "high", None),
    ("email", "email", r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b", 0, "high", None),
    ("url", "url", r"\b(?:https?://|www\.)[^\s<>\"']+", 0, "low", None),
    ("ip_address", "ipv4", r"(?<![\d.])(?:(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)\.){3}(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(?!\d|\.\d)", 0, "medium", None),
    # dates (any element of a date more specific than the year)
    ("date", "date_numeric", r"(?<![\d/.\-])\d{1,2}[/.\-]\d{1,2}[/.\-](?:%s|\d{2})(?![\d/\-]|\.\d)" % _YR, 0, "high", _valid_numeric_date),
    ("date", "date_iso", r"(?<!\d)%s[-/.]%s[-/.]%s(?:[T ]\d{1,2}:\d{2}(?::\d{2})?)?(?!\d)" % (_YR, _MO, _DAY), 0, "high", None),
    ("date", "date_month_day_year", r"\b%s\s+%s%s,?\s+%s\b" % (_MON, _DAY, _ORD, _YR), 0, "high", None),
    ("date", "date_month_day", r"\b%s\s+%s%s(?![\d:])\b" % (_MON, _DAY, _ORD), 0, "medium", None),
    ("date", "date_day_month_year", r"\b%s%s[\s\-]+(?:of\s+)?%s[\s\-,]+(?:%s|\d{2})\b" % (_DAY, _ORD, _MON, _YR), 0, "high", None),
    ("date", "date_day_month", r"\b%s%s[\s\-]+(?:of\s+)?%s" % (_DAY, _ORD, _MON), 0, "medium", None),
    ("date", "date_month_year", r"\b%s,?\s+%s\b" % (_MON, _YR), 0, "low", None),
    ("date", "date_month_day_noyear", r"(?<![\d/.])%s/%s(?![\d/])" % (_MO, _DAY), 0, "low", None),
    # ages over 89
    ("age_over_89", "age_phrase", r"\b(?:9\d|1[01]\d|12[0-5])(?:\s|-)?(?:y/o\b|y\.o\.?|yo[MF]?\b|yrs?\.?(?:\s|-)old|years?(?:\s|-)old|-?year-old|years? of age)", 0, "high", None),
    ("age_over_89", "age_label", r"(?i:\bage[d]?\s*:?\s*)(?:9\d|1[01]\d|12[0-5])\b", 0, "high", None),
    ("age_over_89", "age_sex", r"\b(?:9\d|1[01]\d)\s?(?:M|F)\b", 0, "medium", None),
    # phone / fax
    ("phone_fax", "phone_us", r"(?<![\w\-])(?:\+?1[\s.\-]?)?(?:\(\d{3}\)\s?|\d{3}[\s.\-])\d{3}[\s.\-]\d{4}(?![\w\-])", 0, "high", None),
    ("phone_fax", "phone_intl", r"(?<!\w)\+\d{1,3}[\s.\-]?\d{1,4}[\s.\-]\d{3,4}[\s.\-]\d{3,4}(?!\w)", 0, "high", None),
    ("phone_fax", "phone_7digit", r"(?<![\w\-./])\d{3}-\d{4}(?![\w\-/])", 0, "medium", None),
    # labelled identifiers
    ("mrn", "mrn_label", r"(?i:\b(?:MRN|MR\s?#|med(?:ical)?\s+rec(?:ord)?(?:\s+(?:no\.?|number|#))?|chart\s*(?:#|no\.?|number)|encounter(?:\s+id)?|visit\s+id|CSN|FIN|accession(?:\s+(?:no\.?|number|#))?))\s*[:#]?\s*([A-Za-z]{0,3}\d[A-Za-z0-9\-]{2,})", 1, "high", None),
    ("account", "account_label", r"(?i:\b(?:acct|account)\s*(?:#|no\.?|number)?)\s*[:#]?\s*([A-Za-z]{0,3}\d[A-Za-z0-9\-]{3,})", 1, "high", None),
    ("health_plan", "plan_label", r"(?i:\b(?:member\s*id|subscriber\s*id|policy\s*(?:#|no\.?|number|id)|insurance\s*id|health\s*plan\s*(?:id|#|no\.?)|medicare\s*(?:#|no\.?|id|number)|medicaid\s*(?:#|no\.?|id|number)|MBI|HICN|group\s*(?:#|no\.?)))\s*[:#]?\s*((?=[A-Za-z0-9\-]*\d)[A-Za-z0-9][A-Za-z0-9\-]{3,})", 1, "high", None),
    ("license", "license_label", r"(?i:\b(?:driver'?s?\s+licen[cs]e|licen[cs]e|DL|certificate)\s*(?:#|no\.?|number)?)\s*[:#]\s*((?=[A-Za-z0-9\-]*\d)[A-Za-z0-9][A-Za-z0-9\-]{3,})", 1, "high", None),
    ("device_id", "device_label", r"(?i:\b(?:serial(?:\s+(?:no\.?|number|#))?|S/N|SN|device\s+id|UDI|lot(?:\s+(?:no\.?|#|number))?|VIN|(?:licen[cs]e\s+)?plate(?:\s+(?:no\.?|#|number))?|implant\s+id))\s*[:#]\s*((?=[A-Za-z0-9\-]*\d)[A-Za-z0-9][A-Za-z0-9\-]{3,})", 1, "high", None),
    # addresses and ZIP codes
    ("address", "street", r"\b\d{1,6}\s+(?:[NSEW]\.?\s+)?(?:[A-Z][A-Za-z0-9'\-]*\s+){1,4}%s\b\.?(?:,?\s+(?:Apt|Apartment|Suite|Ste|Unit|#)\.?\s*[A-Za-z0-9\-]+)?" % _STREET_SUFFIX, 0, "high", None),
    ("address", "po_box", r"(?i:\bP\.?\s?O\.?\s+Box\s+\d+)", 0, "high", None),
    ("address", "unit", r"(?i:\b(?:Apt|Apartment|Suite|Unit)\.?\s*#?\s*\d+[A-Za-z]?\b)", 0, "low", None),
    ("zip", "zip_state", r"\b%s\s+(\d{5}(?:-\d{4})?)\b" % _STATES, 1, "high", None),
    ("zip", "zip_label", r"(?i:\bzip(?:\s*code)?)\s*[:#]?\s*(\d{5}(?:-\d{4})?)\b", 1, "high", None),
    ("zip", "zip_plus4", r"(?<![\d\-])\d{5}-\d{4}(?![\d\-])", 0, "medium", None),
    # numbers and codes
    ("long_number", "digits_6plus", r"(?<![\w\-/])(?<!\d[.,])\d{6,}(?![\w\-/]|[.,]\d)", 0, "medium", None),
    ("id_like", "alnum_code", r"\b(?=[A-Z0-9\-]*\d[A-Z0-9\-]*\d[A-Z0-9\-]*\d)(?=[A-Z0-9\-]*[A-Z][A-Z0-9\-]*[A-Z])[A-Z0-9][A-Z0-9\-]{6,}[A-Z0-9]\b", 0, "low", None),
    # names (heuristic only)
    ("name", "patient_label", r"(?i:\b(?:patient(?:\s+name)?|pt(?:\s+name)?|resident|client))\s*[:\-\u2013]\s*(%s%s(?:,?\s+%s(?:%s|[A-Z]\.?)){0,3})" % (_NOT_LABEL, _NAME_TOKEN, _NOT_LABEL, _NAME_TOKEN), 1, "high", None),
    ("name", "name_label", r"(?i:\bname)\s*[:\-\u2013]\s*(%s%s(?:,?\s+%s(?:%s|[A-Z]\.?)){0,3})" % (_NOT_LABEL, _NAME_TOKEN, _NOT_LABEL, _NAME_TOKEN), 1, "medium", None),
    ("name", "honorific", r"\b(?:Mr|Mrs|Ms|Miss|Mx)\.?\s+%s(?:\s+%s)?" % (_NAME_TOKEN, _NAME_TOKEN), 0, "medium", None),
    ("name", "honorific_dr", r"\b(?:Dr|Prof)\.?\s+%s(?:\s+%s)?" % (_NAME_TOKEN, _NAME_TOKEN), 0, "low", None),
]

DETECTORS = [(c, n, re.compile(p), g, conf, v) for (c, n, p, g, conf, v) in _RAW_DETECTORS]

CONTEXT_WORDS = re.compile(
    r"(?i)\b(patient|pt|case|presented|presents|admitted|discharged|hx|son|daughter|wife|husband|"
    r"mother|father|family|named|seen|visit|clinic|bed|room|resident|DOB|MRN)\b")
CAP_PAIR = re.compile(r"\b([A-Z][a-z]+(?:[\-'][A-Z][a-z]+)?)\s+(?:[A-Z]\.\s+)?([A-Z][a-z]+(?:[\-'][A-Z][a-z]+)?)\b")
_STOP = set("""
The This That These Those A An And Or Of In On At For With From To By Case Patient Patients History Presentation
Day Days Week Weeks Month Months Year Years Follow Up Clinical Physical Exam Examination Review Plan Assessment Chief
Complaint Medical Family Social Past Present Illness Discussion Conclusion Conclusions Summary Background Results
Methods Figure Table Slide Learning Objectives Key Points Take Home Emergency Department Hospital University School
Medicine Center Centre Clinic Health System Type Stage Grade Class New North South East West Disease Syndrome Score
Index Test Tests Lab Labs Imaging Diagnosis Differential Treatment Therapy Management Outcome Outcomes Question
Questions Answer References Reference Thank You Acknowledgements Disclosures Introduction Overview Agenda Update
Guidelines Guideline American Association Society College National Institute Journal Trial Study Studies Annual
Meeting Grand Rounds Morning Report Noon Conference Department Division Section Unit Intensive Care Primary Internal
January February March April May June July August September October November December Monday Tuesday Wednesday
Thursday Friday Saturday Sunday Dr Mr Mrs Ms Prof Name DOB MRN Age Sex Male Female Room Bed Visit Seen Admitted
Discharged Presented Presents Family Hx Microsoft Office PowerPoint Google Slides Keynote Content Placeholder
Picture Title Footer Notes Image Text Subtitle Rounds Outside Records Backup Labs Unusual Rash
""".split())

# docProps/app.xml children that carry no personal data (everything else is stripped).
APP_KEEP = {"Application", "AppVersion", "PresentationFormat", "Slides", "Notes", "HiddenSlides",
            "MMClips", "ScaleCrop", "LinksUpToDate", "SharedDoc", "HyperlinksChanged", "DocSecurity"}

# The 18 HIPAA Safe Harbor identifier categories (45 CFR 164.514(b)(2)).
SAFE_HARBOR_18 = [
    "Names",
    "Geographic subdivisions smaller than a state (street, city, county, precinct, ZIP code)",
    "All elements of dates (except year) related to the individual, and all ages over 89",
    "Telephone numbers",
    "Fax numbers",
    "Email addresses",
    "Social Security numbers",
    "Medical record numbers",
    "Health plan beneficiary numbers",
    "Account numbers",
    "Certificate/license numbers",
    "Vehicle identifiers and serial numbers, including license plates",
    "Device identifiers and serial numbers",
    "Web URLs",
    "IP addresses",
    "Biometric identifiers, including finger and voice prints",
    "Full-face photographs and any comparable images",
    "Any other unique identifying number, characteristic, or code (e.g. tattoos, rare conditions)",
]

CONF_RANK = {"high": 3, "medium": 2, "low": 1}

PLACEHOLDERS = {
    "name": "[PATIENT]", "date": "[DATE]", "mrn": "[MRN]", "account": "[ACCOUNT]",
    "health_plan": "[HEALTH PLAN ID]", "age_over_89": "[AGE 90+]", "phone_fax": "[PHONE]",
    "email": "[EMAIL]", "ssn": "[SSN]", "url": "[URL]", "ip_address": "[IP]", "zip": "[ZIP]",
    "address": "[LOCATION]", "location": "[LOCATION]", "long_number": "[ID]", "id_like": "[ID]",
    "device_id": "[DEVICE ID]", "license": "[LICENSE]", "biometric": "[REDACTED]",
    "photo": "[IMAGE REMOVED]", "other": "[REDACTED]",
}


PLACEHOLDER_RE = re.compile(r"\[[A-Z][A-Z0-9 +]*\]")


def detect(text, context_hint=False):
    """Return non-overlapping-ish matches: dicts {category, detector, start, end, value, confidence}."""
    hits = []
    for cat, name, rx, grp, conf, validator in DETECTORS:
        for m in rx.finditer(text):
            if validator and not validator(m):
                continue
            s, e = m.span(grp)
            value = text[s:e]
            if cat == "url":
                stripped = value.rstrip(".,);]'\"")
                e -= len(value) - len(stripped)
                value = stripped
            if not value.strip():
                continue
            hits.append({"category": cat, "detector": name, "start": s, "end": e,
                         "value": value, "confidence": conf})
    if context_hint or CONTEXT_WORDS.search(text):
        for m in CAP_PAIR.finditer(text):
            a, b = m.group(1), m.group(2)
            if a in _STOP or b in _STOP:
                continue
            hits.append({"category": "name", "detector": "capitalized_pair_near_context",
                         "start": m.start(), "end": m.end(), "value": m.group(0), "confidence": "low"})
    # Never flag our own placeholders, e.g. "[AGE 90+]" or "[DATE]".
    ph_spans = [m.span() for m in PLACEHOLDER_RE.finditer(text)]
    hits = [h for h in hits if not any(a < h["end"] and h["start"] < b for a, b in ph_spans)]
    hits.sort(key=lambda h: (h["start"], -(h["end"] - h["start"]), -CONF_RANK[h["confidence"]]))
    kept = []
    for h in hits:
        drop = False
        for k in kept:
            if k["start"] <= h["start"] and h["end"] <= k["end"] and CONF_RANK[k["confidence"]] >= CONF_RANK[h["confidence"]]:
                drop = True
                break
        if not drop:
            kept.append(h)
    return kept


# --------------------------------------------------------------------------
# Masking
# --------------------------------------------------------------------------
def mask(value, category=None, maxlen=48):
    """'Jane Testpatient' -> 'J*** T**********'; '01/02/1950' (date) -> '**/**/1950'."""
    out = []
    for m in re.finditer(r"\d+|[^\W\d_]+|.", value, re.S):
        tok = m.group(0)
        if tok.isdigit():
            if category == "date" and len(tok) == 4 and tok[:2] in ("19", "20"):
                out.append(tok)
            else:
                out.append("*" * len(tok))
        elif tok[0].isalpha():
            out.append(tok[0] + "*" * (len(tok) - 1))
        elif tok in "\r\n\t":
            out.append(" ")
        else:
            out.append(tok)
    s = "".join(out)
    return s if len(s) <= maxlen else s[: maxlen - 3] + "..."


# --------------------------------------------------------------------------
# Images: dimensions, metadata, placeholder PNG, metadata stripping
# --------------------------------------------------------------------------
IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "emf", "wmf", "svg", "webp", "jfif"}
AV_EXTS = {"mp4", "m4v", "mov", "avi", "wmv", "mpg", "mpeg", "mp3", "m4a", "wav", "wma", "aac", "ogg", "webm"}


def image_dims(data):
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack(">II", data[16:24])
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return struct.unpack("<HH", data[6:10])
        if data[:2] == b"BM":
            w, h = struct.unpack("<ii", data[18:26])
            return abs(w), abs(h)
        if data[:2] == b"\xff\xd8":
            i = 2
            while i + 9 < len(data):
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                    i += 2
                    continue
                seglen = struct.unpack(">H", data[i + 2:i + 4])[0]
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    h, w = struct.unpack(">HH", data[i + 5:i + 9])
                    return w, h
                i += 2 + seglen
    except (struct.error, IndexError):
        pass
    return None


def _png_chunks(data):
    i = 8
    while i + 8 <= len(data):
        length = struct.unpack(">I", data[i:i + 4])[0]
        ctype = data[i + 4:i + 8]
        yield ctype, data[i + 8:i + 8 + length], data[i:i + 12 + length]
        i += 12 + length


EXIF_TAGS = {
    0x010E: "ImageDescription", 0x010F: "Make", 0x0110: "Model", 0x0131: "Software",
    0x0132: "DateTime", 0x013B: "Artist", 0x8298: "Copyright", 0x9003: "DateTimeOriginal",
    0x9004: "DateTimeDigitized", 0x9286: "UserComment", 0xA430: "CameraOwnerName",
    0xA431: "BodySerialNumber", 0xA435: "LensSerialNumber", 0x9C9B: "XPTitle",
    0x9C9C: "XPComment", 0x9C9D: "XPAuthor", 0x9C9E: "XPKeywords", 0x9C9F: "XPSubject",
}
EXIF_GROUPS = {
    "camera": {"Make", "Model", "Software", "BodySerialNumber", "LensSerialNumber"},
    "date": {"DateTime", "DateTimeOriginal", "DateTimeDigitized"},
    "owner": {"Artist", "Copyright", "CameraOwnerName", "XPAuthor"},
    "description": {"ImageDescription", "UserComment", "XPTitle", "XPComment", "XPKeywords", "XPSubject"},
}
GPS_TAGS = {0: "GPSVersionID", 1: "GPSLatitudeRef", 2: "GPSLatitude", 3: "GPSLongitudeRef",
            4: "GPSLongitude", 5: "GPSAltitudeRef", 6: "GPSAltitude", 7: "GPSTimeStamp",
            18: "GPSMapDatum", 29: "GPSDateStamp"}
_TYPE_SIZE = {1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8}


def parse_exif(tiff):
    """Parse a TIFF/EXIF block -> {'fields': {name: text}, 'gps_tags': [names]}. Never raises."""
    out = {"fields": {}, "gps_tags": []}
    if tiff[:2] == b"II":
        end = "<"
    elif tiff[:2] == b"MM":
        end = ">"
    else:
        return out

    def u16(o):
        return struct.unpack(end + "H", tiff[o:o + 2])[0]

    def u32(o):
        return struct.unpack(end + "I", tiff[o:o + 4])[0]

    seen = set()

    def read_ifd(off, depth, is_gps):
        if off in seen or depth > 3 or off <= 0 or off + 2 > len(tiff):
            return
        seen.add(off)
        for k in range(min(u16(off), 512)):
            e = off + 2 + 12 * k
            if e + 12 > len(tiff):
                break
            tag, typ, cnt = u16(e), u16(e + 2), u32(e + 4)
            size = _TYPE_SIZE.get(typ, 1) * cnt
            voff = e + 8 if size <= 4 else u32(e + 8)
            raw = tiff[voff:voff + size]
            if is_gps:
                out["gps_tags"].append(GPS_TAGS.get(tag, "GPSTag%d" % tag))
            elif tag == 0x8769:
                read_ifd(u32(e + 8), depth + 1, False)
            elif tag == 0x8825:
                read_ifd(u32(e + 8), depth + 1, True)
            elif tag in EXIF_TAGS:
                name = EXIF_TAGS[tag]
                if name.startswith("XP"):
                    val = raw.decode("utf-16-le", "replace")
                elif name == "UserComment":
                    val = raw[8:].decode("latin-1", "replace")
                else:
                    val = raw.decode("latin-1", "replace")
                val = val.replace("\x00", "").strip()
                if val:
                    out["fields"][name] = val

    try:
        read_ifd(u32(4), 0, False)
    except (struct.error, IndexError):
        pass
    return out


def image_metadata(data):
    """Inspect an image's embedded metadata (PNG and JPEG).

    Returns {'blocks', 'exif_present', 'gps', 'gps_tags', 'flags',
             'exif_fields' (MAY HOLD PHI), 'texts' (MAY HOLD PHI)}.
    """
    info = {"blocks": [], "exif_present": False, "gps": False, "gps_tags": [],
            "exif_fields": {}, "flags": [], "texts": []}

    def add_exif(tiff):
        info["exif_present"] = True
        ex = parse_exif(tiff)
        info["gps_tags"].extend(ex["gps_tags"])
        info["exif_fields"].update(ex["fields"])
        info["texts"].extend(ex["fields"].values())

    if data[:8] == b"\x89PNG\r\n\x1a\n":
        for ctype, body, _raw in _png_chunks(data):
            if ctype in (b"tEXt", b"zTXt", b"iTXt"):
                try:
                    if ctype == b"tEXt":
                        k, txt = body.split(b"\x00", 1)
                        txt = txt.decode("latin-1")
                    elif ctype == b"zTXt":
                        k, rest = body.split(b"\x00", 1)
                        txt = zlib.decompress(rest[1:]).decode("latin-1")
                    else:
                        k, rest = body.split(b"\x00", 1)
                        comp, rest = rest[0], rest[2:]
                        _lang, rest = rest.split(b"\x00", 1)
                        _tk, rest = rest.split(b"\x00", 1)
                        txt = (zlib.decompress(rest) if comp else rest).decode("utf-8", "replace")
                    key = k.decode("latin-1")
                    xmp = key == "XML:com.adobe.xmp"
                    info["blocks"].append("XMP" if xmp else "PNG text")
                    info["texts"].append(re.sub(r"<[^>]+>", " ", txt) if xmp else txt)
                except (ValueError, zlib.error, IndexError):
                    info["blocks"].append("PNG text (unreadable)")
            elif ctype == b"eXIf":
                info["blocks"].append("EXIF")
                add_exif(body)
    elif data[:2] == b"\xff\xd8":
        i = 2
        while i + 4 <= len(data) and data[i] == 0xFF:
            marker = data[i + 1]
            if marker == 0xDA:
                break
            seglen = struct.unpack(">H", data[i + 2:i + 4])[0]
            body = data[i + 4:i + 2 + seglen]
            if marker == 0xE1 and body.startswith(b"Exif\x00\x00"):
                info["blocks"].append("EXIF")
                add_exif(body[6:])
            elif marker == 0xE1:
                info["blocks"].append("XMP" if b"ns.adobe.com/xap" in body else "APP1")
                info["texts"].append(re.sub(r"<[^>]+>", " ", body.decode("utf-8", "replace")))
            elif marker == 0xED:
                info["blocks"].append("IPTC (APP13)")
                info["texts"].extend(s.decode("latin-1") for s in re.findall(rb"[\x20-\x7e]{4,}", body))
            elif marker == 0xFE:
                info["blocks"].append("JPEG comment")
                info["texts"].append(body.decode("latin-1", "replace"))
            elif 0xE2 <= marker <= 0xEF and marker != 0xEE and not body.startswith(b"ICC_PROFILE"):
                info["blocks"].append("APP%d" % (marker - 0xE0))
            i += 2 + seglen
    info["gps"] = bool(info["gps_tags"])
    info["blocks"] = sorted(set(info["blocks"]))
    flags = []
    if info["gps"]:
        pos = {"GPSLatitude", "GPSLongitude"} & set(info["gps_tags"])
        flags.append("GPS location found" if pos else "GPS tags found")
    for group, names in EXIF_GROUPS.items():
        hit = sorted(n for n in names if n in info["exif_fields"])
        if hit:
            flags.append("%s: %s" % (group, ", ".join(hit)))
    other = [b for b in info["blocks"] if b != "EXIF"]
    if other:
        flags.append("other metadata: " + ", ".join(other))
    info["flags"] = flags
    return info


def strip_image_metadata(data):
    """Drop text/EXIF metadata from PNG and JPEG. Other formats are returned unchanged."""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        out = [data[:8]]
        for ctype, _body, raw in _png_chunks(data):
            if ctype in (b"tEXt", b"zTXt", b"iTXt", b"eXIf", b"tIME"):
                continue
            out.append(raw)
        return b"".join(out)
    if data[:2] == b"\xff\xd8":
        out = [data[:2]]
        i = 2
        while i + 4 <= len(data) and data[i] == 0xFF:
            marker = data[i + 1]
            if marker == 0xDA:
                break
            seglen = struct.unpack(">H", data[i + 2:i + 4])[0]
            seg = data[i:i + 2 + seglen]
            body = seg[4:]
            drop = marker == 0xFE or marker == 0xE1 or (
                0xE3 <= marker <= 0xEF and marker != 0xEE) or (marker == 0xE2 and not body.startswith(b"ICC_PROFILE"))
            if not drop:
                out.append(seg)
            i += 2 + seglen
        out.append(data[i:])
        return b"".join(out)
    return data


def grey_png(width, height, grey=0xC0, cap=4096):
    width = max(1, int(width or 320))
    height = max(1, int(height or 240))
    scale = min(1.0, float(cap) / max(width, height))
    width, height = max(1, int(width * scale)), max(1, int(height * scale))
    row = b"\x00" + bytes([grey]) * width
    raw = row * height

    def chunk(tag, body):
        return struct.pack(">I", len(body)) + tag + body + struct.pack(">I", zlib.crc32(tag + body) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 0, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


# --------------------------------------------------------------------------
# Package validity checks (used by scrub's self-check and the tests)
# --------------------------------------------------------------------------
def check_package(path):
    """Return a list of structural problems (empty list = looks consistent)."""
    problems = []
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        return ["not a zip: %s" % exc]
    with z:
        bad = z.testzip()
        if bad:
            problems.append("corrupt zip entry: %s" % bad)
        names = [i.filename for i in z.infolist() if not i.is_dir()]
        if not names or names[0] != "[Content_Types].xml":
            problems.append("[Content_Types].xml is not the first zip entry")
        nameset = set(names)
        docs = {}
        for n in names:
            if is_xml_name(n):
                try:
                    docs[n] = ET.fromstring(z.read(n))
                except ET.ParseError as exc:
                    if n.lower().endswith(".vml"):
                        continue
                    problems.append("XML does not parse: %s (%s)" % (n, exc))
        ct = docs.get("[Content_Types].xml")
        defaults, overrides = {}, {}
        if ct is not None:
            for el in ct:
                if local(el.tag) == "Default":
                    defaults[el.get("Extension", "").lower()] = el.get("ContentType")
                elif local(el.tag) == "Override":
                    overrides[el.get("PartName", "").lstrip("/")] = el.get("ContentType")
        for pn in overrides:
            if pn not in nameset:
                problems.append("content-type override for missing part: %s" % pn)
        for n in names:
            if n == "[Content_Types].xml":
                continue
            ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
            if n not in overrides and ext not in defaults:
                problems.append("no content type for part: %s" % n)
        rel_ids = {}
        for n in names:
            if not n.endswith(".rels") or n not in docs:
                continue
            src = source_for_rels(n)
            if src is None:
                continue
            if src != "" and src not in nameset:
                problems.append("rels file for missing part: %s" % n)
                continue
            ids = set()
            for el in docs[n]:
                if local(el.tag) != "Relationship":
                    continue
                ids.add(el.get("Id"))
                if el.get("TargetMode") == "External":
                    continue
                tgt = resolve_target(src, el.get("Target", ""))
                if tgt and tgt not in nameset:
                    problems.append("relationship %s in %s points to missing part %s" % (el.get("Id"), n, tgt))
            rel_ids[src] = ids
        rattr = "{%s}" % R
        for n, root in docs.items():
            if n.endswith(".rels") or n == "[Content_Types].xml":
                continue
            ids = rel_ids.get(n, set())
            for el in root.iter():
                for k, v in el.attrib.items():
                    if k.startswith(rattr) and v and v not in ids:
                        problems.append("%s: <%s %s=%r> has no matching relationship" % (n, local(el.tag), local(k), v))
    return problems
