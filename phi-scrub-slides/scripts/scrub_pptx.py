#!/usr/bin/env python3
"""Apply an approved, fully decided de-identification plan to a .pptx.
Python standard library only. Fail-closed.

    python3 scrub_pptx.py deck.pptx --plan deck.plan.json --out deck.deidentified.pptx

- The input file is never modified.
- Refuses to run while any finding or image in the plan is UNRESOLVED, or the plan
  was made for a different file, unless --force (which is written loudly into the report).
- Refuses macro-enabled decks (.pptm) unless --strip-macros.
- After writing, re-scans the output and searches every remaining XML part and all
  decoded text for every original finding's value. Any hit = no output file and a
  nonzero exit (unless --force).
- Writes <out>.scrub-report.md (masked values only) with residual risks and a human
  attestation checklist built on the 18 HIPAA Safe Harbor categories.
"""

import argparse
import datetime
import hashlib
import json
import os
import posixpath
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from phi_common import (  # noqa: E402
    A_P, A_T, APP_KEEP, P, PLACEHOLDERS, R, SAFE_HARBOR_18, VERSION, Package, grey_png, image_dims,
    is_xml_name, local, mask, relative_target, rels_path_for, source_for_rels, strip_image_metadata,
)
import scan_pptx  # noqa: E402
from scan_pptx import AV_DECISIONS, DEFAULT_ACTIONS, IMAGE_DECISIONS, covered_by  # noqa: E402

REMOVABLE_REF_ELEMENTS = {"hlinkClick", "hlinkHover", "hlinkMouseOver", "sldId", "sld", "externalData",
                          "tags", "custData"}
NAME_ATTR_ELEMENTS = {"cNvPr", "cSld", "section", "custShow", "cmAuthor", "author"}
MACRO_MAIN = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
PPTX_MAIN = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"


class Replacer:
    def __init__(self, reps):
        self.reps = []
        for r in reps:
            flags = 0 if r.get("match_case", True) else re.I
            self.reps.append((r["find"], re.compile(re.escape(r["find"]), flags), r["replace"]))
        self.counts = {r[0]: 0 for r in self.reps}

    def text(self, s):
        if not s:
            return s
        for find, rx, repl in self.reps:
            s, n = rx.subn(lambda _m, repl=repl: repl, s)
            self.counts[find] += n
        return s

    def paragraph(self, ts):
        """Replace across <a:t> runs: match the joined paragraph text, write the
        placeholder into the first affected run, empty the matched part of the others."""
        chars = [[ch, i] for i, t in enumerate(ts) for ch in (t.text or "")]
        s = "".join(c for c, _ in chars)
        changed = False
        for find, rx, repl in self.reps:
            pos = 0
            while True:
                m = rx.search(s, pos)
                if not m:
                    break
                a, b = m.span()
                if a == b:
                    pos = a + 1
                    continue
                owner = chars[a][1]
                chars[a:b] = [[ch, owner] for ch in repl]
                s = s[:a] + repl + s[b:]
                pos = a + len(repl)
                self.counts[find] += 1
                changed = True
        if changed:
            texts = [""] * len(ts)
            for ch, i in chars:
                texts[i] += ch
            for t, tx in zip(ts, texts):
                t.text = tx
        return changed


# ---------------------------------------------------------------------------
def check_plan(plan, report, strip_macros):
    """Validate a plan against a fresh scan of the input. Returns (reps, actions, images, errors, kept)."""
    errors = []
    actions = dict(DEFAULT_ACTIONS)
    actions.update(plan.get("actions", {}))
    if plan.get("source_sha256") and plan["source_sha256"] != report["sha256"]:
        errors.append("plan was made for a different file (sha256 mismatch); re-run the scan")
    hidden = report["structure"]["hidden_slides"]

    reps = {}
    for r in plan.get("replacements", []):
        f = r.get("find", "")
        if not isinstance(f, str) or len(f.strip()) < 3:
            errors.append("replacement 'find' must be at least 3 characters: %r" % mask(str(f)))
            continue
        reps[f] = {"find": f, "replace": str(r.get("replace", "[REDACTED]")),
                   "match_case": bool(r.get("match_case", True)), "category": r.get("category")}

    decided = {d.get("id"): d for d in plan.get("findings", [])}
    kept = []
    for f in report["findings"]:
        d = decided.get(f["id"])
        where = "%s %s (%s) at %s: %s" % (f["id"], f["category"], f["confidence"], f["location"], f["masked"])
        if d is None or d.get("value") not in (None, f["value"]):
            errors.append("no decision for finding " + where)
            continue
        dec = str(d.get("decision", "UNRESOLVED"))
        if dec == "replace":
            if not f["replaceable"]:
                errors.append("finding %s is inside image metadata or an embedded file; it cannot be replaced as "
                              "text. Use strip_image_metadata / embedded_objects / an image decision." % f["id"])
            elif f["value"] not in reps:
                reps[f["value"]] = {"find": f["value"], "replace": PLACEHOLDERS.get(f["category"], "[REDACTED]"),
                                    "match_case": f["category"] != "name", "category": f["category"]}
        elif dec == "keep":
            if not str(d.get("reason", "")).strip():
                errors.append("finding %s is marked keep without a reason" % f["id"])
            kept.append(f["value"])
        elif dec.startswith("removed_by:"):
            cov = covered_by(f, actions, hidden)
            if cov is None:
                errors.append("finding %s says %s but no enabled action removes %s" % (f["id"], dec, f["location"]))
        else:
            errors.append("UNRESOLVED finding " + where)

    images = {}
    planned = plan.get("images", {})
    for im in report["images"]:
        spec = planned.get(im["part"])
        if spec is None:
            for k, v in planned.items():
                if posixpath.basename(k) == posixpath.basename(im["part"]):
                    spec = v
        if isinstance(spec, str):
            spec = {"decision": spec}
        spec = dict(spec or {})
        dec = spec.get("decision", spec.get("action", "UNRESOLVED"))
        if dec == "placeholder":
            dec = "replace"
        allowed = AV_DECISIONS if im["audio_video"] else IMAGE_DECISIONS
        if dec not in allowed:
            errors.append("UNRESOLVED %s %s (slides %s): decide %s" % (
                "audio/video" if im["audio_video"] else "image", im["part"], im["slides"] or "-",
                " / ".join(sorted(allowed - {"placeholder"}))))
        spec["decision"] = dec
        images[im["part"]] = spec

    st = report["structure"]
    if (st["macros"]["parts"] or st["macros"]["macro_enabled"]) and not strip_macros:
        errors.append("deck contains VBA macros (.pptm). Refusing. Re-run with --strip-macros to remove them "
                      "and save as a macro-free .pptx")
    ordered = sorted(reps.values(), key=lambda r: -len(r["find"]))
    return ordered, actions, images, errors, kept


# ---------------------------------------------------------------------------
class Scrubber:
    def __init__(self, pkg, reps, actions, images, strip_macros):
        self.pkg = pkg
        self.rep = Replacer(reps)
        self.a = actions
        self.images = images
        self.strip_macros = strip_macros
        self.log = []
        self.removed_rids = {}
        self.by_hand = []
        self.warnings = []
        self.unresolved = []

    # -- helpers ------------------------------------------------------------------
    def drop_rel(self, rel):
        self.pkg.remove_rel(rel)
        self.removed_rids.setdefault(rel.source, set()).add(rel.id)

    def delete(self, part, why):
        if part and self.pkg.exists(part):
            self.pkg.delete(part)
            self.log.append("removed %s (%s)" % (part, why))

    def rels_of_type(self, *types):
        return [r for r in self.pkg.all_rels() if not r.external and r.is_type(*types)]

    @staticmethod
    def _refs_rid(el, rid):
        return any(v == rid for k, v in el.attrib.items() if k.startswith("{%s}" % R))

    # -- structural actions ---------------------------------------------------------
    def hidden_slides(self):
        hidden = [s for s in self.pkg.slides() if s["hidden"]]
        if self.a["hidden_slides"] != "remove":
            if hidden:
                self.warnings.append("hidden slides KEPT by plan: %s" % [s["num"] for s in hidden])
            return
        gone = set()
        for s in hidden:
            gone.add(s["sld_id"])
            self.delete(s["part"], "hidden slide %d" % s["num"])
        if gone:
            pres = self.pkg.main_part()
            d = self.pkg.doc(pres)
            for lst in d.root.iter():
                if local(lst.tag) == "sldIdLst" and lst.tag != "{%s}sldIdLst" % P:
                    for el in list(lst):
                        if el.get("id") in gone:
                            lst.remove(el)
            self.pkg.mark(pres)

    def comments(self):
        if self.a["remove_comments"]:
            for rel in self.rels_of_type("comments", "commentAuthors", "authors"):
                self.delete(rel.resolved, "comments / comment authors")

    def notes(self):
        if self.a["speaker_notes"] == "delete":
            for rel in self.rels_of_type("notesSlide"):
                self.delete(rel.resolved, "speaker notes")

    def thumbnail(self):
        if self.a["remove_thumbnail"]:
            for rel in self.pkg.rels(""):
                if rel.is_type("thumbnail"):
                    self.delete(rel.resolved, "thumbnail picture of slide 1")

    def custom_parts(self):
        if self.a["strip_metadata"]:
            for rel in self.pkg.rels(""):
                if rel.is_type("custom-properties"):
                    self.delete(rel.resolved, "custom document properties")
            for rel in self.rels_of_type("changesInfo", "revisionInfo"):
                self.delete(rel.resolved, "revision history")
        if self.a["remove_custom_xml"]:
            for rel in self.rels_of_type("customXml", "customData"):
                self.drop_rel(rel)
                self.delete(rel.resolved, "custom XML data")
        if self.a["remove_tags"]:
            for rel in self.rels_of_type("tags"):
                self.drop_rel(rel)
                self.delete(rel.resolved, "hidden tag data")

    def macros(self):
        vba = self.rels_of_type("vbaProject")
        main = self.pkg.main_part()
        ct = self.pkg.doc("[Content_Types].xml")
        macro_ct = [el for el in ct.root if local(el.tag) == "Override" and el.get("ContentType") == MACRO_MAIN]
        if not vba and not macro_ct:
            return
        if not self.strip_macros:
            raise SystemExit("macro-enabled deck: refusing (use --strip-macros)")
        for rel in vba:
            self.drop_rel(rel)
            self.delete(rel.resolved, "VBA macro project")
        for n in self.pkg.parts():
            if "vbadata" in n.lower():
                self.delete(n, "VBA macro data")
        for el in macro_ct:
            el.set("ContentType", PPTX_MAIN)
        for el in ct.root:
            if local(el.tag) == "Default" and el.get("Extension", "").lower() == "bin" and \
                    "vbaProject" in el.get("ContentType", ""):
                ct.root.remove(el)
                break
        self.pkg.mark("[Content_Types].xml")
        self.warnings.append("MACROS STRIPPED: the VBA project was removed and the deck saved as a macro-free "
                             ".pptx. Anything the macros did will no longer work. (%s)" % main)

    def embedded_objects(self):
        emb = [n for n in self.pkg.parts() if "/embeddings/" in n.lower()]
        if not emb:
            return
        if self.a["embedded_objects"] != "remove":
            for n in emb:
                self.warnings.append("embedded object KEPT by plan, not scanned safely: %s" % n)
            return
        for rel in self.pkg.all_rels():
            if rel.external or rel.resolved not in emb:
                continue
            src = rel.source
            if "/charts/" in src:
                self.drop_rel(rel)  # fix_dangling removes <c:externalData>; the chart keeps its cached values
                self.log.append("removed workbook behind %s (chart picture and cached values kept)" % src)
                continue
            d = self.pkg.try_doc(src) if is_xml_name(src) else None
            if d is None:
                self.unresolved.append("embedded object %s referenced from unparsable %s" % (rel.resolved, src))
                continue
            parents = {c: p for p in d.root.iter() for c in p}
            gone_rids = set()
            for el in [e for e in d.root.iter() if self._refs_rid(e, rel.id)]:
                if el is not d.root and el not in parents:
                    continue  # already removed together with an earlier mc:AlternateContent branch
                frame = el
                while frame is not None and local(frame.tag) != "graphicFrame":
                    frame = parents.get(frame)
                if frame is None:
                    self.unresolved.append("embedded object %s is used by <%s> in %s, which this tool cannot "
                                           "remove safely" % (rel.resolved, local(el.tag), src))
                    continue
                top = frame
                while parents.get(top) is not None and local(parents[top].tag) in ("Choice", "Fallback"):
                    top = parents[parents[top]]  # the enclosing mc:AlternateContent
                if top not in parents:
                    continue
                for sub in top.iter():
                    gone_rids.update(v for k, v in sub.attrib.items() if k.startswith("{%s}" % R))
                parents[top].remove(top)
                parents = {c: p for p in d.root.iter() for c in p}
                self.pkg.mark(src)
            still = {v for e in d.root.iter() for k, v in e.attrib.items() if k.startswith("{%s}" % R)}
            for r2 in self.pkg.rels(src):
                if r2.id in gone_rids and r2.id not in still:
                    self.drop_rel(r2)
            self.log.append("removed embedded object shape(s) on %s" % src)
        for n in emb:
            self.delete(n, "embedded object (cannot be scanned safely)")

    def hyperlinks(self):
        mode = self.a["hyperlinks"]
        if mode == "keep":
            return
        finds = [f.lower() for f, _rx, _r in self.rep.reps]
        for rel in list(self.pkg.all_rels()):
            if not (rel.external and rel.is_type("hyperlink")):
                continue
            if mode == "remove_all" or any(f in rel.target.lower() for f in finds):
                self.drop_rel(rel)
                self.log.append("removed hyperlink %s" % mask(rel.target, "url"))

    # -- images ---------------------------------------------------------------------
    def images_pass(self):
        for n, spec in self.images.items():
            if not self.pkg.exists(n):
                continue
            dec = spec["decision"]
            if spec.get("by_hand"):
                self.by_hand.append("%s (slides %s): %s" % (n, spec.get("slides", "-"), spec["by_hand"]))
            if dec == "replace":
                self.placeholder(n)
            elif dec == "remove":
                self.remove_image(n)
            elif dec == "mask":
                self.mask_image(n, spec.get("boxes", []))
            if self.pkg.exists(n) and self.a["strip_image_metadata"]:
                data = self.pkg.blobs[n]
                new = strip_image_metadata(data)
                if new != data:
                    self.pkg.set_blob(n, new)
                    self.log.append("stripped hidden metadata (EXIF/GPS/XMP/IPTC/text) from %s" % n)

    def _refs(self, part):
        return [r for r in self.pkg.all_rels() if not r.external and r.resolved == part]

    def placeholder(self, part):
        dims = image_dims(self.pkg.blobs[part]) or (320, 240)
        stem = posixpath.splitext(part)[0]
        new, i = stem + "_redacted.png", 1
        while self.pkg.exists(new):
            i += 1
            new = "%s_redacted%d.png" % (stem, i)
        self.pkg.set_blob(new, grey_png(*dims))
        for rel in self._refs(part):
            rel.el.set("Target", relative_target(rel.source, new))
            self.pkg.mark(rels_path_for(rel.source))
        self.ensure_default("png", "image/png")
        self.delete(part, "image replaced by grey placeholder %s" % posixpath.basename(new))

    def remove_image(self, part):
        leftover = False
        for rel in self._refs(part):
            d = self.pkg.try_doc(rel.source)
            if d is None:
                leftover = True
                continue
            parents = {c: p for p in d.root.iter() for c in p}
            for pic in list(d.root.iter("{%s}pic" % P)):
                if any(self._refs_rid(el, rel.id) for el in pic.iter()) and pic in parents:
                    parents[pic].remove(pic)
                    self.pkg.mark(rel.source)
            if any(self._refs_rid(el, rel.id) for el in d.root.iter()):
                leftover = True
        if leftover:
            self.warnings.append("%s is also used outside a picture shape (e.g. a background); replaced with a "
                                 "grey placeholder instead of removed" % part)
            self.placeholder(part)
        else:
            self.delete(part, "picture removed from slides")

    def mask_image(self, part, boxes):
        try:
            from PIL import Image, ImageDraw  # optional extra, never required
        except ImportError:
            raise SystemExit("image %s: 'mask' needs Pillow, which is not installed. Use 'replace' (grey "
                             "placeholder), or crop/cover the region by hand in PowerPoint/Slides." % part)
        import io
        im = Image.open(io.BytesIO(self.pkg.blobs[part]))
        fmt = im.format or "PNG"
        draw = ImageDraw.Draw(im)
        for box in boxes:
            draw.rectangle([int(v) for v in box], fill=0 if im.mode in ("L", "P", "1") else (0, 0, 0))
        buf = io.BytesIO()
        im.save(buf, format=fmt)
        self.pkg.set_blob(part, buf.getvalue())
        self.log.append("masked %d region(s) in %s (Pillow)" % (len(boxes), part))

    def ensure_default(self, ext, ctype):
        d = self.pkg.doc("[Content_Types].xml")
        for el in d.root:
            if local(el.tag) == "Default" and el.get("Extension", "").lower() == ext:
                return
        import xml.etree.ElementTree as ET
        d.root.insert(0, ET.Element("{http://schemas.openxmlformats.org/package/2006/content-types}Default",
                                    {"Extension": ext, "ContentType": ctype}))
        self.pkg.mark("[Content_Types].xml")

    # -- text ---------------------------------------------------------------------------
    def text_pass(self):
        clear_all = self.a["alt_text"] == "clear_all"
        clear_matched = self.a["alt_text"] == "clear_matched"
        for name in self.pkg.parts():
            if name == "[Content_Types].xml":
                continue
            if name.endswith(".rels"):
                src = source_for_rels(name)
                if src is None:
                    continue
                for rel in self.pkg.rels(src):
                    if rel.external:
                        new = self.rep.text(rel.target)
                        if new != rel.target:
                            rel.el.set("Target", new)
                            self.pkg.mark(name)
                continue
            if not is_xml_name(name):
                continue
            d = self.pkg.try_doc(name)
            if d is None:
                raw = self.pkg.blobs[name].decode("utf-8", "replace")
                new = self.rep.text(raw)
                if new != raw:
                    self.pkg.set_blob(name, new.encode("utf-8"))
                    self.warnings.append("%s could not be parsed; replaced as plain text" % name)
                continue
            changed = False
            for p in d.root.iter(A_P):
                ts = list(p.iter(A_T))
                if ts and self.rep.paragraph(ts):
                    changed = True
            for el in d.root.iter():
                ln = local(el.tag)
                if el.tag != A_T and el.text and el.text.strip():
                    new = self.rep.text(el.text)
                    if new != el.text:
                        el.text = new
                        changed = True
                attrs = []
                if ln == "cNvPr":
                    attrs += ["descr", "title"]
                if ln in NAME_ATTR_ELEMENTS:
                    attrs.append("name")
                if ln in ("cmAuthor", "author"):
                    attrs += ["initials", "userId"]
                if ln == "tag":
                    attrs.append("val")
                for at in attrs:
                    v = el.get(at)
                    if not v:
                        continue
                    new = self.rep.text(v)
                    if at in ("descr", "title") and (clear_all or (clear_matched and new != v)):
                        del el.attrib[at]
                        changed = True
                    elif new != v:
                        el.set(at, new)
                        changed = True
            if changed:
                self.pkg.mark(name)

    def metadata(self):
        if not self.a["strip_metadata"]:
            return
        if self.pkg.exists("docProps/core.xml"):
            d = self.pkg.doc("docProps/core.xml")
            for el in list(d.root):
                d.root.remove(el)
            self.pkg.mark("docProps/core.xml")
            self.log.append("cleared document properties (author, last modified by, title, subject, keywords, dates)")
        if self.pkg.exists("docProps/app.xml"):
            d = self.pkg.doc("docProps/app.xml")
            for el in list(d.root):
                if local(el.tag) not in APP_KEEP:
                    d.root.remove(el)
            self.pkg.mark("docProps/app.xml")
            self.log.append("cleared app properties (company, manager, template, slide-title list)")

    # -- consistency -----------------------------------------------------------------------
    def garbage_collect(self):
        while True:
            changed = False
            for rel in list(self.pkg.all_rels()):
                if not rel.external and rel.resolved and not self.pkg.exists(rel.resolved):
                    self.drop_rel(rel)
                    changed = True
            reach, todo = set(), [""]
            while todo:
                cur = todo.pop()
                for rel in self.pkg.rels(cur):
                    if not rel.external and rel.resolved not in reach and self.pkg.exists(rel.resolved):
                        reach.add(rel.resolved)
                        todo.append(rel.resolved)
            for n in self.pkg.parts():
                if n == "[Content_Types].xml":
                    continue
                if n.endswith(".rels"):
                    src = source_for_rels(n)
                    if src not in ("", None) and not self.pkg.exists(src):
                        self.pkg.delete(n)
                        changed = True
                    continue
                if n not in reach:
                    self.delete(n, "no longer referenced")
                    changed = True
            if not changed:
                break

    def fix_dangling(self):
        rattr = "{%s}" % R
        for src, rids in self.removed_rids.items():
            if not src or not self.pkg.exists(src) or not is_xml_name(src):
                continue
            d = self.pkg.try_doc(src)
            if d is None:
                continue
            live = {r.id for r in self.pkg.rels(src)}
            parents = {c: p for p in d.root.iter() for c in p}
            for el in list(d.root.iter()):
                bad = [k for k, v in el.attrib.items() if k.startswith(rattr) and v in rids and v not in live]
                if not bad:
                    continue
                if local(el.tag) in REMOVABLE_REF_ELEMENTS and el in parents:
                    parents[el].remove(el)
                else:
                    self.unresolved.append("%s: <%s> still points at a removed part" % (src, local(el.tag)))
                self.pkg.mark(src)

    def content_types(self):
        d = self.pkg.doc("[Content_Types].xml")
        for el in list(d.root):
            if local(el.tag) == "Override" and not self.pkg.exists(el.get("PartName", "").lstrip("/")):
                d.root.remove(el)
                self.pkg.mark("[Content_Types].xml")

    def app_counts(self):
        if not self.pkg.exists("docProps/app.xml"):
            return
        slides = self.pkg.slides()
        notes = {r.resolved for r in self.pkg.all_rels() if r.is_type("notesSlide") and self.pkg.exists(r.resolved or "")}
        values = {"Slides": len(slides), "HiddenSlides": sum(1 for s in slides if s["hidden"]), "Notes": len(notes)}
        d = self.pkg.doc("docProps/app.xml")
        for el in d.root:
            if local(el.tag) in values and el.text != str(values[local(el.tag)]):
                el.text = str(values[local(el.tag)])
                self.pkg.mark("docProps/app.xml")

    def run(self):
        self.macros()
        self.hidden_slides()
        self.comments()
        self.notes()
        self.thumbnail()
        self.custom_parts()
        self.embedded_objects()
        self.hyperlinks()
        self.images_pass()
        self.garbage_collect()
        self.text_pass()
        self.metadata()
        self.garbage_collect()
        self.fix_dangling()
        self.content_types()
        self.app_counts()


# ---------------------------------------------------------------------------
RESIDUAL_RISK = [
    "What is *inside* pictures: faces, eyes, tattoos, birthmarks, scars, jewelry, room or bed numbers, signs, "
    "logos, views from windows, wristbands, badges, screens, burned-in dates, DICOM/EHR overlays. Pixels are "
    "never read by this tool.",
    "Audio and video: frames and soundtrack are never inspected.",
    "Names, places and context that do not follow a pattern: a rare diagnosis in a small town, an unusual "
    "occupation, a famous event, a date next to a rare event, small cell counts in tables (n < 11).",
    "Text drawn as shapes, WordArt converted to curves, text inside EMF/WMF/SVG artwork that was not parsed, "
    "and anything in parts listed as 'not scanned' in the scan.",
    "Charts: cached labels and values are scanned and replaced, but a chart picture can still show identifying "
    "outliers.",
    "Google Slides: the ORIGINAL Slides file keeps its version history, comments, suggestions, sharing links "
    "and linked Sheets data. None of that is in the .pptx and none of it is cleaned by this tool.",
]

ATTESTATION = [
    "I opened the cleaned deck in PowerPoint, Keynote or Google Slides and it displays correctly.",
    "I looked at every remaining image at full size (face/eyes, marks/tattoos/jewelry, location clues, "
    "wristbands/badges/screens, burned-in text) and cropped, covered or replaced what identifies the patient.",
    "I played every remaining audio/video clip, or removed it.",
    "I considered small-cell counts and rare-case details (rare diagnosis + place, occupation, event).",
    "I considered dates and timing near rare events, and ages over 89.",
    "For Google Slides: the cleaned deck was uploaded as a NEW file; the original's sharing, version history "
    "and comments are handled per my institution's policy.",
    "I followed my institution's de-identification / presentation policy.",
]


def write_report(path, ctx):
    L = ["# PHI scrub report", "",
         "Masked values only. Tool: phi-scrub-slides %s. Run: %s" % (VERSION, ctx["when"]), ""]
    if ctx["forced"]:
        L += ["> **WARNING: --force WAS USED. The plan had unresolved items and/or verification failed. "
              "THIS DECK MAY STILL CONTAIN PHI. Do not present or share it.**", ""]
        L += ["> - %s" % e for e in ctx["plan_errors"]] + [""]
    L += ["- Input (unchanged): `%s` (sha256 %s)" % (ctx["input"], ctx["sha"][:16]),
          "- Output: `%s`" % (ctx["output"] or "NOT WRITTEN"),
          "- Verification: **%s**" % ("PASS" if not ctx["problems"] else "FAIL"), ""]
    if ctx["problems"]:
        L += ["## Verification problems", ""] + ["- %s" % p for p in ctx["problems"]] + [""]
    L += ["## What was removed or replaced", ""]
    L += ["- %d x %s -> %s" % (n, m, r) for m, r, n in ctx["replacements"]]
    L += ["- %s" % x for x in ctx["log"]]
    if ctx["warnings"]:
        L += ["", "## Warnings", ""] + ["- %s" % w for w in ctx["warnings"]]
    if ctx["kept"]:
        L += ["", "## Kept on purpose (decision: keep)", ""] + ["- %s: %s" % k for k in ctx["kept"]]
    L += ["", "## Do by hand before presenting", ""]
    L += ["- crop or cover this region by hand: %s" % b for b in ctx["by_hand"]] or ["- (no manual edits recorded)"]
    L += ["", "## Residual risk: what this tool cannot see", ""] + ["- %s" % r for r in RESIDUAL_RISK]
    L += ["", "## Human attestation checklist (HIPAA Safe Harbor, 45 CFR 164.514(b)(2))", "",
          "Tick each category after checking the WHOLE deck: text, pictures, charts, tables.", ""]
    L += ["- [ ] %d. %s" % (i, c) for i, c in enumerate(SAFE_HARBOR_18, 1)]
    L += [""] + ["- [ ] %s" % a for a in ATTESTATION]
    L += ["", "This tool reduces risk. It does not certify de-identification and is not legal advice.", ""]
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Apply a fully decided PHI scrub plan to a .pptx (stdlib only). "
                                             "Writes a new file; fails closed.")
    ap.add_argument("pptx")
    ap.add_argument("--plan", required=True, help="plan JSON from scan_pptx.py --plan-out, fully decided")
    ap.add_argument("--out", help="output path (default: <name>.deidentified.pptx)")
    ap.add_argument("--strip-macros", action="store_true", help="remove VBA macros from a .pptm instead of refusing")
    ap.add_argument("--force", action="store_true",
                    help="write output despite unresolved items or failed verification (LOUD warning in report)")
    ap.add_argument("--overwrite", action="store_true", help="replace an existing output file")
    ap.add_argument("--show-values", action="store_true", help="print raw values in the terminal summary")
    args = ap.parse_args(argv)

    src = os.path.abspath(args.pptx)
    out = os.path.abspath(args.out or os.path.splitext(args.pptx)[0] + ".deidentified.pptx")
    if out.lower().endswith(".pptm"):
        out = out[:-5] + ".pptx"
    if out == src:
        raise SystemExit("refusing to overwrite the input deck; choose a different --out")
    if os.path.exists(out) and not args.overwrite:
        raise SystemExit("output exists: %s (use --overwrite)" % out)
    report_path = os.path.splitext(out)[0] + ".scrub-report.md"

    with open(args.plan, encoding="utf-8") as fh:
        plan = json.load(fh)
    report, _ = scan_pptx.scan(src)
    reps, actions, images, errors, kept = check_plan(plan, report, args.strip_macros)
    if errors and not args.force:
        print("REFUSING: the plan is not fully decided (fail-closed). Nothing was written.")
        for e in errors:
            print("  - " + (e if args.show_values else e))
        print("Decide every item in the plan (see its _HOW_TO), or pass --force (not recommended).")
        return 1
    if errors:
        print("!" * 72 + "\n--force: proceeding with %d unresolved plan item(s). Output MAY CONTAIN PHI.\n" % len(errors)
              + "!" * 72)

    pkg = Package(src)
    s = Scrubber(pkg, reps, actions, images, args.strip_macros or args.force)
    s.run()
    tmp = os.path.join(os.path.dirname(out), ".%s.partial" % os.path.basename(out))
    pkg.save(tmp)

    # Final verification pass: re-scan the output, grep for every original value.
    out_report, out_pkg = scan_pptx.scan(tmp)
    out_report["input"] = tmp
    originals = {f["value"]: f["category"] for f in report["findings"] if f["value"] not in kept}
    problems = scan_pptx.verify_output(out_report, out_pkg, {"replacements": reps, "actions": actions},
                                       originals, kept)
    problems += ["unresolved: " + u for u in s.unresolved]

    written = None
    if problems and not args.force:
        os.remove(tmp)
    else:
        os.replace(tmp, out)
        written = out

    show = (lambda r: r["find"]) if args.show_values else (lambda r: mask(r["find"], r.get("category")))
    ctx = {
        "when": datetime.datetime.now().isoformat(timespec="seconds"), "forced": bool(args.force and (errors or problems)),
        "plan_errors": errors, "input": src, "sha": report["sha256"], "output": written, "problems": problems,
        "replacements": [(mask(r["find"], r.get("category")), r["replace"], s.rep.counts[r["find"]]) for r in reps],
        "log": s.log, "warnings": s.warnings, "by_hand": s.by_hand,
        "kept": [(f["id"], "%s %s" % (f["category"], f["masked"])) for f in report["findings"] if f["value"] in kept],
    }
    write_report(report_path, ctx)

    print("phi-scrub-slides %s" % VERSION)
    print("Input (unchanged): %s" % src)
    print("Text replacements:")
    for r in reps:
        print("  %4d x  %-40s -> %s" % (s.rep.counts[r["find"]], show(r), r["replace"]))
    print("Structural changes:")
    for line in s.log:
        print("  - " + line)
    for w in s.warnings:
        print("  ! " + w)
    print("Final verification (re-scan + search for every original value): %s" % ("PASS" if not problems else "FAIL"))
    for p in problems:
        print("  ! " + p)
    print("Report (masked, with residual-risk + attestation checklist): %s" % report_path)
    if written:
        print("Output: %s" % written)
        if problems or errors:
            print("!" * 72 + "\nWRITTEN WITH --force. THIS DECK MAY STILL CONTAIN PHI.\n" + "!" * 72)
    else:
        print("NO OUTPUT WRITTEN (fail-closed). Fix the plan and run again.")
    print("Still to check BY EYE: every remaining image, audio/video, charts, rare-case details.")
    for b in s.by_hand:
        print("  - crop or cover this region by hand: " + b)
    if problems:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
