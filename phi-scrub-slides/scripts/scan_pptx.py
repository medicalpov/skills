#!/usr/bin/env python3
"""Inventory every place PHI can hide in a .pptx (PowerPoint, or Google Slides
downloaded as .pptx). Python standard library only. Never modifies the deck.

    python3 scan_pptx.py deck.pptx --out deck.phi-report.json \
        --plan-out deck.plan.json --images-dir deck_images

The terminal shows masked values ("J*** T**********", "**/**/1950"); raw values
only with --show-values. The report, the plan and the images folder contain raw
PHI: keep them on this computer and delete them when done.

Check a scrubbed deck against the plan (exit code 2 on any problem):

    python3 scan_pptx.py deck.deidentified.pptx --verify-plan deck.plan.json
"""

import argparse
import hashlib
import io
import json
import os
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from phi_common import (  # noqa: E402
    APP_KEEP, AV_EXTS, CONF_RANK, CONTEXT_WORDS, IMAGE_EXTS, P, PHI_WARNING, PLACEHOLDERS, R, VERSION,
    Package, _find_cnvpr, build_context, check_package, detect, image_dims, image_metadata,
    is_xml_name, local, mask, text_items,
)

CHECKLIST = [
    ("face", "Face or eyes (full or partial face, eye close-ups)"),
    ("marks", "Tattoos, birthmarks, scars, piercings, jewelry or other distinctive marks"),
    ("location", "Location clues: room/bed numbers, hospital or clinic signs, logos, windows/views, street signs"),
    ("ids_in_frame", "Wristbands, ID badges, charts, monitors or screens in frame"),
    ("burned_in_text", "Burned-in text: date stamps, DICOM/EHR overlays, rulers or labels with patient info"),
]

# Fail-closed defaults (a plan may only loosen them explicitly).
DEFAULT_ACTIONS = {
    "strip_metadata": True,        # docProps personal fields, custom props, revision history
    "remove_custom_xml": True,
    "remove_tags": True,           # ppt/tags (hidden add-in data)
    "remove_comments": True,       # comments + comment authors
    "speaker_notes": "delete",     # delete | keep
    "remove_thumbnail": True,      # docProps/thumbnail.jpeg is a picture of slide 1
    "hidden_slides": "remove",     # remove | keep (keep makes their findings need decisions)
    "hyperlinks": "remove_all",    # remove_all | remove_matching | keep
    "alt_text": "clear_matched",   # clear_matched | clear_all | keep
    "strip_image_metadata": True,  # EXIF / GPS / XMP / IPTC / PNG text
    "embedded_objects": "remove",  # remove | keep  (xlsx/docx/OLE cannot be scanned safely)
}
IMAGE_DECISIONS = {"keep", "replace", "placeholder", "remove", "mask"}
AV_DECISIONS = {"keep", "remove"}
GPS_VALUE = "GPS coordinates in image file"


def _ext(name):
    return name.rsplit(".", 1)[-1].lower() if "." in name else ""


def _finding(cat, detector, conf, value, part, kind, location, replaceable, ctx, item=None, top=None):
    return {"category": cat, "detector": detector, "confidence": conf, "value": value,
            "masked": mask(value, cat), "part": part, "kind": kind, "location": location,
            "slides": ctx.get(part, {}).get("slides", []), "top": top,
            "replaceable": replaceable, "text_item": item}


# ---------------------------------------------------------------------------
def scan_images(pkg, ctx):
    media = [n for n in pkg.parts()
             if not n.lower().startswith("docprops/") and "/embeddings/" not in n.lower()
             and (_ext(n) in IMAGE_EXTS or _ext(n) in AV_EXTS or "/media/" in n.lower())]
    refs = {}
    for rel in pkg.all_rels():
        if not rel.external and rel.resolved in media:
            refs.setdefault(rel.resolved, []).append(rel)
    images, findings = [], []
    for n in media:
        data = pkg.blobs[n]
        slides, used_in, alts = set(), [], []
        for rel in refs.get(n, []):
            src_ctx = ctx.get(rel.source, {})
            slides.update(src_ctx.get("slides", []))
            used_in.append(src_ctx.get("label", rel.source))
            d = pkg.try_doc(rel.source) if is_xml_name(rel.source) else None
            if d is None:
                continue
            for pic in d.root.iter("{%s}pic" % P):
                if any(v == rel.id for el in pic.iter() for k, v in el.attrib.items() if k.startswith("{%s}" % R)):
                    cnv = _find_cnvpr(pic)
                    if cnv is not None:
                        alts += [cnv.get(a) for a in ("descr", "title") if cnv.get(a, "").strip()]
        meta = image_metadata(data)
        dims = image_dims(data)
        is_av = _ext(n) in AV_EXTS or pkg.content_type(n).startswith(("video/", "audio/"))
        images.append({
            "part": n, "type": _ext(n), "bytes": len(data),
            "width": dims[0] if dims else None, "height": dims[1] if dims else None,
            "slides": sorted(slides), "used_in": sorted(set(used_in)),
            "alt_text": alts, "audio_video": is_av,
            "exif_present": meta["exif_present"], "gps": meta["gps"], "gps_tags": meta["gps_tags"],
            "metadata_blocks": meta["blocks"], "metadata_flags": meta["flags"],
            "exif_fields": meta["exif_fields"],
            "note": ("audio/video cannot be inspected by this tool; remove it, or verify by ear and eye"
                     if is_av else "view it: faces, marks, location clues, IDs in frame, burned-in text"),
        })
        loc = "%s (image metadata)" % n
        if meta["gps"]:
            findings.append(_finding("location", "exif_gps", "high", GPS_VALUE,
                                     n, "image_metadata", loc, False, ctx))
        for field, val in meta["exif_fields"].items():
            cat = {"DateTime": "date", "DateTimeOriginal": "date", "DateTimeDigitized": "date",
                   "Artist": "name", "CameraOwnerName": "name", "XPAuthor": "name", "Copyright": "name",
                   "BodySerialNumber": "device_id", "LensSerialNumber": "device_id"}.get(field)
            if cat:
                findings.append(_finding(cat, "exif_" + field, "medium", val, n, "image_metadata",
                                         "%s EXIF %s" % (n, field), False, ctx))
        for txt in meta["texts"]:
            for h in detect(txt, context_hint=True):
                findings.append(_finding(h["category"], h["detector"], h["confidence"], h["value"],
                                         n, "image_metadata", loc, False, ctx))
    return images, findings


def scan_embeddings(pkg, ctx):
    out, findings = [], []
    chart_wb = set()
    for rel in pkg.all_rels():
        if not rel.external and "/charts/" in rel.source and rel.is_type("package", "oleObject"):
            chart_wb.add(rel.resolved)
    for n in pkg.parts():
        if "/embeddings/" not in n.lower():
            continue
        data = pkg.blobs[n]
        entry = {"part": n, "type": _ext(n), "bytes": len(data), "chart_workbook": n in chart_wb,
                 "where": ctx.get(n, {}).get("label", n), "scanned": False, "findings": 0}
        if data[:2] == b"PK":
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    entry["scanned"] = True
                    for inner in z.namelist():
                        if not is_xml_name(inner) or inner.endswith(".rels"):
                            continue
                        try:
                            root = ET.fromstring(z.read(inner))
                        except ET.ParseError:
                            continue
                        for el in root.iter():
                            if el.text and el.text.strip():
                                for h in detect(el.text):
                                    entry["findings"] += 1
                                    findings.append(_finding(h["category"], h["detector"], h["confidence"],
                                                             h["value"], n, "embedded_file",
                                                             "%s -> %s" % (entry["where"], inner), False, ctx))
            except zipfile.BadZipFile:
                pass
        out.append(entry)
    return out, findings


def scan_structure(pkg, ctx, items):
    s = {}
    meta = {}
    for name in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"):
        if not pkg.exists(name):
            continue
        d = pkg.try_doc(name)
        if d is None:
            continue
        for top in d.root:
            txt = " | ".join(t.strip() for t in top.itertext() if t.strip())
            key = "%s:%s" % (posixpath.basename(name), local(top.tag))
            if top.get("name"):
                key += "[%s]" % top.get("name")
            if txt:
                meta[key] = txt
    s["metadata"] = meta
    comment_parts, author_parts = [], []
    for rel in pkg.all_rels():
        if rel.external:
            continue
        if rel.is_type("comments"):
            comment_parts.append(rel.resolved)
        elif rel.is_type("commentAuthors", "authors"):
            author_parts.append(rel.resolved)
    n_comments = 0
    for c in comment_parts:
        d = pkg.try_doc(c) if pkg.exists(c) else None
        if d is not None:
            n_comments += sum(1 for el in d.root if local(el.tag) == "cm")
    s["comments"] = {"parts": sorted(set(comment_parts)), "count": n_comments,
                     "author_parts": sorted(set(author_parts))}
    notes = {}
    for it in items:
        if it["kind"] == "notes" and it["field"] == "paragraph":
            notes[it["part"]] = notes.get(it["part"], 0) + 1
    s["speaker_notes"] = [{"part": p, "where": ctx[p]["label"], "paragraphs": c} for p, c in sorted(notes.items())]
    slides = pkg.slides()
    s["slide_count"] = len(slides)
    s["hidden_slides"] = [x["num"] for x in slides if x["hidden"]]
    thumbs = [r.resolved for r in pkg.rels("") if r.is_type("thumbnail")]
    s["thumbnail"] = [t for t in thumbs if t and pkg.exists(t)]
    s["external_links"] = sum(1 for it in items if it["kind"] == "hyperlink")
    s["custom_xml_parts"] = [n for n in pkg.parts() if n.lower().startswith("customxml/") and not n.endswith(".rels")]
    s["tag_parts"] = [n for n in pkg.parts() if ctx.get(n, {}).get("kind") == "tags"]
    s["font_parts"] = [n for n in pkg.parts() if ctx.get(n, {}).get("kind") == "font"]
    s["revision_history_parts"] = [n for n in pkg.parts() if ctx.get(n, {}).get("kind") == "revision_history"]
    main_ct = pkg.content_type(pkg.main_part())
    s["macros"] = {"macro_enabled": "macroEnabled" in main_ct,
                   "parts": [n for n in pkg.parts() if ctx.get(n, {}).get("kind") == "macro"]}
    known = set(s["thumbnail"]) | set(s["font_parts"]) | set(s["macros"]["parts"])
    s["unscanned_binary_parts"] = [
        n for n in pkg.parts()
        if not is_xml_name(n) and "/media/" not in n.lower() and "/embeddings/" not in n.lower()
        and _ext(n) not in IMAGE_EXTS and _ext(n) not in AV_EXTS and n not in known
    ]
    s["parse_errors"] = dict(pkg.parse_errors)
    return s


def scan(path):
    pkg = Package(path)
    ctx = build_context(pkg)
    items = text_items(pkg, ctx)
    part_has_context = {it["part"] for it in items if CONTEXT_WORDS.search(it["text"])}
    findings = []
    for i, it in enumerate(items, 1):
        it["id"] = "T%04d" % i
        for h in detect(it["text"], context_hint=it["part"] in part_has_context):
            findings.append(_finding(h["category"], h["detector"], h["confidence"], h["value"],
                                     it["part"], it["kind"], it["location"], True, ctx, it["id"], it.get("top")))
    images, img_findings = scan_images(pkg, ctx)
    embeds, emb_findings = scan_embeddings(pkg, ctx)
    findings += img_findings + emb_findings
    for i, f in enumerate(findings, 1):
        f["id"] = "F%03d" % i
    structure = scan_structure(pkg, ctx, items)
    structure["embeddings"] = embeds
    with open(path, "rb") as fh:
        digest = hashlib.sha256(fh.read()).hexdigest()
    report = {
        "_WARNING": PHI_WARNING,
        "tool": "phi-scrub-slides scan %s" % VERSION,
        "input": os.path.abspath(path),
        "sha256": digest,
        "summary": {
            "findings": len(findings),
            "by_confidence": {c: sum(1 for f in findings if f["confidence"] == c) for c in ("high", "medium", "low")},
            "images": sum(1 for im in images if not im["audio_video"]),
            "audio_video": sum(1 for im in images if im["audio_video"]),
            "images_with_exif": sum(1 for im in images if im["exif_present"]),
            "images_with_gps": sum(1 for im in images if im["gps"]),
        },
        "findings": findings,
        "images": images,
        "structure": structure,
        "text_items": items,
    }
    return report, pkg


# ---------------------------------------------------------------------------
def covered_by(f, actions, hidden_slides):
    """Name of the structural action that removes this finding's location, or None."""
    k, part = f["kind"], f["part"]
    if f["slides"] and actions["hidden_slides"] == "remove" and all(s in hidden_slides for s in f["slides"]):
        return "hidden_slides"
    if k == "notes" and actions["speaker_notes"] == "delete":
        return "speaker_notes"
    if k in ("comments", "comment_authors") and actions["remove_comments"]:
        return "remove_comments"
    if k == "metadata" and actions["strip_metadata"]:
        if part.endswith(("core.xml", "custom.xml")):
            return "strip_metadata"
        if part.endswith("app.xml") and (f.get("top") or "") not in APP_KEEP:
            return "strip_metadata"
    if k == "revision_history" and actions["strip_metadata"]:
        return "strip_metadata"
    if k == "custom_xml" and actions["remove_custom_xml"]:
        return "remove_custom_xml"
    if k == "tags" and actions["remove_tags"]:
        return "remove_tags"
    if k == "hyperlink" and actions["hyperlinks"] == "remove_all":
        return "hyperlinks"
    if k == "image_metadata" and actions["strip_image_metadata"]:
        return "strip_image_metadata"
    if k in ("embedded_file", "embedding") and actions["embedded_objects"] == "remove":
        return "embedded_objects"
    return None


def suggested_plan(report):
    actions = dict(DEFAULT_ACTIONS)
    hidden = report["structure"]["hidden_slides"]
    reps = {}
    for f in report["findings"]:
        if f["replaceable"] and f["confidence"] in ("high", "medium") and not covered_by(f, actions, hidden):
            cur = reps.get(f["value"])
            if cur is None or CONF_RANK[f["confidence"]] > CONF_RANK[cur["confidence"]]:
                reps[f["value"]] = f
    decisions = []
    for f in report["findings"]:
        cov = covered_by(f, actions, hidden)
        if cov:
            dec = "removed_by:" + cov
        elif f["value"] in reps and f["replaceable"]:
            dec = "replace"
        else:
            dec = "UNRESOLVED"
        decisions.append({"id": f["id"], "decision": dec, "reason": "", "category": f["category"],
                          "confidence": f["confidence"], "location": f["location"], "value": f["value"]})
    replacements = [{"find": v, "replace": PLACEHOLDERS.get(f["category"], "[REDACTED]"),
                     "category": f["category"], "match_case": f["category"] != "name"}
                    for v, f in sorted(reps.items(), key=lambda kv: -len(kv[0]))]
    images = {}
    for im in report["images"]:
        images[im["part"]] = {
            "decision": "UNRESOLVED",
            "allowed": sorted(AV_DECISIONS if im["audio_video"] else IMAGE_DECISIONS - {"placeholder"}),
            "slides": im["slides"], "by_hand": "", "note": im["note"],
        }
    return {
        "_WARNING": PHI_WARNING,
        "_HOW_TO": (
            "Fail-closed plan. scrub_pptx.py refuses to run while any finding or image is UNRESOLVED. "
            "findings[].decision: 'replace' (value goes in replacements), 'keep' (needs a reason), or "
            "'removed_by:<action>' (only valid while that action is on). Add names, places and other "
            "context identifiers you find by reading to 'replacements'. images[].decision: keep | replace "
            "(grey placeholder) | remove | mask (needs Pillow, with 'boxes'); audio/video: keep | remove. "
            "Put manual crop/cover instructions in 'by_hand'."),
        "version": 2,
        "source_sha256": report["sha256"],
        "actions": actions,
        "replacements": replacements,
        "findings": decisions,
        "images": images,
    }


# ---------------------------------------------------------------------------
def review_name(im, idx):
    base = posixpath.basename(im["part"])
    if im["slides"]:
        return "slide%02d_%s" % (im["slides"][0], base)
    if im["used_in"]:
        return "template_%s" % base
    return "unused_%02d_%s" % (idx, base)


def write_image_review(report, pkg, out_dir, deck_name):
    os.makedirs(out_dir, exist_ok=True)
    lines = [
        "# Image review: %s" % deck_name, "",
        "> **%s**" % PHI_WARNING, "",
        "Every image needs a human look. The scrub strips hidden metadata (EXIF, GPS, XMP, IPTC, PNG text) "
        "by default, but no script can see what is *in* the picture.", "",
        "For each image tick what you see, then decide: **keep**, **crop by hand**, **cover a region by hand**, "
        "**replace with a grey placeholder**, or **remove**.", "",
        "Look for:", "",
    ]
    lines += ["- %s" % label for _k, label in CHECKLIST]
    lines.append("")
    ordered = sorted(report["images"], key=lambda x: (x["slides"] or [9999], x["part"]))
    for idx, im in enumerate(ordered, 1):
        fname = review_name(im, idx)
        im["review_file"] = fname
        with open(os.path.join(out_dir, fname), "wb") as fh:
            fh.write(pkg.blobs[im["part"]])
        where = ("slide " + ", ".join(str(s) for s in im["slides"])) if im["slides"] else (
            ", ".join(im["used_in"]) or "not used on any slide")
        dims = "%sx%s px" % (im["width"], im["height"]) if im["width"] else "size unknown"
        lines += ["## %d. %s" % (idx, fname), "",
                  "- Where: %s" % where,
                  "- In the deck: `%s` (%s, %s, %d bytes)" % (im["part"], im["type"].upper(), dims, im["bytes"])]
        if im["audio_video"]:
            lines.append("- **Audio/video: cannot be inspected by this tool. Remove it, or verify every frame "
                         "and the soundtrack by eye and ear.**")
        lines.append("- Alt text: %s" % ("; ".join('"%s"' % a for a in im["alt_text"]) if im["alt_text"] else "(none)"))
        lines.append("- EXIF: %s%s" % ("yes" if im["exif_present"] else "no",
                                        " - **GPS location found**" if im["gps"] else ""))
        if im["metadata_flags"]:
            lines.append("- Metadata: %s (stripped by default)" % "; ".join(im["metadata_flags"]))
            for k, v in sorted(im["exif_fields"].items()):
                lines.append("  - %s: %s" % (k, v))
        lines.append("- Visual check:")
        lines += ["  - [ ] %s" % label for _k, label in CHECKLIST]
        lines += ["- Decision: keep / crop by hand / cover by hand (where: ______) / replace with placeholder / remove", ""]
    with open(os.path.join(out_dir, "images.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


# ---------------------------------------------------------------------------
def print_summary(report, show_values, paths):
    v = (lambda f: f["value"]) if show_values else (lambda f: f["masked"])
    s = report["summary"]
    print("PHI scan: %s" % os.path.basename(report["input"]))
    for label, p in paths:
        if p:
            print("%s (CONTAINS PHI, keep local): %s" % (label, p))
    print("Findings: %d  (high %d, medium %d, low %d)%s" % (
        s["findings"], s["by_confidence"]["high"], s["by_confidence"]["medium"], s["by_confidence"]["low"],
        "" if show_values else "  [values masked; --show-values prints them]"))
    for f in report["findings"]:
        print("  %-5s %-6s %-12s %-52.52s %s" % (f["id"], f["confidence"], f["category"], f["location"], v(f)))
    st = report["structure"]
    print("Structure:")
    print("  document metadata fields: %s" % (", ".join(sorted(st["metadata"])) or "none"))
    print("  comments: %d (comment-author parts: %d)" % (st["comments"]["count"], len(st["comments"]["author_parts"])))
    print("  speaker notes: %d slide(s) with text" % len(st["speaker_notes"]))
    print("  hidden slides: %s" % (", ".join(map(str, st["hidden_slides"])) or "none"))
    print("  thumbnail (a picture of slide 1): %s" % ("present" if st["thumbnail"] else "none"))
    print("  external links: %d" % st["external_links"])
    print("  embedded objects: %d%s" % (len(st["embeddings"]),
                                        " (removed by default; cannot be scanned safely)" if st["embeddings"] else ""))
    if st["custom_xml_parts"] or st["tag_parts"]:
        print("  custom XML parts: %d, tag parts: %d (removed by default)" % (len(st["custom_xml_parts"]), len(st["tag_parts"])))
    if st["font_parts"]:
        print("  embedded fonts: %d (kept; fonts do not carry patient data)" % len(st["font_parts"]))
    if st["macros"]["macro_enabled"] or st["macros"]["parts"]:
        print("  MACROS: %s (scrub refuses unless --strip-macros)" % (", ".join(st["macros"]["parts"]) or "macro-enabled deck"))
    if st["revision_history_parts"]:
        print("  revision history parts: %d" % len(st["revision_history_parts"]))
    if st["unscanned_binary_parts"]:
        print("  NOT scanned (binary): %s" % ", ".join(st["unscanned_binary_parts"]))
    if st["parse_errors"]:
        print("  could not parse: %s" % ", ".join(st["parse_errors"]))
    print("Images: %d, audio/video: %d. Each needs a human decision (EXIF in %d, GPS in %d)" % (
        s["images"], s["audio_video"], s["images_with_exif"], s["images_with_gps"]))
    for im in report["images"]:
        where = ("slide " + ",".join(map(str, im["slides"]))) if im["slides"] else "template/unused"
        flags = "; ".join(im["metadata_flags"]) or "no metadata"
        print("  %-30s %-12s %s%s" % (im["part"], where, flags,
                                      "  [AUDIO/VIDEO: cannot be inspected; remove or verify by ear/eye]"
                                      if im["audio_video"] else ""))


def verify_output(out_report, out_pkg, plan, original_values, kept_values=()):
    """Fail-closed checks on a scrubbed deck. Returns a list of problems (masked).

    original_values: {raw value: category} for every original finding not kept by decision,
    plus every planned replacement. Each is searched (case-insensitive) in the decoded text
    and in every remaining XML part."""
    problems = []
    kept = {v.lower() for v in kept_values}
    texts = [it["text"].lower() for it in out_report["text_items"]]
    xml_raw = [out_pkg.blobs[n].decode("utf-8", "ignore").lower() for n in out_pkg.parts() if is_xml_name(n)]
    values = dict(original_values)
    for rep in plan.get("replacements", []):
        values.setdefault(rep["find"], rep.get("category"))
    for val, cat in values.items():
        low = val.lower()
        if low in kept or val == GPS_VALUE:
            continue
        esc = escape(low, {'"': "&quot;"})
        hits = sum(t.count(low) for t in texts) + sum(r.count(low) + (r.count(esc) if esc != low else 0)
                                                      for r in xml_raw)
        if hits:
            problems.append("value still present %d time(s): %s" % (hits, mask(val, cat)))
    for f in out_report["findings"]:
        if f["value"].lower() not in kept:
            problems.append("re-scan finding %s (%s, %s) at %s: %s" % (
                f["id"], f["category"], f["confidence"], f["location"], f["masked"]))
    a = dict(DEFAULT_ACTIONS)
    a.update(plan.get("actions", {}))
    st = out_report["structure"]
    if a["remove_comments"] and (st["comments"]["parts"] or st["comments"]["author_parts"]):
        problems.append("comments still present")
    if a["speaker_notes"] == "delete" and st["speaker_notes"]:
        problems.append("speaker notes still present")
    if a["remove_thumbnail"] and st["thumbnail"]:
        problems.append("thumbnail still present")
    if a["hidden_slides"] == "remove" and st["hidden_slides"]:
        problems.append("hidden slides still present")
    if a["embedded_objects"] == "remove" and st["embeddings"]:
        problems.append("embedded objects still present: %s" % ", ".join(e["part"] for e in st["embeddings"]))
    if a["remove_custom_xml"] and st["custom_xml_parts"]:
        problems.append("custom XML still present")
    if a["remove_tags"] and st["tag_parts"]:
        problems.append("tag parts still present")
    if st["macros"]["parts"] or st["macros"]["macro_enabled"]:
        problems.append("macros still present")
    if a["hyperlinks"] == "remove_all" and st["external_links"]:
        problems.append("external links still present")
    if a["strip_metadata"]:
        left = [k for k in st["metadata"] if k.startswith(("core.xml:", "custom.xml:"))
                or (k.startswith("app.xml:") and k.split(":", 1)[1] not in APP_KEEP)]
        if left:
            problems.append("document metadata still present: %s" % ", ".join(left))
    if a["strip_image_metadata"]:
        for im in out_report["images"]:
            if im["exif_present"] or im["gps"] or im["metadata_blocks"]:
                problems.append("image metadata still present in %s (%s)" % (im["part"], "; ".join(im["metadata_flags"])))
    problems += ["package: " + p for p in check_package(out_report["input"])]
    return problems


def plan_values(plan):
    """(original values not kept, kept values) from a plan's finding decisions."""
    kept = [d["value"] for d in plan.get("findings", []) if d.get("decision") == "keep" and d.get("value")]
    orig = {d["value"]: d.get("category") for d in plan.get("findings", [])
            if d.get("value") and d.get("decision") != "keep"}
    return orig, kept


def main(argv=None):
    ap = argparse.ArgumentParser(description="Scan a .pptx (PowerPoint or Google Slides export) for PHI. "
                                             "Stdlib only; never modifies the deck.")
    ap.add_argument("pptx")
    ap.add_argument("--out", help="write the full JSON report here (CONTAINS PHI)")
    ap.add_argument("--plan-out", help="write a fail-closed draft plan here (CONTAINS PHI)")
    ap.add_argument("--images-dir", help="extract images + images.md review index here (CONTAINS PHI)")
    ap.add_argument("--verify-plan", help="check a scrubbed deck against this plan (exit 2 on any problem)")
    ap.add_argument("--show-values", action="store_true", help="print raw identifiers to the terminal")
    args = ap.parse_args(argv)

    report, pkg = scan(args.pptx)
    if args.images_dir:
        write_image_review(report, pkg, args.images_dir, os.path.basename(args.pptx))
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(report, fh, indent=2, ensure_ascii=False)
    if args.plan_out:
        with open(args.plan_out, "w", encoding="utf-8") as fh:
            json.dump(suggested_plan(report), fh, indent=2, ensure_ascii=False)
    print_summary(report, args.show_values, [("Report", args.out), ("Draft plan", args.plan_out),
                                             ("Image review folder", args.images_dir)])
    if args.verify_plan:
        with open(args.verify_plan, encoding="utf-8") as fh:
            plan = json.load(fh)
        orig, kept = plan_values(plan)
        problems = verify_output(report, pkg, plan, orig, kept)
        print("Verification against plan:")
        if problems:
            for p in problems:
                print("  FAIL " + p)
            return 2
        print("  PASS: no planned or originally found identifier remains; no comments, notes, hidden slides, "
              "thumbnail, embedded objects, document or image metadata; package structure consistent.")
        print("  Still required: a human look at every image and the attestation checklist.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
