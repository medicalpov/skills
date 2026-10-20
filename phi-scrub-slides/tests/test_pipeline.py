"""End-to-end tests: build a synthetic deck -> scan -> decide plan -> scrub -> verify.

Run:  python3 -m unittest discover -s tests -v     (or tests/run_tests.sh)
Stdlib only. All identifiers are fake (see make_fixture.py).
"""

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
sys.path.insert(0, HERE)
sys.path.insert(0, SCRIPTS)

import make_fixture  # noqa: E402
from phi_common import check_package, image_metadata  # noqa: E402
import scan_pptx  # noqa: E402

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# What Claude adds after reading the inventory: names / places the regexes cannot know.
CONTEXT_PHI = [
    {"find": "Jane Testpatient", "replace": "[PATIENT]", "category": "name", "match_case": False},
    {"find": "Testpatient", "replace": "[PATIENT]", "category": "name", "match_case": False},
    {"find": "Fake Hospital of Testville", "replace": "[LOCATION]", "category": "location", "match_case": True},
    {"find": "Springfield", "replace": "[LOCATION]", "category": "location", "match_case": True},
]
PHI_WORDS = ("testpatient", "springfield", "testville", "fakename")


def sha(path):
    with open(path, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()


def run(script, *args):
    return subprocess.run([sys.executable, os.path.join(SCRIPTS, script)] + list(args),
                          capture_output=True, text=True)


def decide(plan, image_decisions=None):
    """Simulate the reviewed, approved plan: every finding and image decided."""
    plan["replacements"] = CONTEXT_PHI + plan["replacements"]
    for d in plan["findings"]:
        if d["decision"] != "UNRESOLVED":
            continue
        if any(w in d["value"].lower() for w in PHI_WORDS):
            d["decision"] = "replace"
        else:
            d["decision"] = "keep"
            d["reason"] = "synthetic test: not patient-identifying (presenter name / generic text)"
    image_decisions = image_decisions or {
        "ppt/media/image1.png": {"decision": "replace"},
        "ppt/media/image2.jpeg": {"decision": "keep", "by_hand": "lower-left corner: tattoo on forearm"},
    }
    for part, spec in image_decisions.items():
        plan["images"][part].update(spec)
    return plan


class PipelineTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="phi_scrub_test_")
        self.deck = make_fixture.build(os.path.join(self.tmp, "case.pptx"))
        self.plan_path = os.path.join(self.tmp, "case.plan.json")
        self.out = os.path.join(self.tmp, "case.deidentified.pptx")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def make_plan(self, **kw):
        r = run("scan_pptx.py", self.deck, "--out", os.path.join(self.tmp, "r.json"),
                "--plan-out", self.plan_path, "--images-dir", os.path.join(self.tmp, "imgs"))
        self.assertEqual(r.returncode, 0, r.stderr)
        self.scan_stdout = r.stdout
        with open(self.plan_path) as fh:
            plan = json.load(fh)
        plan = decide(plan, **kw)
        with open(self.plan_path, "w") as fh:
            json.dump(plan, fh)
        return plan

    # ------------------------------------------------------------------
    def test_fixture_is_consistent(self):
        self.assertEqual(check_package(self.deck), [])

    def test_scan_masks_terminal_output_and_finds_exif_gps(self):
        self.make_plan()
        for w in ("testpatient", "00012345", "01/02/1950", "555-0100", "fakename"):
            self.assertNotIn(w, self.scan_stdout.lower(), "raw value leaked to terminal: " + w)
        self.assertIn("GPS location found", self.scan_stdout)
        with open(os.path.join(self.tmp, "imgs", "images.md")) as fh:
            index = fh.read()
        self.assertIn("slide03_image2.jpeg", index)
        self.assertIn("GPS location found", index)
        self.assertIn("Tattoos", index)
        self.assertTrue(os.path.exists(os.path.join(self.tmp, "imgs", "slide02_image1.png")))

    def test_full_pipeline(self):
        before = sha(self.deck)
        self.make_plan()
        r = run("scrub_pptx.py", self.deck, "--plan", self.plan_path, "--out", self.out)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertEqual(sha(self.deck), before, "input deck was modified")
        self.assertNotIn("testpatient", r.stdout.lower(), "raw value leaked to terminal")
        self.assertIn("Final verification (re-scan + search for every original value): PASS", r.stdout)

        # Valid zip, every XML part parses, rels / content types consistent.
        self.assertEqual(check_package(self.out), [])
        with zipfile.ZipFile(self.out) as z:
            self.assertIsNone(z.testzip())
            names = z.namelist()
            self.assertEqual(names[0], "[Content_Types].xml")
            blobs = {n: z.read(n) for n in names}
        for n, b in blobs.items():
            if n.endswith((".xml", ".rels")):
                ET.fromstring(b)

        # Zero fake identifiers anywhere (every part, XML and binary).
        for n, b in blobs.items():
            low = b.decode("latin-1").lower()
            for ident in make_fixture.FAKE_IDENTIFIERS:
                self.assertNotIn(ident, low, "%r survived in %s" % (ident, n))

        # Structural removals.
        for gone in ("docProps/thumbnail.jpeg", "ppt/commentAuthors.xml", "ppt/comments/comment1.xml",
                     "ppt/slides/slide4.xml", "ppt/media/image1.png"):
            self.assertNotIn(gone, names)
        self.assertFalse([n for n in names if "notesSlide" in n])
        self.assertFalse([n for n in names if "/embeddings/" in n])
        self.assertIn("ppt/media/image1_redacted.png", names)

        # Kept JPEG: same pixels, no EXIF / GPS.
        meta = image_metadata(blobs["ppt/media/image2.jpeg"])
        self.assertFalse(meta["exif_present"] or meta["gps"] or meta["blocks"], meta)
        self.assertTrue(blobs["ppt/media/image2.jpeg"].endswith(b"\x3f\xff\xd9"))

        # Run-split paragraph rewritten into the first run, formatting runs kept.
        slide2 = ET.fromstring(blobs["ppt/slides/slide2.xml"])
        first = [p for p in slide2.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}p")][1]
        runs = [t.text or "" for t in first.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")]
        self.assertEqual("".join(runs), "[PATIENT] is a [AGE 90+] seen in clinic on [DATE].")
        self.assertEqual(len(runs), 4)
        # mc:Ignorable prefix still declared after re-serialization.
        self.assertRegex(blobs["ppt/slides/slide2.xml"].decode(), r'xmlns:p14="')
        self.assertIn('mc:Ignorable="p14"', blobs["ppt/slides/slide2.xml"].decode())

        # Hidden slide removed with presentation.xml + rels + sections consistent.
        pres = ET.fromstring(blobs["ppt/presentation.xml"])
        sld_ids = pres.find("{%s}sldIdLst" % P)
        self.assertEqual(len(sld_ids), 4)
        prels = ET.fromstring(blobs["ppt/_rels/presentation.xml.rels"])
        rel_ids = {el.get("Id"): el.get("Target") for el in prels}
        for s in sld_ids:
            self.assertIn(s.get("{%s}id" % R), rel_ids)
        self.assertNotIn("slides/slide4.xml", rel_ids.values())
        self.assertNotIn('id="259"', blobs["ppt/presentation.xml"].decode())

        # Chart kept, its workbook gone, rels valid; OLE object removed.
        self.assertIn("ppt/charts/chart1.xml", names)
        self.assertNotIn("externalData", blobs["ppt/charts/chart1.xml"].decode())
        self.assertNotIn("embeddings", blobs.get("ppt/charts/_rels/chart1.xml.rels", b"").decode())
        self.assertIn("Lesion size", blobs["ppt/charts/chart1.xml"].decode())
        slide5 = blobs["ppt/slides/slide5.xml"].decode()
        self.assertNotIn("oleObj", slide5)
        self.assertNotIn("AlternateContent", slide5)

        # Metadata, notes, links gone; report written with residual risk + 18 categories.
        self.assertNotIn("creator", blobs["docProps/core.xml"].decode())
        self.assertNotIn("hlinkClick", blobs["ppt/slides/slide2.xml"].decode())
        report = os.path.splitext(self.out)[0] + ".scrub-report.md"
        with open(report) as fh:
            md = fh.read()
        self.assertIn("Residual risk: what this tool cannot see", md)
        self.assertEqual(len(re.findall(r"^- \[ \] \d+\. ", md, re.M)), 18)
        self.assertIn("crop or cover this region by hand", md)
        self.assertNotIn("testpatient", md.lower())

        # Independent re-scan of the output passes.
        v = run("scan_pptx.py", self.out, "--verify-plan", self.plan_path)
        self.assertEqual(v.returncode, 0, v.stdout)

    def test_verification_catches_unscrubbed_deck(self):
        self.make_plan()
        v = run("scan_pptx.py", self.deck, "--verify-plan", self.plan_path)
        self.assertEqual(v.returncode, 2)
        self.assertIn("FAIL", v.stdout)
        self.assertNotIn("testpatient", v.stdout.lower())

    def test_refuses_unresolved_image(self):
        self.make_plan(image_decisions={"ppt/media/image1.png": {"decision": "replace"}})
        r = run("scrub_pptx.py", self.deck, "--plan", self.plan_path, "--out", self.out)
        self.assertEqual(r.returncode, 1)
        self.assertIn("UNRESOLVED image ppt/media/image2.jpeg", r.stdout)
        self.assertFalse(os.path.exists(self.out))

    def test_refuses_unresolved_finding(self):
        plan = self.make_plan()
        target = next(d for d in plan["findings"] if d["decision"] == "keep")
        target["decision"] = "UNRESOLVED"
        with open(self.plan_path, "w") as fh:
            json.dump(plan, fh)
        r = run("scrub_pptx.py", self.deck, "--plan", self.plan_path, "--out", self.out)
        self.assertEqual(r.returncode, 1)
        self.assertIn("UNRESOLVED finding %s" % target["id"], r.stdout)
        self.assertFalse(os.path.exists(self.out))

    def test_refuses_keep_without_reason_and_wrong_deck(self):
        plan = self.make_plan()
        target = next(d for d in plan["findings"] if d["decision"] == "keep")
        target["reason"] = ""
        plan["source_sha256"] = "0" * 64
        with open(self.plan_path, "w") as fh:
            json.dump(plan, fh)
        r = run("scrub_pptx.py", self.deck, "--plan", self.plan_path, "--out", self.out)
        self.assertEqual(r.returncode, 1)
        self.assertIn("without a reason", r.stdout)
        self.assertIn("different file", r.stdout)

    def test_force_writes_loud_warning(self):
        self.make_plan(image_decisions={"ppt/media/image1.png": {"decision": "replace"}})
        r = run("scrub_pptx.py", self.deck, "--plan", self.plan_path, "--out", self.out, "--force")
        self.assertTrue(os.path.exists(self.out), r.stdout)
        self.assertIn("--force", r.stdout)
        with open(os.path.splitext(self.out)[0] + ".scrub-report.md") as fh:
            self.assertIn("--force WAS USED", fh.read())

    def test_macro_deck_refused_then_stripped(self):
        self.deck = make_fixture.build(os.path.join(self.tmp, "case.pptm"), macro=True)
        self.make_plan()
        r = run("scrub_pptx.py", self.deck, "--plan", self.plan_path, "--out", self.out)
        self.assertEqual(r.returncode, 1)
        self.assertIn("macros", r.stdout.lower())
        self.assertFalse(os.path.exists(self.out))
        r = run("scrub_pptx.py", self.deck, "--plan", self.plan_path, "--out", self.out, "--strip-macros")
        self.assertEqual(r.returncode, 0, r.stdout)
        self.assertIn("MACROS STRIPPED", r.stdout)
        with zipfile.ZipFile(self.out) as z:
            self.assertNotIn("ppt/vbaProject.bin", z.namelist())
            ct = z.read("[Content_Types].xml").decode()
        self.assertNotIn("macroEnabled", ct)
        self.assertNotIn("vbaProject", ct)
        self.assertEqual(check_package(self.out), [])

    def test_draft_plan_is_fail_closed(self):
        report, _ = scan_pptx.scan(self.deck)
        plan = scan_pptx.suggested_plan(report)
        self.assertTrue(all(v["decision"] == "UNRESOLVED" for v in plan["images"].values()))
        self.assertEqual(plan["actions"]["embedded_objects"], "remove")
        self.assertEqual(plan["actions"]["hyperlinks"], "remove_all")


if __name__ == "__main__":
    unittest.main()
