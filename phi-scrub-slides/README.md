# phi-scrub-slides

**Remove patient identifiers from PowerPoint and Google Slides decks before you present or share them.**

You photographed a lesion for grand rounds. The photo also shows a tattoo, the clinic room number on
the door behind the patient, and, invisibly, a GPS tag and the date inside the image file. The
speaker notes still say the patient's name, a hidden backup slide has the MRN, and the file's
"author" field is the clinic's shared login.

This tool is built for decks like that: clinical photography (dermatology, wound care, plastics,
ophthalmology, ENT, radiology screenshots), case presentations and M&M slides. It looks in every place a
.pptx can hide PHI, strips what can be stripped automatically, and gives you a checklist for what only
a human can judge.

It is a Claude skill **and** two plain Python scripts. You can use it with Claude, or entirely offline
with no AI at all.

---

## What it checks

| Where | What happens |
|---|---|
| Slide text, tables, grouped shapes, SmartArt, chart labels | Pattern search + your review; replaced with `[PATIENT]`, `[DATE]`, `[MRN]`, `[AGE 90+]`, `[LOCATION]` ... |
| Speaker notes | Deleted by default (Google Slides exports them too) |
| Comments and comment authors | Deleted |
| Hidden slides | Removed |
| Alt text on pictures | Replaced; cleared if it held an identifier |
| Footers, slide layouts and masters | Searched and replaced |
| Hyperlinks | Removed (links often carry MRNs or portal IDs) |
| File properties: author, last modified by, title, company, keywords | Cleared |
| Thumbnail (a hidden picture of slide 1 inside the file) | Deleted |
| **Photos: EXIF, GPS location, camera, owner, date, serial number, XMP, IPTC** | Reported per image, stripped by default |
| **Photos: what is in the picture** | You decide per image (see below) |
| Embedded Excel/Word/OLE objects | Removed (cannot be scanned safely); charts keep their picture |
| Audio / video | Cannot be inspected: remove, or verify by eye and ear |
| Macros (.pptm) | Refused, or stripped on request |

### The 18 HIPAA Safe Harbor identifiers

1. Names
2. Geographic units smaller than a state (street, city, county, ZIP code)
3. All elements of dates (except year) related to the patient, and all ages over 89
4. Telephone numbers
5. Fax numbers
6. Email addresses
7. Social Security numbers
8. Medical record numbers
9. Health plan beneficiary numbers
10. Account numbers
11. Certificate / license numbers
12. Vehicle identifiers and license plates
13. Device identifiers and serial numbers
14. Web URLs
15. IP addresses
16. Biometric identifiers (fingerprints, voice prints)
17. Full-face photographs and any comparable images
18. Any other unique identifying number, characteristic or code (a tattoo, a rare condition in a small town)

Pattern matching can find many of 3-15. Names (1), places (2) and 16-18 need a human, or Claude
reading the deck with you.

### Every image gets a decision

For each picture, look for:

- face or eyes;
- tattoos, birthmarks, scars, piercings, jewelry or other distinctive marks;
- location clues: room or bed numbers, hospital or clinic signs, logos, windows or views, street signs;
- wristbands, ID badges, charts, monitors or screens in frame;
- burned-in text: date stamps, DICOM or EHR overlays, rulers labelled with patient info.

Then choose: **keep**, **replace** with a grey placeholder, **remove**, or keep it and **crop or cover a
region by hand** in PowerPoint (the report reminds you exactly where). The tool will not write a cleaned
file until every image has a decision.

---

## Use it with Claude

1. Install the skill: copy the `phi-scrub-slides` folder into your Claude skills folder
   (Claude Code: `~/.claude/skills/`), or upload it as a skill in the Claude app.
2. Say: *"Remove the PHI from my slides"* and attach the .pptx.
3. Claude scans, reads the deck for names and context, asks before looking at your images, shows you a
   table of what it found (masked), and waits for your approval before writing a new file.

**Important:** using it inside Claude sends your deck to Claude. Only do this with real patient data in
an environment your institution has approved for PHI (for example one covered by a Business Associate
Agreement). Otherwise, use the offline route below.

## Run it without AI (offline, two commands)

You need Python 3 (preinstalled on most Macs; nothing else to install).

```bash
# 1. Scan: makes a report, a draft plan, and a folder of the deck's images to look at
python3 scripts/scan_pptx.py MyTalk.pptx --out MyTalk.phi-report.json \
    --plan-out MyTalk.plan.json --images-dir MyTalk_review

# (open MyTalk.plan.json in a text editor: add names/places to "replacements",
#  and change every "UNRESOLVED" to a decision - see the _HOW_TO line at the top)

# 2. Scrub: writes MyTalk.deidentified.pptx and MyTalk.deidentified.scrub-report.md
python3 scripts/scrub_pptx.py MyTalk.pptx --plan MyTalk.plan.json
```

Shortcut if you want every image replaced with a grey box: set each image's `"decision"` to `"replace"`.

Optional check afterwards (the scrub already does this internally):

```bash
python3 scripts/scan_pptx.py MyTalk.deidentified.pptx --verify-plan MyTalk.plan.json
```

Your original file is never changed. The report, plan and review folder contain PHI: keep them on your
computer and delete them when you are done.

## Google Slides

1. In Slides: **File > Download > Microsoft PowerPoint (.pptx)**.
2. Scan and scrub the .pptx as above.
3. Upload the cleaned file as a **new** Slides file (Drive > New > File upload, or File > Import slides
   into a new presentation).

Traps specific to Google Slides:

- **Version history**: the original file keeps every past version, PHI included. Do not try to clean the
  original in place. Make a new file from the scrubbed .pptx, then restrict or delete the original
  according to your institution's policy.
- **Comments and suggestions** may not come through the export, but they still live on the original.
- **Sharing**: "anyone with the link" access and share links belong to the original file.
- **Images from Drive or Google Photos** can carry EXIF and GPS; the scrub strips it.
- **Linked charts** (from Sheets) point back to the source spreadsheet, which this tool does not touch.
- **Speaker notes do export.** They are deleted by default.

---

## Limits (please read)

- **This reduces risk. It does not certify de-identification.** You, a human, must review the cleaned
  deck before presenting or sharing it.
- **It cannot see inside pictures.** Faces, tattoos, room numbers, wristbands and burned-in text in
  screenshots, DICOM overlays or photos must be checked by eye. Only file metadata is removed automatically.
- **Audio and video are not inspected.**
- **Names and context are hard.** Patterns catch MRNs, dates, phone numbers and the like. A rare
  diagnosis in a small town, an unusual job, or a date next to a newsworthy event can still identify
  someone. Read the attestation checklist in the report.
- **Embedded objects are removed, not cleaned.** Charts keep their picture and labels (which are scanned).
- **Not tested in every PowerPoint version.** Always open the cleaned file and click through it before
  you present.
- **Running inside Claude sends the file to Claude.** Use an approved environment, or run it offline.

## Disclaimer

Not legal advice. Not a substitute for your institution's privacy office, IRB or de-identification
policy; follow them. Provided as is, without warranty.

## For developers

- Python 3 standard library only (zipfile, xml.etree, re, json, argparse, struct, zlib).
- `tests/run_tests.sh` builds a synthetic deck with obviously fake identifiers, then
  scans, decides, scrubs and verifies it (`python3 -m unittest discover -s tests -v`).
