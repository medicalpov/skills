---
name: phi-scrub-slides
description: Remove PHI (protected health information) from a PowerPoint or Google Slides deck before it is presented or shared. Use when a clinician asks to remove PHI from slides, de-identify a presentation, scrub patient information from a .pptx, make a deck HIPAA-safe, strip EXIF/GPS from clinical photos in slides, or check a Google Slides deck for patient identifiers. Works on .pptx (Google Slides: File > Download > Microsoft PowerPoint). Fail-closed; stdlib Python only.
---

# PHI scrub for PowerPoint / Google Slides

Clinical decks leak PHI in places people do not look: speaker notes, hidden slides, comments, alt text,
the thumbnail, document properties, chart workbooks, and above all **photos** (the lesion photo
also shows a tattoo, the room number, and a GPS tag in the file). This skill finds those places,
proposes a plan, and only writes a new file once **every** finding and image has a decision.

Scripts (Python 3 standard library only, no pip) live in `scripts/` next to this file.
They never modify the input file and never print raw identifiers unless `--show-values` is given.

## Ground rules (tell the user up front)

1. **Data environment.** Running this inside Claude sends the deck to Claude. For real patient data,
   use only an institution-approved, BAA-covered environment, or have the user run the two scripts
   locally with no AI (see README "Run it without AI"). Ask before continuing if unsure.
2. **This reduces risk; it does not certify de-identification.** A human must review the result.
3. **Never present the raw values back in chat** unless the user asks. Use the masked previews.
4. Reports, plans and the image folder contain PHI: keep them local, delete when done.

## Workflow

### 1. Get a .pptx
- PowerPoint: use the file as is (`.pptm` macro decks are refused unless the user agrees to `--strip-macros`).
- Google Slides: **File > Download > Microsoft PowerPoint (.pptx)**. See "Google Slides traps" below.

### 2. Scan
```bash
python3 scripts/scan_pptx.py deck.pptx --out deck.phi-report.json \
    --plan-out deck.plan.json --images-dir deck_review
```
`deck_review/` gets every image named by slide (`slide03_image2.jpeg`) plus `images.md`
(slide, alt text, EXIF flags, GPS, visual checklist).

### 3. Read the inventory yourself and add what regex cannot see
Read `text_items` in the report (every paragraph, table cell, note, comment, alt text, chart label,
layout/master footer, metadata field). Add to `replacements` in the plan:
- **names** of patients and family (all spellings, first-name-only, initials), staff names that
  identify the patient's location, and the deck's own footer text;
- **places** smaller than a state: city, town, county, hospital, clinic, unit, room;
- **context identifiers**: rare diagnosis + small place, occupation, notable event, exact timing
  around a rare event, small-cell counts (n < 11).
Then give **every** `UNRESOLVED` finding a decision: `replace`, or `keep` with a `reason`
(e.g. "presenter's own name"). `removed_by:<action>` entries are already covered by a default action.

### 4. Images (ask the user before viewing them)
Images are the main risk in clinical decks. If the user agrees, view each file in the review folder
and check it against this list:
- **Face or eyes** (full or partial face; eye close-ups).
- **Distinctive marks**: tattoos, birthmarks, scars, piercings, jewelry (Safe Harbor: "full-face
  photographs and any comparable images" and "any other unique identifying characteristic").
- **Location clues**: room or bed numbers, hospital/clinic signs, logos, windows or views, street signs.
- **IDs in frame**: wristbands, ID badges, charts, monitors or screens.
- **Burned-in text**: date stamps, DICOM/EHR overlays, rulers or labels with patient info.

For each image propose one of: `keep` / `replace` (grey placeholder, same size) / `remove`
(or `mask` with pixel `boxes`, only if Pillow happens to be installed). If the image is still needed but
part of it identifies the patient, set `keep` and fill `by_hand` with **where** to crop or cover, e.g.
`"crop or cover by hand: tattoo, lower-left forearm"`. Hidden metadata (EXIF, GPS, camera, owner,
dates, XMP, IPTC, PNG text) is stripped from every kept image by default.
Audio/video cannot be inspected: decide `remove`, or `keep` only after the user has watched and listened.

### 5. Show the user the plan
A short table with masked values: ID, category, where, preview, proposed decision. Then images
(slide, flags, proposed decision, any by-hand instruction). Then the structural defaults below.

### 6. Get explicit approval. Do not scrub without it.

### 7. Scrub
```bash
python3 scripts/scrub_pptx.py deck.pptx --plan deck.plan.json --out deck.deidentified.pptx
```
Fail-closed: it refuses while anything is `UNRESOLVED`, if the plan was made for another file, or on
macros. After writing it re-scans the output and searches every XML part for every original value;
any hit means **no output file** and exit code 2. Do not use `--force` unless the user insists; it
stamps a loud warning into the report.

### 8. Verify (mandatory)
```bash
python3 scripts/scan_pptx.py deck.deidentified.pptx --verify-plan deck.plan.json
```
Must print `PASS`. If not, fix the plan and repeat 7-8.

### 9. Hand back
- the new `.deidentified.pptx` (Google Slides: import as a **new** file);
- the `.scrub-report.md`: what was removed, what was kept and why, the by-hand list, residual risk,
  and the Safe Harbor attestation checklist;
- a short list of what the user must still check by eye: every image, charts, embedded media,
  rare-case details, dates near rare events. Remind them it has not been opened in PowerPoint here.

## Default structural actions (fail-closed)

| Action | Default |
|---|---|
| Speaker notes | deleted |
| Comments + comment authors | deleted |
| Thumbnail (picture of slide 1) | deleted |
| Document properties (author, last modified by, title, company, ...), custom props, revision history | cleared |
| Hidden slides | removed (`"hidden_slides": "keep"` makes their findings need decisions) |
| External hyperlinks | removed |
| Alt text that matched a replacement | cleared |
| Image metadata (EXIF/GPS/XMP/IPTC/PNG text) | stripped |
| Embedded objects (xlsx, docx, OLE) | removed; charts keep their cached picture/labels (scanned) |
| customXml, ppt/tags | removed |
| Macros (.pptm) | refused, or stripped with `--strip-macros` |

## Google Slides traps (say these every time the source is Google Slides)

- **Version history** keeps every past version of the original, PHI included. Never "clean" the
  original in place. Upload the scrubbed .pptx as a **new** file, then restrict or delete the original
  per institutional policy.
- **Comments and suggestions** may not come through the .pptx export, but they still live on the original.
- **Sharing**: share links and "anyone with the link" access belong to the original file.
- **Images from Drive/Photos** may carry EXIF/GPS; the scrub strips it.
- **Linked objects** (charts linked to Sheets) point back to the source data; the source sheet is untouched.
- **Speaker notes do export**, and are deleted by default.
