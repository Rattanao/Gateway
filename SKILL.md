---
name: manifest-enter-edi-check
description: >
  Reconciles an ocean-freight carrier's MANIFEST (.xls/.xlsx, free-form cargo report)
  against a customer's ENTER document (.pdf, per-B/L confirmation/amendment sheet,
  often with hand-typed red annotations for STATUS/DG/REEFER TEMP/TRANSIT — including
  notes that cover a whole range of sub-B/Ls at once) before filing a Thailand customs
  import declaration (ใบขนสินค้าขาเข้า). Auto-detects which of two known ENTER PDF
  layouts is in use. Produces a single color-coded EDI.xlsx report comparing every
  sub-B/L on CONSIGNEE, CONTAINER NO., STATUS, TOTAL PACKAGE, PACKAGING type, GROSS
  WEIGHT, MEASUREMENT, MARKS, DESCRIPTION, REEFER TEMP, DG (CLASS/UN), and
  TRANSIT/TRANSHIPMENT, plus a TOTAL row cross-checking summed values against the
  totals each document states for itself, and a critical row for any B/L that exists
  in ENTER but is entirely missing from MANIFEST (cargo at risk of being left off the
  declaration).
  Use this whenever the user uploads or references files named/shaped like
  MANIFEST.xls, MANIFEST_*.xls, or ENTER.pdf, or asks in Thai or English to
  "ตรวจสอบเอกสารขาเข้าเรือ", "เทียบ MANIFEST กับ ENTER", "ทำ EDI", "ตรวจ B/L",
  or similar import-document reconciliation — even if only one of the two files is
  attached yet, or the wording is loose. This is Rattana's own confirmed report format
  (built and iterated across real jobs for multiple customers/carriers) — trust it
  over improvising a new layout from scratch.
---

# MANIFEST × ENTER → EDI.xlsx reconciliation

This packages a working, already-validated script (`scripts/build_edi.py`) plus the
exact report conventions the user settled on after many rounds of feedback. Don't
re-derive the format — run the script and follow the conventions below.

## When the user gives you files

1. Save the MANIFEST file and the ENTER PDF somewhere sensible (their existing
   `input/` folder if the project already has one, otherwise the folder they
   uploaded to). Any filename containing "MANIFEST"/"ENTER" is fine — the script
   also just picks up the only `.xls`/`.xlsx` and only `.pdf` in the folder if the
   names don't match.
2. **If `EDI.xlsx` already exists at the destination and is open in Excel**, saving
   will raise `PermissionError: [Errno 13]`. Build to a scratch path first, verify,
   then copy over the real path (retry the copy once — Excel sometimes releases the
   lock a moment later); if it's still locked, hand the user the scratch copy and ask
   them to close Excel before you copy it into place.
3. Run:
   ```bash
   python scripts/build_edi.py --manifest <path> --enter <path> --outdir <folder>
   ```
   (or with no args at all if both files sit in the current folder / an `input/`
   subfolder — it auto-detects them). It prints a summary: total B/Ls, how many rows
   need review, and a list of exactly what's wrong per B/L — relay that to the user
   before they open the spreadsheet, leading with anything DG/REEFER/TRANSIT-related
   since those carry legal weight on the declaration.
4. Send the resulting `EDI.xlsx` to the user.

## If the MANIFEST layout doesn't match

`parse_manifest()` in the script assumes one specific free-form carrier layout
(column 0 = B/L number / S:/C:/N: lines, column 5 **or 6** = marks + container number
lines, column 10 = package count + STATUS + description [+ DG/REEFER/TRANSIT tokens],
column 16 = gross weight then measurement — see the header comment in the script and
`references/format-notes.md`). If a new MANIFEST file's columns are laid out
differently, read a raw dump of it first (`xlrd`, print every row) before assuming
the parser will work — adjust the column indices in `parse_manifest()` to match
rather than guessing blind. The B/L prefix itself (e.g. "HASL" for Heung-A) is
auto-detected from the file, so a different carrier with the same layout shape
usually needs no code changes at all.

The marks/container column specifically is **auto-detected** too
(`detect_marks_col()`), not hardcoded to 5 — some MANIFEST exports (confirmed on a real
Heung-A KMTC file) put the actual data one column to the right of where the header
cell's merge range visually suggests (header "MARKS AND NUMBERS" merges starting at
column 7, but the data underneath starts at column 6). Trust the auto-detected column
over the header position if they ever disagree — it's picked from where the container-
number pattern (`\d+\.\s*[A-Z]{4}\d{6,7}`, e.g. `1. ABCD1000002`) actually appears most
often in the data rows, not from where a header label happens to sit.

Also watch for the whole-file grand-total footer (`PORT TOTAL:` / `TOTAL MEASUREMENT:`
/ `TOTAL GROSSWEIGHT:` / `TOTAL PACK:`) leaking into the **last** B/L block's MARKS —
the per-page `PortOfDischarge: ...` footer that normally ends a block sits in column 0
and is caught by `END.match(_cell(rows[j], 0))`, but the true end-of-file grand total on
a real file sat in **column 1** instead (`PORT TOTAL:` at col 1, not col 0). The end-of-
block check scans every column in the row for an `END` match now, not just column 0 —
don't narrow that back down to column 0 only.

## If the ENTER layout doesn't match

`parse_enter()` auto-detects and handles **four** known block layouts (it tries the
literal anchor `B/L NO. :` — with a space before the colon — first, then `B/L NO.:`
without the space, then falls back to the column-position-based layouts C and D if no
anchor is found anywhere in the document). See `references/format-notes.md` for a
full worked example of each:

- **Layout A**: separate `MARK&NO.` / `NO. OF PKGS` / `DESCRIPTION OF GOODS` columns,
  block ends on the anchor `CNTR :` then `B/L NO. :` (space before the colon) then the
  consignee name.
- **Layout B**: no separate MARKS column — one free-text paragraph covers marks and
  description together — block ends on `CONSIGNEE :` / `CONT.NO.:` / `B/L NO.:` (no
  space before this colon) in that order. Because marks and description aren't really
  separable in this layout, `ENT[bl]['combined_text'] = True` is set and the report
  compares the combined text as a **word bag** (`same_bag()`, order-independent)
  instead of the strict character-for-character check used for layout A — flag it
  orange (needs human confirmation) rather than red (confident mismatch) when the
  word sets don't match, since the discrepancy may just be how the two documents
  happened to wrap the same content.
- **Layout C** (`parse_enter_layout_c()`): a real ruled table with no `B/L NO.` anchor
  text anywhere — seen on Cargoport (Thailand) Ltd.'s "AMENDMENT" letter form. Columns
  are identified by x0 position bands instead of anchor text, and a sub-B/L's number
  is split across two lines (base number, then its suffix letter directly below in the
  same column) rather than printed as one string — this is also why the generic
  fallback parser misses 8 of 9 sub-B/Ls on this layout (it can't reassemble a key
  that's never printed contiguously). STATUS and CONTAINER NO. are usually stated once
  for the whole document rather than per block. See `references/format-notes.md` for
  the full column-band map and the reasoning behind each design choice.
- **Layout D** (`parse_enter_layout_d()`): a real ruled table like layout C, seen on AGN
  International Logistics (Thailand) Co., Ltd.'s "AMENDMENT" form — anchor is the
  literal header text `B/L CHANGE NO.` in the leftmost column. Unlike layout C, each
  sub-B/L's full number is printed on **one line** (never split base+suffix), and the
  CONSIGNEE column has a `NOTIFY:` sub-line stacked directly under the consignee name
  in the *same* x0 band — split the column's lines at that label to get `cons`/`notify`
  separately (`rec['notify']` — a new ENT field only this layout populates). The
  `CARGO MOVEMENT n` line sits inside the DESCRIPTION column's x0 band but is **not**
  goods description — filter it out of `desc` (same exclusion category as MANIFEST's
  STATUS token/container-type code) and let the separate CARGO MOVEMENT anchor-scan (see
  below) pick it up on its own. This form's FreeText annotations aren't limited to
  STATUS/DG/REEFER/TRANSIT corrections the way other carriers' are — it also uses one
  for `NOTIFY:` (duplicating the printed NOTIFY line) — `_apply_annot()` must route a
  `NOTIFY:`-prefixed annotation to its own bucket, not the STATUS catch-all, or it
  silently corrupts `status_raw` for that B/L.
  Some AGN files draw the **whole page sideways**: the content is landscape but
  `page.rotation == 0`, and every line has `dir == (0,-1)`. In that case the raw x/y are
  swapped and nothing lands in the right column band. `_page_lines()` detects the majority
  line direction and de-rotates coordinates before layout D reads them. Use it, not raw
  `bbox`, for any new layout-D code. On these files, QUANTITY can also wrap
  (`203` / `CARTONS`), so join all lines in the QTY band. STATUS is printed once in the
  header (`STATUS : *** เร่งเปิดตู้เข้าโกดังด่วน ***`) and applies to every B/L unless a
  FreeText annotation overrides it. Its Thai comes out garbled by the font
  (`เร=งเป@ดตCDเขDาโกดIงด=วน`), and `_LD_TH_FIX` maps it back.
- **Layout F** (`parse_enter_layout_f()`): CMS Logistics "INBOUND CARGO MANIFEST" —
  page 1 is a cover letter (totals + "AS PER ATTACHED SHEET"), then per block: goods
  rows (marks x0<155 / pkgs <255 / desc <480 / KGS.+CBM ≥480), then `CONSIGNEE : <name>`
  with `D/O No. <B/L>` on the same row, then `CNTR NO: <cntr>/<size>`. The B/L is only
  printed after the `D/O No.` label, so annotation anchoring strips that label first
  (otherwise the page has no anchors and FreeText STATUS notes are silently dropped).
  Blocks **split across pages** (goods at the bottom of one page, CONSIGNEE at the top of
  the next) are handled by carrying leftover rows forward. The totals line is
  `TOTAL QUANTITY n PACKAGES G.W. n KGS. MEAS. n CBM.`; a whole-document STATUS note
  like `เปิดตู้เข้าโกดัง-ทุกBL` applies to every B/L — so does any plain FreeText STATUS
  note placed **below** the `TOTAL QUANTITY` line (THE M CORP puts `เปิดตู้เข้าโกดัง` there
  with no "ทุกBL" wording). THE M CORP GLOBAL uses this same layout, printing the label as
  `D / O No .` with spaces, and some of its PDFs use control characters (`\x04`/`\x05`/
  `\x07`/`\x08`) instead of spaces between words. `pymupdf.Page.get_text` is wrapped at
  import time (`_clean_get_text`) to turn those into spaces everywhere — keep it, or
  anchors fail and openpyxl raises `IllegalCharacterError`.
  **Never copy a job folder's local `build_edi.py` over this skill's script** — that has
  already wiped out layout F twice. Merge the changes into this file instead.

Both layouts also read **FreeText PDF annotations** (text the customer typed on top
of the PDF, usually in red) and assign each one to whichever B/L block sits below it
on the page by y-coordinate — this is how hand-added STATUS corrections, DG CLASS/UN,
REEFER temperatures, and INTRANSIT/TRANSHIPMENT notes that aren't in the document's
original printed text get picked up at all. **Never skip these** — per the user's
SOP, annotated text is the customer's real instruction and overrides the printed
form. Watch for **range annotations**: some customers write one note covering a whole
span of sub-B/Ls instead of repeating it per block, e.g.
`HASLK01260000002-T    STATUS :เปิดตู้เข้าโกดัง` meaning the base B/L through its
`-T`-suffixed sibling (682, 682A, 682B, … 682T) all share that STATUS —
`parse_enter()` detects the `<base>-<letter> <instruction>` shape and expands it to
every sub-B/L in the range rather than misfiling the whole note onto whichever single
block happens to sit nearest by position.

If a new ENTER PDF matches neither known layout, dump `page.get_text()` and
`page.annots()` for a page or two first to see the actual anchors and annotation
positions before touching the parser — don't guess blind.

## Report conventions — apply these by default, don't ask

These were fixed through direct user feedback; treat them as settled defaults for
this report type, not something to re-decide each time.

- **Font:** Aptos Narrow, size 9, throughout the whole sheet.
- **Background:** white unless there's a specific reason to color a cell — color
  carries meaning here, so don't add it decoratively. Red (`FDD7D7`) = mismatch
  (prefix the cell text with `⚠`), orange (`FCE4D6`) = needs human confirmation,
  yellow (`FFF2CC`) = no ENTER doc to compare against for that B/L, green (`C6EFCE`)
  = pass (only used in the summary column), navy (`1F4E78`) fill + white bold = header.
- **Column order:** an urgent-summary column comes **first** (column A), then B/L
  NO., then CNEE.(ENTER)/CNEE.(MANIFEST), then **TAX ID vs NOTIFY** right after those
  two (user-requested position — it's a MANIFEST-internal check, not an ENTER/MANIFEST
  pair, so it doesn't fit the alternating pattern but belongs visually next to the
  consignee columns it's most related to), then the remaining ENTER/MANIFEST pairs for
  container, status, package count+type, weight, measurement, marks, description, then
  reefer/DG/transit, then SHED NO./ประเทศปลายทาง/CARGO MOVEMENT, then a detail-notes
  column **last**. See `HEAD` in the script for the exact list — don't reorder it
  without the user asking. The four check functions and all the per-B/L comparison
  logic in `build_report()` are written against a fixed *internal* column numbering
  (`mk()` calls, `cs[...]`, `COL_SHED`/`COL_DESTCODE`/`COL_MOVEMENT`/`COL_TAXID`/
  `COL_NOTE`) that does **not** match `HEAD`'s visual order — `_OLD_TO_NEW` /
  `_reorder()` / `_remap_cols()` translate from internal numbering to the real sheet
  position exactly once, right before each of the four places a row gets written (the
  per-B/L row, the "no ENTER" row, the "B/L missing from MANIFEST" row, and the TOTAL
  row). If the user asks to move a column again, change `HEAD` and `_OLD_TO_NEW` only —
  don't renumber the `mk()` calls scattered through the comparison logic.
- **CNEE. and NOTIFY should be the same company** (per the user's SOP) — TAX ID vs
  NOTIFY (`check_taxid()`) enforces this against MANIFEST's own N:/NOTIFY PARTY field
  by default. The one exception is a B/L with a TRANSIT/TRANSHIPMENT sentence: there,
  the correct notify party may be one ENTER states separately (a forwarder handling the
  in-transit leg) rather than MANIFEST's own N: — `compute_extra_checks()` passes
  ENTER's own `rec['notify']` (only layout D populates this field so far) as a second
  acceptable match, but **only** when the B/L has transit text; a non-transit B/L must
  still match MANIFEST's N: exactly.
- **STATUS canon:** CY, CY/CY, FCL, ลากตู้ → `CY`; LCL, ขน(ส่ง) → `LCL`; CFS,
  LCL/CFS, เปิดตู้ → `LCL/CFS`. Show the raw text plus `(= <canonical>)`.
  If the customer's ENTER document uses different colloquial terms for the same
  three buckets, extend `canon_status()` rather than leaving them unclassified.
- **PACKAGING:** reduce to the bare noun (CASE/PALLET/CARTON/BOX/PACKAGE/SET/…) —
  never leave a raw manifest code, a package count, or a `(s)` suffix in this column.
- **MARKS / DESCRIPTION:** compare strictly — normalize by stripping everything but
  letters and digits, then require exact equality (`same_detail()`). This tolerates
  only the line-wrapping/spacing differences that free-form MANIFESTs introduce (a
  stray mid-word space from PDF/Excel extraction never counts as a mismatch); any
  real wording difference (an extra/missing word or sentence) must still be flagged,
  even a small one. Compare MARKS(ENTER) vs MARKS(MANIFEST) and DESC(ENTER) vs
  DESC(MANIFEST) **as separate pairs**, not as one combined blob — flagging a column
  that is byte-identical to its counterpart just because the *other* column in the
  same row disagrees is confusing to the reviewer. For layout-B ENTER files (real
  x0-split MARKS/DESCRIPTION columns, see format-notes.md) flag orange, not red,
  since the MANIFEST's own split point comes from Excel line-wrapping and can
  legitimately land the same content on the other side of the line from ENTER's real
  column boundary — flag **both documents'** cells for whichever field actually
  differs, not just the MANIFEST side, so a reviewer scanning only the ENTER column
  still sees that something needs checking. **Never truncate or drop content** — the DESCRIPTION (MANIFEST) cell must
  read exactly like the source file's "NO. OF PKGS / DESCRIPTION OF GOODS" column,
  including any DG/REEFER/TRANSIT sentence embedded in it (those get *also* pulled
  into their own dedicated columns for quick scanning, but must stay in DESCRIPTION
  too — extracting them is additive, never destructive). The only content that's
  legitimately excluded from DESCRIPTION is pure structural noise that isn't
  "description of goods" in any reading and already has its own column elsewhere:
  the STATUS token and the container-type code (e.g. `45G0X1 PART CNTR`).
- **DG / REEFER TEMP / TRANSIT-TRANSHIPMENT:** flag red on any B/L where these show
  up **at all**, in either document — not just when the two documents disagree with
  each other. These always need customs follow-up regardless of agreement. The
  TRANSIT/TRANSHIPMENT **column** shows only the distilled phrase via
  `extract_transit_phrase()` — `"INTRANSIT TO <destination>"` or
  `"TRANSHIPMENT TO <destination>"` — never the raw sentence or an `ENTER:`/
  `MANIFEST:` source label; some MANIFESTs write a full travel narrative
  (`"IN TRANSIT CARGO FROM HOCHIMINH PORT, VIETNAM VIA LAEM CHABANG PORT BY SEA TO
  MYWADDY, MYANMAR BY TRUCK"`) and only the actual destination after the last `TO`
  (before a `VIA`/`BY`/`FROM` clause or the sentence end) belongs on-screen — the full
  sentence still belongs in DESCRIPTION per the rule above, just not repeated in this
  column. When neither document mentions transit/transhipment for a B/L, show `-`
  in this column (not `LOCAL`).
- **B/L in ENTER but missing from MANIFEST entirely:** its own critical red row
  (`⚠⚠ ... วิกฤต`), not just a console message — this is cargo that risks being left
  off the customs declaration altogether, distinct from (and worse than) a B/L that's
  in MANIFEST with no ENTER to compare against (which gets a yellow row instead).
- **TOTAL row:** sums every numeric column for both ENTER and MANIFEST (right-align
  those columns, comma-separated thousands), a packaging-kind breakdown, counts of
  STATUS/DG/REEFER/TRANSIT flags, and cross-checks against whatever total each
  source document states about itself (MANIFEST's own "GRAND TOTAL" footer, ENTER's
  own "PACKAGES/KGS/CBM" summary) — not just against the other document. Distinct
  containers count once each, not once per B/L sharing them.
- Row 2 = `Vessel <name> | Master B/L <number>` in plain black bold, one line — both
  values come from the parsed data, don't hardcode a shipment's vessel or B/L number.
  Row 3 = the color/legend key, bold red font.
- Row heights are fixed (not auto-fit) and header/data columns are sized to their
  own header text by default, except MARKS/DESC. which get extra width for full text.
- **Column alignment:** every data column is left-aligned, except G.W. KGM (ENTER),
  G.W. KGM (MANIFEST), MEAS. MTQ (ENTER), and MEAS. MTQ (MANIFEST), which are
  right-aligned. Header row (row 4) stays centered. Don't right-align TOTAL PKG/PKG.
  or any other column just because it holds a number.

If the user asks for a formatting tweak that isn't covered above (a column width, a
color, a reordering), make the change and treat it as durable for *this* report going
forward, the same way the conventions above were built up — update this SKILL.md's
list if the change seems likely to matter for the next job too, not just this one.

## 4 extra fixed-rule checks (SHED NO. / ประเทศปลายทาง / CARGO MOVEMENT / TAX ID)

**Display rule (user, 2026-10-09):** in all 4 columns, a value that passes its locked rule
shows `-`, not the value itself. The actual MANIFEST value is shown only when it's wrong
(red) or needs a person to check it (orange). `-` therefore means either "correct" or "no rule
applies".

Added per the user's fixed customs rules (not a MANIFEST↔ENTER comparison — these check
the MANIFEST's own printed values against fixed reference tables). They live in their
own 4 columns at the end of the report (before the notes column) and are computed by
`compute_extra_checks()` / `check_shed()` / `check_dest_code()` / `check_movement()` /
`check_taxid()` in `build_edi.py`. All 4 values are printed on **one column (col 15) of
MANIFEST.xls**, as 4 consecutive lines per B/L block: `1) SHED NO. ####` /
`2) ประเทศปลายทาง (XX)` / `3) CARGO MOVEMENT (X)` / `4) TAX ID: (ชื่อ)` — parsed into
`rec['shed_no']` / `rec['dest_code_raw']` / `rec['movement_raw']` / `rec['tax_id_raw']`
in `parse_manifest()`.

- **1) SHED NO. — user-specified SHED (takes priority):** the user states the job's SHED
  number when asking for the check (e.g. "ตรวจสอบเอกสาร SHED NO. 0141"). Pass it as
  `--shed 0141`. Then **every** B/L must match that number: a match shows `-`, a mismatch
  is red. The one exception is LAOS, which keeps the old rule (0124 when discharge =
  BANGKOK). If the user didn't state a SHED, ask for it, or run without `--shed` to fall
  back to the 7-category rules below.
- **1) SHED NO. (fallback without --shed)** — only 7 categories have a required SHED NO.; everything else shows
  `-` (no rule). The category is decided by the **PORT OF DISCHARGE** field (tracked
  per-page in `parse_manifest()` as `rec['port_discharge']`) plus, for two of the
  categories, other MANIFEST content:
  - `USED ENGINE` keyword anywhere in DESCRIPTION, **and** discharge = BANGKOK → `0126`
  - destination country = LAOS (from the TRANSIT sentence), **and** discharge = BANGKOK
    → `0124`
  - DG detected on the B/L, **and** discharge = LAEM CHABANG → `2826`
  - discharge field itself reads `THLKR` → `0332` ; `BMT` → `0110` ; `SCT` → `0302` ;
    `UNITHAI` → `0113`. **No real example of these 4 has been seen yet** (confirmed by
    the user: they show up as the PORT OF DISCHARGE value itself, not a separate
    "Place of Delivery" field) — the code is ready for them but unverified against a
    real file. If one of these ever fires, double-check the match against the actual
    cell text before trusting it blindly.
- **2) ประเทศปลายทาง (XX)** — only checked when the B/L has a TRANSIT/TRANSSHIPMENT
  sentence with a recognized destination country (`DEST_CODE` table: LAOS→LA (user-confirmed 2026-10-12; was LL),
  MYANMAR→MM, CHINA→CN, MALAYSIA→MY, PHILIPPINES→PH, CAMBODIA→KH, MARSHALL ISLAND→MH,
  INDIA→IN). No transit sentence, or an unrecognized country → `-`, not checked.
- **3) CARGO MOVEMENT** — TRANSIT to anywhere except Laos must read `(7-TRANSIT)`.
  TRANSIT to Laos is the one exception: MANIFEST's code (`7-TRANSIT` or
  `L-INTRANSIT TO LAOS`) must match whatever ENTER's own `CARGO MOVEMENT <n>` line says
  for that same B/L (plain printed text on the ENTER PDF, not a FreeText annotation —
  parsed in `parse_enter()` into `ent_movement`/`rec['movement_raw']`, matched to the
  nearest B/L anchor **above** it — the opposite direction from FreeText annotations,
  since this text sits below its own block's B/L-number line, not above the next
  block's). If ENTER has no movement value for that B/L, flag **orange** ("ต้องเทียบ
  ENTER ด้วยคน"), not red — a real gap, not a confident mismatch. No transit sentence
  at all → `-` (STATUS `O-OTHER`/`5-CONSOLIDATE` cases are correct by definition).
- **4) TAX ID: (ชื่อ)** — must company-match (`comatch()`, same fuzzy logic as
  CONSIGNEE) against the B/L's own `N :`/NOTIFY PARTY name. Empty `TAX ID: ()` → `-`,
  not checked (common in real files).

## Web version: Check EDI Quality (GitHub Pages)

The Quality ENTER form (layout B — `MARKS & NOS | DESCRIPTIONS OF GOODS | GROSS WT.& MEAS`,
anchors `CONSIGNEE :` / `CONT.NO.:` / `B/L NO.:`) also exists as an in-browser page so
colleagues can use it without installing Python: repo `Rattanao/check-edi-Quality`,
live at https://rattanao.github.io/check-edi-Quality/ (Settings → Pages → Deploy from a
branch → `main` / root; the root `index.html` redirects to `docs/`).

- **Limitation (state it on the page and to the user): the web page supports only
  Quality's AMENDMENT form.** `isQualityAmendment()` rejects anything else (including the
  AGN form, which has `B/L CHANGE NO.`) with a message instead of guessing. `CNTRS*`
  files are ignored (not compared). For other customers' ENTER forms use the Python
  script / this skill, or build a sibling page the same way (`check-edi-AGN` is the
  template: `docs/index.html` + `docs/core.js`, same UI, only the ENTER parser differs).
- `docs/core.js` is a JS port of `build_edi.py` (parseManifest + report logic shared with
  the AGN page; `parseEnter` rewritten for the Quality form). pdf.js gives word-level
  items `{x,y,w,t}`, so the port groups items into rows, splits each row into segments at
  gaps > 20pt, and classifies by `rel = x − page label x` (the x of the `CONSIGNEE`/
  `CONT.NO.`/`B/L NO.` labels on that page — page margins differ between files, e.g.
  20.6pt vs 56.6pt, so never hard-code absolute x): `rel < 100` → MARKS, `rel ≥ 100` →
  DESCRIPTION/values, `rel ≥ 300` → KGS/CBM column. Classify by **segment**, not by single
  word — long MARKS lines legitimately contain words far to the right.
- Same report rules as the Python script: MARKS/DESC compared per pair with
  `sameDetail()` (spaces ignored), mismatch → orange on both documents' cells; TRANSIT
  column `-` when absent; consignee joined across wrapped lines; blocks may straddle a
  page break (rows are indexed globally across pages).
- **Verify any change to the port against Python:** run `build_edi.py`, run the JS port in
  node (`pdfjs-dist@3.11.174`, `xlsx@0.18.5`, same versions as the CDN tags), and diff
  every cell — the 2026-10-09 Quality job (9 B/L × 28 columns) matched with 0 differences.
- Never commit customer files: `.gitignore` excludes `input/`, `output/`, `*.xls*`, `*.pdf`.

## Reference

- `scripts/build_edi.py` — the full implementation; read its docstring and inline
  comments before modifying anything, they document every column-index assumption.
- `references/format-notes.md` — worked examples of the exact MANIFEST row shape and
  **both** known ENTER layouts (A and B) this parser expects, plus the B/L-range
  annotation pattern, taken from real validated jobs across different customers.
