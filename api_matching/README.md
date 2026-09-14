# Nexus header-file API repair

Rebuilds the `API` column of the Nexus well header file from an Enverus well
export, by matching each row to its Enverus counterpart on well name.

## The problem

The `API` column as received is bucketed, not merely imprecise. 320 populated
cells collapse to **38 distinct values**, every one a multiple of 100:

| Value in header file | True APIs that fall in it | Count | Span |
|---|---|---|---|
| `3505124600` | `3505124509` – `3505124735` | 55 | 226 |
| `3505125000` | `3505124993` – `3505125218` | 29 | 225 |
| `3505124000` | `3505123978` – `3505124200` | 19 | 222 |

The state and county prefix (first 5 digits) survives intact — it agreed with
the `County` column in **301 of 301** unambiguous matches. Everything after it
is a bucket label, so well identity has to be recovered by name.

## Matching

Candidates are always restricted to the row's county. Tiers, strictest first:

1. **`exact_name`** — exact match on the normalized well name, unique in the
   county.
2. **`exact_name_tiebreak`** — several wells share the name; resolved by which
   candidate agrees on more independent fields.
3. **`fuzzy_corroborated`** — similar name, accepted only when independent
   fields agree: at least one corroborating field above 0.90 similarity, at
   least two between 0.75 and 0.90. Below 0.75 nothing is applied.

Two normalizations do most of the work:

- **Suffix canonicalization.** Operators spell the horizontal designator
  inconsistently, so letter runs inside a mixed alphanumeric token are sorted:
  `WILD BILL 6-31-6-7HXW` and `...-7XHW` reduce to the same key. Pure-word
  tokens are left alone.
- **Lease root + Well Number.** The `Lease` and `Well Number` columns disagree
  often enough to matter — `LEONA FAYE 1-20-29-32XHS` carries well number
  `1-20HS`, and Enverus calls the well `LEONA FAYE 1-20HS`. Recombining the
  lease root with the well number recovers the vendor's name and resolved 7
  rows that no name-only match reaches.

**Corroborating fields** are spud date, first production date, lateral length
(graded by relative difference) and operator. Well number, field name and
proximity to the original bucketed API break ties but never justify a match on
their own.

Corroboration outranks name similarity in the ranking, which is what makes the
hard rows come out right. `LEONA FAYE 1-20-29-32XHS` scores 0.96 against
`LEONA FAYE 2-20-29-32XHS` — a different well — and only 0.83 against the
correct `LEONA FAYE 1-20HS`, which agrees on spud, first production, lateral
and well number. Ranking on similarity alone picks the wrong well.

## Undrilled locations

Rows with a blank API are PUD locations. They get tier 1 only — no fuzzy
matching. An undrilled location has no spud date, its first-production date is
a forecast, and its lateral length is a planned round number, so nothing can
corroborate a name; the nearest name is nearly always a sibling well on the
same pad. Where a near name exists it is recorded in the notes as a lead and
the cell is left blank.

## Output

- Corrected `API` in place, highlighted where changed. Every other cell of the
  original sheet is left byte-identical.
- Audit columns appended to the right: API Source, Original API (as received),
  Matched Enverus Well Name, Match Method, Match Confidence, Corroborating
  Fields, Match Notes.
- A QC workbook: Summary, Match Detail (every row, with runner-up candidates
  and API drift), and a Review Queue of everything flagged plus every undrilled
  location where a near name exists and someone who knows the asset could
  confirm or reject it.

The Summary carries both a live `COUNTIF` over the detail sheet and the count
the run produced, in adjacent columns. openpyxl writes formulas with no cached
value, so the live column reads blank until a spreadsheet application opens the
file; the generated column keeps the numbers readable anywhere, and the two
cross-check each other.

## Running it

```sh
python3 match_apis.py \
  --header  "HeaderData_-_Nexus_14Sep2026.xlsx" \
  --enverus "env_csv-Wells-22c45_2026-09-14.csv" \
  --out     "HeaderData_-_Nexus_14Sep2026_API-corrected.xlsx" \
  --report  "API_Match_QC_Report.xlsx"
```

Requires `openpyxl` and `pandas`. The Enverus export is read with a fixed
column subset; see `build_env.py`.

## Result on the 14 Sep 2026 files

| | Rows |
|---|---|
| API corrected | 320 |
| API already correct | 1 |
| Undrilled location, left blank | 190 |
| Flagged for review | 0 |

Every accepted match agreed on **at least two** independent fields; 233 of 321
agreed on all five. No two rows resolved to the same API. Every corrected API
landed within its original bucket (drift −174 to +218, none beyond ±250) and
kept the original's county code.
