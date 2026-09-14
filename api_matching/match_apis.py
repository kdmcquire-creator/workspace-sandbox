#!/usr/bin/env python3
"""Repair the API numbers in the Nexus header file from an Enverus well export.

The header file's API column is bucketed: 320 populated cells collapse to only
38 distinct values, because the trailing 5-digit unique well number has been
rounded to a shared 100-multiple. The state+county prefix (first 5 digits)
survives intact, so the well identity has to be recovered by name.

Matching runs in tiers, strictest first:

  T1  exact normalized (or suffix-canonical) well name within county, unique
  T2  same, several candidates -- resolved by corroborating field agreement
  T3  fuzzy name within county, accepted only when independent fields agree

Rows with no original API are undrilled PUD locations. They get T1 only: an
undrilled location has no spud date, no real first-production date and only a
planned lateral length, so nothing can corroborate a fuzzy name and the
nearest name is almost always a different well on the same pad.
"""

import argparse
import collections
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from build_env import load_enverus
from wellmatch import (
    canon_name, digits, iso_date, name_similarity, norm_county, norm_name,
    to_float,
)

SHEET = "Header Data"

# Corroboration weights. Only the four the brief names as independent fields
# count toward `corr`; the rest break ties without ever justifying a match.
W_SPUD = 3.0
W_FIRST_PROD = 2.5
W_LATERAL = {0.005: 2.5, 0.02: 1.5, 0.05: 0.5}
W_WELLNO = 2.0
W_OPERATOR = 1.0
W_FIELD = 0.5
W_API_NEAR = {250: 1.0, 400: 0.5}

# Empirical: across 301 unambiguous exact-name matches the bucketed API sat
# within -174..+218 of the true API, and the county prefix agreed every time.
API_DRIFT_MAX = 400

SIM_STRONG = 0.90   # above this, one corroborating field is enough
SIM_FLOOR = 0.75    # below this, never auto-apply

OPERATOR_NOISE = {
    "LLC", "LLP", "LP", "INC", "INCORPORATED", "CORP", "CORPORATION", "CO",
    "COMPANY", "COMPANIES", "OIL", "GAS", "ENERGY", "ENERGIES", "RESOURCES",
    "RESOURCE", "PETROLEUM", "OPERATING", "OPERATOR", "PRODUCTION",
    "PRODUCING", "EXPLORATION", "NATURAL", "USA", "US", "AMERICA", "PARTNERS",
    "HOLDINGS", "MIDSTREAM", "GROUP", "THE", "AND", "OF", "E", "P",
}


def op_tokens(value):
    toks = [t for t in norm_name(value).split() if t not in OPERATOR_NOISE]
    return toks or norm_name(value).split()


def operator_match(a, b):
    ta, tb = op_tokens(a), op_tokens(b)
    if not ta or not tb:
        return False
    return ta[0] == tb[0] or set(ta) <= set(tb) or set(tb) <= set(ta)


def lateral_weight(a, b):
    fa, fb = to_float(a), to_float(b)
    if not fa or not fb:
        return 0.0
    rel = abs(fa - fb) / max(fa, fb)
    for tol in sorted(W_LATERAL):
        if rel <= tol:
            return W_LATERAL[tol]
    return 0.0


def api_weight(original, candidate):
    """Reward candidates near the bucketed original; require the county code."""
    if not original or original[:5] != candidate[:5]:
        return 0.0
    delta = abs(int(candidate) - int(original))
    for tol in sorted(W_API_NEAR):
        if delta <= tol:
            return W_API_NEAR[tol]
    return 0.0


class Row:
    """One header-file row, with the fields matching depends on pre-normalized."""

    __slots__ = ("excel_row", "prop_id", "api", "rsv_cat", "lease", "well_no",
                 "field", "operator", "county", "spud", "first_prod", "lateral",
                 "name_n", "name_c", "wellno_n", "alt_names")

    def __init__(self, excel_row, values, col):
        def get(name):
            v = values[col[name]]
            return None if v is None or str(v).strip() == "" else str(v).strip()

        self.excel_row = excel_row
        self.prop_id = get("Prop_ID")
        self.api = digits(get("API")) or None
        self.rsv_cat = get("RsvCat")
        self.lease = get("Lease")
        self.well_no = get("Well Number")
        self.field = get("Field")
        self.operator = get("Operator")
        self.county = norm_county(get("County"))
        self.spud = iso_date(get("Spud"))
        self.first_prod = iso_date(get("First Prod"))
        self.lateral = to_float(get("Lateral Length"))
        self.name_n = norm_name(self.lease)
        self.name_c = canon_name(self.lease)
        self.wellno_n = norm_name(self.well_no)
        self.alt_names = self._alt_names()

    def _alt_names(self):
        """Lease name rebuilt around the Well Number column.

        The two columns disagree often enough to matter -- 'LEONA FAYE
        1-20-29-32XHS' carries well number '1-20HS', and Enverus calls it
        'LEONA FAYE 1-20HS'. Pairing the lease root with the well number
        recovers the name the vendor actually used.
        """
        if not self.wellno_n or not self.name_n:
            return []
        toks = self.name_n.split()
        wtoks = self.wellno_n.split()
        if toks[-len(wtoks):] == wtoks:
            root = toks[:-len(wtoks)]
        else:
            root = [t for t in toks if not any(c.isdigit() for c in t)]
        if not root:
            return []
        alt = " ".join(root + wtoks)
        return [] if alt == self.name_n else [alt]


def score(row, cand):
    """Return (weight, corroborations, similarity) for one candidate well."""
    sims = [name_similarity(row.name_c, cand.name_c),
            name_similarity(row.name_n, cand.name_n)]
    for alt in row.alt_names:
        sims.append(name_similarity(canon_name(alt), cand.name_c))
    sim = max(sims)

    corr = 0
    weight = 0.0
    if row.spud and cand.SpudDate and row.spud == cand.SpudDate:
        weight += W_SPUD
        corr += 1
    if row.first_prod and cand.FirstProdDate and row.first_prod == cand.FirstProdDate:
        weight += W_FIRST_PROD
        corr += 1
    lat = lateral_weight(row.lateral, cand.LateralLength_FT)
    if lat:
        weight += lat
        corr += 1
    if row.operator and cand.ENVOperator and operator_match(row.operator, cand.ENVOperator):
        weight += W_OPERATOR
        corr += 1

    if row.wellno_n and cand.wellno_n and row.wellno_n == cand.wellno_n:
        weight += W_WELLNO
    if row.field and cand.Field and norm_name(row.field) == norm_name(cand.Field):
        weight += W_FIELD
    weight += api_weight(row.api, cand.Unformatted_API_UWI)

    return weight, corr, sim


def narrow(row, pool, token_idx):
    """Candidates sharing a name token with the header row, or its well number.

    Scanning every well in a county adds thousands of unrelated names without
    ever changing the winner.
    """
    tokens = {t for key in {row.name_n, row.name_c, *row.alt_names}
              for t in key.split()}
    found = {}
    for tok in tokens:
        for cand in token_idx.get((row.county, tok), []):
            found[cand.Unformatted_API_UWI] = cand
    if row.wellno_n:
        for cand in pool:
            if cand.wellno_n == row.wellno_n:
                found[cand.Unformatted_API_UWI] = cand
    return list(found.values())


Candidate = collections.namedtuple(
    "Candidate", "api weight corr sim name")


def rank(row, wells):
    ranked = []
    for cand in wells:
        weight, corr, sim = score(row, cand)
        ranked.append(Candidate(cand.Unformatted_API_UWI, weight, corr, sim,
                                cand.WellName))
    ranked.sort(key=lambda c: (c.weight, c.sim), reverse=True)
    return ranked


def accepts_fuzzy(cand):
    if cand.sim >= SIM_STRONG:
        return cand.corr >= 1
    if cand.sim >= SIM_FLOOR:
        return cand.corr >= 2
    return False


def match_row(row, by_county, exact_idx, token_idx):
    """Return (api, method, confidence, note, candidates_considered)."""
    if not row.county:
        return None, "no_match", "review", "county missing from header row", []

    hits = sorted(set(exact_idx.get((row.county, row.name_n), [])
                      + exact_idx.get((row.county, row.name_c), [])))
    for alt in row.alt_names:
        hits = sorted(set(hits + exact_idx.get((row.county, canon_name(alt)), [])))

    pool = by_county.get(row.county, [])
    lookup = {c.Unformatted_API_UWI: c for c in pool}

    if len(hits) == 1:
        cand = rank(row, [lookup[hits[0]]])[0]
        env_keys = {norm_name(cand.name), canon_name(cand.name)}
        via = ("Lease" if env_keys & {row.name_n, row.name_c}
               else "Lease root + Well Number")
        note = "exact well name in county, unique (matched on %s)" % via
        if row.api and cand.api[:5] != row.api[:5]:
            return None, "no_match", "review", (
                "exact name match %s but county code disagrees with original API"
                % cand.api), [cand]
        if cand.corr == 0:
            return cand, "exact_name", "medium", (
                note + "; no other field corroborates"), [cand]
        return cand, "exact_name", "high", note, [cand]

    if not row.api:
        # Undrilled location: no field can corroborate, so exact name or nothing.
        # Planned laterals are round numbers and first-production dates are
        # forecasts, so the nearest name is nearly always a sibling well on the
        # same pad. Surface it as a lead, never as a value.
        if hits:
            return None, "no_match", "review", (
                "undrilled location; %d wells share this name" % len(hits)), []
        ranked = rank(row, narrow(row, pool, token_idx))
        near = ""
        if ranked and ranked[0].sim >= SIM_FLOOR:
            near = ("; nearest name in county is %s %s (similarity %.2f)"
                    " -- NOT applied, most likely a different well on the"
                    " same pad" % (ranked[0].api, ranked[0].name, ranked[0].sim))
        return None, "no_match", "no_api_expected", (
            "undrilled location; no exact name match in county" + near), ranked[:4]

    if len(hits) > 1:
        ranked = rank(row, [lookup[h] for h in hits])
        best, runner = ranked[0], ranked[1]
        if best.weight > runner.weight and best.corr >= 1:
            return best, "exact_name_tiebreak", "high", (
                "%d wells share this name; chosen on %d corroborating field(s)"
                % (len(hits), best.corr)), ranked
        return None, "no_match", "review", (
            "%d wells share this name and no field separates them" % len(hits)
        ), ranked

    # The county code in the bucketed original API held in all 301 unambiguous
    # matches, so a candidate outside it is not this well.
    narrowed = [c for c in narrow(row, pool, token_idx)
                if c.Unformatted_API_UWI[:5] == row.api[:5]]
    if not narrowed:
        return None, "no_match", "review", "no candidate in county", []

    ranked = rank(row, narrowed)
    best = ranked[0]
    runner = ranked[1] if len(ranked) > 1 else None
    if not accepts_fuzzy(best):
        return None, "no_match", "review", (
            "best candidate %s (%s) scored %.2f similarity with %d corroborating "
            "field(s) -- below the acceptance bar"
            % (best.api, best.name, best.sim, best.corr)), ranked
    if runner and runner.weight >= best.weight:
        return None, "no_match", "review", (
            "candidates %s and %s are indistinguishable" % (best.api, runner.api)
        ), ranked
    if abs(int(best.api) - int(row.api)) > API_DRIFT_MAX:
        return None, "no_match", "review", (
            "best candidate %s sits %d beyond the original API's bucket"
            % (best.api, abs(int(best.api) - int(row.api)))), ranked
    return best, "fuzzy_corroborated", ("high" if best.corr >= 3 else "medium"), (
        "name similarity %.2f with %d corroborating field(s)"
        % (best.sim, best.corr)), ranked


AUDIT_HEADERS = [
    "API Source", "Original API (as received)", "Matched Enverus Well Name",
    "Match Method", "Match Confidence", "Corroborating Fields", "Match Notes",
]


def corroborating_fields(row, cand, lookup):
    if not cand:
        return ""
    well = lookup[cand.api]
    hit = []
    if row.spud and well.SpudDate and row.spud == well.SpudDate:
        hit.append("Spud")
    if row.first_prod and well.FirstProdDate and row.first_prod == well.FirstProdDate:
        hit.append("First Prod")
    if lateral_weight(row.lateral, well.LateralLength_FT):
        hit.append("Lateral Length")
    if row.operator and well.ENVOperator and operator_match(row.operator, well.ENVOperator):
        hit.append("Operator")
    if row.wellno_n and well.wellno_n and row.wellno_n == well.wellno_n:
        hit.append("Well Number")
    return ", ".join(hit)


SUMMARY_ROWS = [
    ("Rows in header file", "COUNTA"),
    ("", None),
    ("API corrected from Enverus", 'Status=corrected'),
    ("API already correct", 'Status=unchanged'),
    ("Undrilled location - left blank", 'Status=no_api_expected'),
    ("Flagged for review - original retained", 'Status=flagged_kept_original'),
    ("Flagged for review - still blank", 'Status=flagged_blank'),
    ("", None),
    ("Matched on exact well name", 'Match Method=exact_name'),
    ("Matched on exact name, tie broken by data", 'Match Method=exact_name_tiebreak'),
    ("Matched on corroborated fuzzy name", 'Match Method=fuzzy_corroborated'),
    ("", None),
    ("High confidence", 'Confidence=high'),
    ("Medium confidence", 'Confidence=medium'),
]


def write_report(path, report, header_file, enverus_file):
    """QC workbook: a summary that recalculates, the full detail, a review queue."""
    fields = list(report[0])
    head_font = Font(name="Calibri", size=11, bold=True)
    body_font = Font(name="Calibri", size=11)
    head_fill = PatternFill("solid", fgColor="D9E1F2")

    wb = openpyxl.Workbook()

    def sheet(title, rows):
        ws = wb.create_sheet(title)
        for i, name in enumerate(fields, start=1):
            cell = ws.cell(row=1, column=i, value=name)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="left", vertical="center")
        for r, record in enumerate(rows, start=2):
            for i, name in enumerate(fields, start=1):
                cell = ws.cell(row=r, column=i, value=record[name])
                cell.font = body_font
                cell.alignment = Alignment(horizontal="left", vertical="top")
        for i, name in enumerate(fields, start=1):
            width = max(len(name), *(len(str(rec[name] or "")) for rec in rows)) \
                if rows else len(name)
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = \
                max(10, min(60, width + 2))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        return ws

    sheet("Match Detail", report)
    review = [r for r in report if not r["Corrected API"]
              and r["Status"] != "no_api_expected"]
    unresolved = [r for r in report
                  if not r["Corrected API"] and "nearest name" in (r["Notes"] or "")]
    sheet("Review Queue", review or unresolved)

    ws = wb["Sheet"]
    ws.title = "Summary"
    last = len(report) + 1
    col_of = {name: openpyxl.utils.get_column_letter(i)
              for i, name in enumerate(fields, start=1)}

    ws["A1"] = "API correction QC summary"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True)
    for r, (label, value) in enumerate(
            [("Header file", header_file), ("Enverus export", enverus_file)], start=3):
        ws.cell(row=r, column=1, value=label).font = body_font
        ws.cell(row=r, column=2, value=value).font = body_font

    row = 6
    for i, title in enumerate(("", "Live count", "Count when generated"), start=1):
        cell = ws.cell(row=row, column=i, value=title)
        cell.font = head_font
        cell.fill = head_fill if title else PatternFill()
    ws.cell(row=row, column=2).alignment = Alignment(horizontal="right")
    ws.cell(row=row, column=3).alignment = Alignment(horizontal="right")

    row = 7
    for label, rule in SUMMARY_ROWS:
        if not label:
            row += 1
            continue
        ws.cell(row=row, column=1, value=label).font = body_font
        if rule == "COUNTA":
            formula = "=COUNTA('Match Detail'!%s2:%s%d)" % (
                col_of["Prop_ID"], col_of["Prop_ID"], last)
            tally = sum(1 for rec in report if rec["Prop_ID"])
        else:
            field, value = rule.split("=", 1)
            formula = '=COUNTIF(\'Match Detail\'!%s2:%s%d,"%s")' % (
                col_of[field], col_of[field], last, value)
            tally = sum(1 for rec in report if rec[field] == value)
        for column, content in ((2, formula), (3, tally)):
            cell = ws.cell(row=row, column=column, value=content)
            cell.font = body_font
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right")
        row += 1

    ws.cell(row=row + 1, column=1,
            value="'Live count' recalculates against the Match Detail sheet. "
                  "'Count when generated' is what this run produced -- the two "
                  "should agree unless Match Detail has been edited."
            ).font = Font(name="Calibri", size=9, italic=True)

    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 22
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--header", required=True, help="Nexus header .xlsx")
    ap.add_argument("--enverus", required=True, help="Enverus wells .csv")
    ap.add_argument("--out", required=True, help="updated .xlsx to write")
    ap.add_argument("--report", required=True, help="QC .xlsx to write")
    args = ap.parse_args()

    wells = load_enverus(args.enverus)
    by_county = collections.defaultdict(list)
    exact_idx = collections.defaultdict(list)
    token_idx = collections.defaultdict(list)
    for cand in wells.itertuples():
        by_county[cand.county_n].append(cand)
        exact_idx[(cand.county_n, cand.name_n)].append(cand.Unformatted_API_UWI)
        if cand.name_c != cand.name_n:
            exact_idx[(cand.county_n, cand.name_c)].append(cand.Unformatted_API_UWI)
        for tok in set(cand.name_n.split()) | set(cand.name_c.split()):
            token_idx[(cand.county_n, tok)].append(cand)
    lookup = {c.Unformatted_API_UWI: c for c in wells.itertuples()}

    wb = openpyxl.load_workbook(args.header)
    ws = wb[SHEET]
    headers = [c.value for c in ws[1]]
    col = {name: i for i, name in enumerate(headers)}
    api_col = col["API"] + 1

    first_audit = ws.max_column + 1
    head_font = Font(name="Calibri", size=11, bold=True)
    body_font = Font(name="Calibri", size=11)
    changed_fill = PatternFill("solid", fgColor="FFF2CC")
    review_fill = PatternFill("solid", fgColor="FCE4D6")

    for offset, title in enumerate(AUDIT_HEADERS):
        cell = ws.cell(row=1, column=first_audit + offset, value=title)
        cell.font = head_font
        cell.alignment = Alignment(horizontal="left")

    report = []
    tally = collections.Counter()

    for excel_row in range(2, ws.max_row + 1):
        values = [c.value for c in ws[excel_row]]
        row = Row(excel_row, values, col)
        cand, method, confidence, note, ranked = match_row(
            row, by_county, exact_idx, token_idx)

        new_api = cand.api if cand else None
        matched_name = lookup[cand.api].WellName if cand else ""
        corroborated = corroborating_fields(row, cand, lookup)

        if cand:
            status = "unchanged" if new_api == row.api else "corrected"
            ws.cell(row=excel_row, column=api_col, value=new_api).font = body_font
        elif row.api:
            status = "flagged_kept_original"
        else:
            status = "flagged_blank" if confidence == "review" else "no_api_expected"

        source = {
            "corrected": "Enverus match",
            "unchanged": "Enverus match (already correct)",
            "flagged_kept_original": "UNMATCHED - original value retained, do not use",
            "flagged_blank": "UNMATCHED - review",
            "no_api_expected": "Undrilled location - no matching well in Enverus",
        }[status]

        audit = [source, row.api or "", matched_name, method, confidence,
                 corroborated, note]
        for offset, value in enumerate(audit):
            cell = ws.cell(row=excel_row, column=first_audit + offset, value=value)
            cell.font = body_font
            cell.alignment = Alignment(horizontal="left", vertical="top")

        if status == "corrected":
            ws.cell(row=excel_row, column=api_col).fill = changed_fill
        elif status.startswith("flagged"):
            ws.cell(row=excel_row, column=api_col).fill = review_fill

        tally[status] += 1
        tally["method:" + method] += 1

        report.append({
            "Excel Row": excel_row,
            "Prop_ID": row.prop_id,
            "RsvCat": row.rsv_cat,
            "Lease": row.lease,
            "Well Number": row.well_no,
            "County": row.county,
            "Operator (header)": row.operator,
            "Original API": row.api or "",
            "Corrected API": new_api or "",
            "Status": status,
            "Match Method": method,
            "Confidence": confidence,
            "Corroborating Fields": corroborated,
            "Matched Enverus Well Name": matched_name,
            "Matched Operator": lookup[cand.api].ENVOperator if cand else "",
            "Matched Spud": lookup[cand.api].SpudDate if cand else "",
            "Matched First Prod": lookup[cand.api].FirstProdDate if cand else "",
            "Matched Lateral": lookup[cand.api].LateralLength_FT if cand else "",
            "Matched Status": lookup[cand.api].ENVWellStatus if cand else "",
            "API Drift (corrected - original)": (
                int(new_api) - int(row.api) if new_api and row.api else ""),
            "Notes": note,
            "Runner-up Candidates": " | ".join(
                "%s %s (sim %.2f, corr %d)" % (c.api, c.name, c.sim, c.corr)
                for c in ranked[1:4]),
        })

    for offset, title in enumerate(AUDIT_HEADERS):
        letter = openpyxl.utils.get_column_letter(first_audit + offset)
        ws.column_dimensions[letter].width = max(16, min(52, len(title) + 10))
    ws.freeze_panes = "A2"

    wb.save(args.out)

    write_report(args.report, report, os.path.basename(args.header),
                 os.path.basename(args.enverus))

    for key in sorted(tally):
        print("%-34s %d" % (key, tally[key]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
