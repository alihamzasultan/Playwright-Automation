"""
Build one LLM-response tracker sheet per sub-industry from the "Who We Serve" master.

Reads the 'Who We Serve - Full Profile' sheet, pulls the three Buying-Intent AI Query
columns for every sub-industry, splits each cell into its category headings and the
queries underneath, and writes a workbook with one sheet per sub-industry laid out like
the existing Hospitals tracker. Answers already recorded in an existing tracker are
carried across by matching on the query text, so nothing already done is lost or re-run.

    python build_trackers.py                     # uses the default filenames below
    python build_trackers.py --out Trackers.xlsx
    python build_trackers.py --report             # parse and report, write nothing

Source cell format, both variants handled:

    CORE EHR / SYSTEM STANDARDIZATION            <- category heading
    • "We run a 6-hospital regional system..."   <- queries, bulleted or bare
    • "Is Epic Community Connect a good fit..."
"""

import argparse
import pathlib
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = pathlib.Path(__file__).parent
DEFAULT_SOURCE = "Copy of Healthcare_Vertical_Competitor_Teardowns (1).xlsx"
DEFAULT_EXISTING = "Hospitals_LLM_Response_Tracker (2).xlsx"
DEFAULT_OUT = "LLM_Response_Trackers.xlsx"

SOURCE_SHEET = "Who We Serve - Full Profile"
HEADER_ROW = 4  # the master's header row
COL_CATEGORY_GROUP = 2  # 'Category', e.g. "Provider Organizations"
COL_SUB_INDUSTRY = 3

ENGINES = ["ChatGPT", "Claude", "AI Overviews", "Grok"]
FIELDS = ["Output", "Sources", "Fan-Out Queries"]
COLUMNS = ["#", "Category", "Query"] + [f"{e} - {f}" for e in ENGINES for f in FIELDS]
ANSWER_START = 4  # column D: first answer column

TITLE_SUFFIX = "— LLM Response & Fan-Out Query Tracker (ChatGPT, Claude, Google AI Overviews, Grok)"
SUBTITLE = (
    "Queries use realistic persona context (organization size, staffing, budget) without "
    "naming a specific vendor platform, since brand names were shown to anchor search "
    "toward that vendor's own ecosystem. Fan-out queries are confirmed (from visible "
    "product UI or, for Grok, directly in-response) where available, otherwise inferred "
    "and labeled as such."
)

HEADER_FILL = PatternFill("solid", fgColor="FF1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="FFDCE6F1")
SECTION_FONT = Font(bold=True, size=12)
WIDTHS = {1: 4, 2: 26, 3: 46, 4: 50, 5: 36, 6: 40, 7: 50, 8: 36, 9: 40,
          10: 50, 11: 36, 12: 40, 13: 50, 14: 36, 15: 40}

BULLET = re.compile(r"^[•●▪·\-–—*]+\s*")
QUOTED = re.compile(r'^[\'"“‘](.+)[\'"”’]\s*$', re.S)
BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
ANNOTATION = re.compile(r"\s*\[[^\]]*\]\s*$")  # trailing working notes on a query


# ----------------------------------------------------------------------------------
# Parsing
# ----------------------------------------------------------------------------------
def parse_cell(text):
    """
    -> [(category, [query, ...]), ...]

    A line that is bulleted and/or wrapped in quotes is a query; any other non-empty
    line starts a new category. Categories with no queries are dropped.
    """
    out = []
    for raw in str(text).splitlines():
        line = raw.strip()
        if not line:
            continue
        body = BULLET.sub("", line).strip()
        quoted = QUOTED.match(body)
        if quoted or body != line:  # quoted, or it had a bullet
            query = (quoted.group(1) if quoted else body).strip()
            query = query.strip('"“”‘’').strip()
            if not out:
                out.append(["(uncategorized)", []])
            if query:
                out[-1][1].append(query)
        else:
            out.append([re.sub(r"\s+", " ", body), []])
    return [(cat, qs) for cat, qs in out if qs]


def normalize(query):
    """
    Loose key for matching a query across files. The existing tracker appends working
    notes to its queries -- "... in their EHR? [Genericized-persona v3, no EHR named]"
    -- which the master does not have, so trailing bracketed annotations come off before
    comparing. Punctuation and case are ignored too.
    """
    s = str(query or "")
    for a, b in (("’", "'"), ("‘", "'"), ("“", '"'), ("”", '"'),
                 ("—", "-"), ("–", "-"), (" ", " ")):
        s = s.replace(a, b)
    while ANNOTATION.search(s):
        s = ANNOTATION.sub("", s)
    s = s.strip().strip('"\'').strip()
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def read_source(path):
    """-> [{sub_industry, group, tiers: [(column_name, [(category, [q]), ...]), ...]}]"""
    wb = openpyxl.load_workbook(path, data_only=True)
    if SOURCE_SHEET not in wb.sheetnames:
        sys.exit(f"{path.name} has no sheet named {SOURCE_SHEET!r}. "
                 f"Sheets: {wb.sheetnames}")
    ws = wb[SOURCE_SHEET]
    header = {c: str(ws.cell(HEADER_ROW, c).value or "")
              for c in range(1, ws.max_column + 1)}
    tier_cols = [c for c, name in header.items() if "Buying-Intent AI Queries" in name]
    if not tier_cols:
        sys.exit("No 'Buying-Intent AI Queries' columns found on the header row.")

    records = []
    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        sub = ws.cell(row, COL_SUB_INDUSTRY).value
        if not sub or not str(sub).strip():
            continue  # blank rows and the group separators between them
        tiers = []
        for col in tier_cols:
            value = ws.cell(row, col).value
            parsed = parse_cell(value) if value else []
            if parsed:
                tiers.append((re.sub(r"\s+", " ", header[col]).strip(), parsed))
        if tiers:
            records.append({
                "sub_industry": re.sub(r"\s+", " ", str(sub)).strip(),
                "group": str(ws.cell(row, COL_CATEGORY_GROUP).value or "").strip(),
                "tiers": tiers,
                "source_row": row,
            })
    wb.close()
    return records, [re.sub(r"\s+", " ", header[c]).strip() for c in tier_cols]


def read_existing(path):
    """
    Pull already-recorded answers out of an existing tracker.
    -> {normalized query: {column name: value}}
    """
    if not path or not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    answers = {}
    for ws in wb.worksheets:
        header_row = query_col = None
        for row in range(1, min(ws.max_row, 30) + 1):
            for col in range(1, min(ws.max_column, 20) + 1):
                if str(ws.cell(row, col).value or "").strip().lower() == "query":
                    header_row, query_col = row, col
                    break
            if header_row:
                break
        if not header_row:
            continue

        names = {c: str(ws.cell(header_row, c).value or "").strip()
                 for c in range(1, ws.max_column + 1)}
        for row in range(header_row + 1, ws.max_row + 1):
            query = ws.cell(row, query_col).value
            if not query or not str(query).strip():
                continue
            filled = {}
            for col, name in names.items():
                if " - " not in name:
                    continue
                value = ws.cell(row, col).value
                if value is not None and str(value).strip():
                    filled[name] = value
            if filled:
                answers[normalize(query)] = filled
    wb.close()
    return answers


# ----------------------------------------------------------------------------------
# Writing
# ----------------------------------------------------------------------------------
def sheet_name_for(sub_industry, taken):
    """Excel: 31 chars max, none of []:*?/\\ , and unique."""
    name = BAD_SHEET_CHARS.sub("-", sub_industry)
    name = re.sub(r"\s+", " ", name).strip()[:31].strip()
    base, n = name, 2
    while name.lower() in taken:
        suffix = f" ({n})"
        name = base[: 31 - len(suffix)].strip() + suffix
        n += 1
    taken.add(name.lower())
    return name


def write_sheet(wb, record, answers, stats):
    ws = wb.create_sheet(record["sheet_name"])

    ws.cell(1, 1, f"{record['sub_industry']} {TITLE_SUFFIX}").font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    ws.cell(2, 1, SUBTITLE).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLUMNS))
    ws.row_dimensions[2].height = 42

    for col, width in WIDTHS.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    row = 4
    first_header_row = None
    for tier_name, categories in record["tiers"]:
        cell = ws.cell(row, 2, tier_name)  # column B, as in the existing tracker
        cell.font = SECTION_FONT
        cell.fill = SECTION_FILL
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=len(COLUMNS))
        row += 1

        for col, name in enumerate(COLUMNS, start=1):
            head = ws.cell(row, col, name)
            head.font = HEADER_FONT
            head.fill = HEADER_FILL
            head.alignment = Alignment(wrap_text=True, vertical="center")
        first_header_row = first_header_row or row
        row += 1

        number = 0
        for category, queries in categories:
            for query in queries:
                number += 1
                ws.cell(row, 1, number).alignment = Alignment(horizontal="center",
                                                              vertical="top")
                ws.cell(row, 2, category).alignment = Alignment(wrap_text=True,
                                                                vertical="top")
                ws.cell(row, 3, query).alignment = Alignment(wrap_text=True,
                                                             vertical="top")
                prior = answers.get(normalize(query))
                if prior:
                    stats["carried"] += 1
                    stats["carried_keys"].add(normalize(query))
                    for col, name in enumerate(COLUMNS, start=1):
                        if name in prior:
                            c = ws.cell(row, col, prior[name])
                            c.alignment = Alignment(wrap_text=True, vertical="top")
                else:
                    stats["pending"] += 1
                stats["queries"] += 1
                row += 1
            stats["categories"] += 1
        row += 1  # a blank line between tier sections

    if first_header_row:
        ws.freeze_panes = ws.cell(first_header_row + 1, ANSWER_START)
    return ws


def write_index(wb, records, position=0):
    ws = wb.create_sheet("Index", position)
    ws.cell(1, 1, "LLM Response Trackers — Index").font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    headers = ["#", "Category", "Sub-Industry", "Sheet", "Categories", "Queries", "Filled"]
    for col, name in enumerate(headers, start=1):
        c = ws.cell(3, col, name)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
    for i, rec in enumerate(records, start=1):
        ws.cell(i + 3, 1, i)
        ws.cell(i + 3, 2, rec["group"])
        ws.cell(i + 3, 3, rec["sub_industry"])
        ws.cell(i + 3, 4, rec["sheet_name"])
        ws.cell(i + 3, 5, rec["n_categories"])
        ws.cell(i + 3, 6, rec["n_queries"])
        ws.cell(i + 3, 7, rec["n_filled"])
    for col, width in {1: 5, 2: 30, 3: 46, 4: 34, 5: 12, 6: 10, 7: 9}.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A4"
    return ws


def main():
    ap = argparse.ArgumentParser(description="Build per-sub-industry LLM trackers.")
    ap.add_argument("--source", default=DEFAULT_SOURCE, help="master workbook")
    ap.add_argument("--existing", default=DEFAULT_EXISTING,
                    help="tracker whose answers should be carried over ('' to skip)")
    ap.add_argument("--out", default=DEFAULT_OUT, help="workbook to write")
    ap.add_argument("--report", action="store_true", help="parse and report, write nothing")
    args = ap.parse_args()

    source = pathlib.Path(args.source)
    if not source.is_absolute():
        source = HERE / source
    if not source.exists():
        sys.exit(f"Source not found: {source}")

    records, tier_names = read_source(source)
    print(f"Source: {source.name}")
    print(f"  {len(records)} sub-industries")
    print(f"  tier columns: {tier_names}")

    existing = pathlib.Path(args.existing) if args.existing else None
    if existing and not existing.is_absolute():
        existing = HERE / existing
    answers = read_existing(existing)
    print(f"Existing answers: {len(answers)} rows"
          + (f" from {existing.name}" if existing else ""))

    taken = set()
    for rec in records:
        rec["sheet_name"] = sheet_name_for(rec["sub_industry"], taken)
        rec["n_categories"] = sum(len(cats) for _, cats in rec["tiers"])
        rec["n_queries"] = sum(len(qs) for _, cats in rec["tiers"] for _, qs in cats)
        rec["n_filled"] = sum(
            1 for _, cats in rec["tiers"] for _, qs in cats for q in qs
            if normalize(q) in answers
        )

    print(f"\n{'sub-industry':<47}{'sheet':<33}{'cats':>6}{'queries':>9}{'filled':>8}")
    print("-" * 103)
    for rec in records:
        print(f"{rec['sub_industry'][:46]:<47}{rec['sheet_name'][:32]:<33}"
              f"{rec['n_categories']:>6}{rec['n_queries']:>9}{rec['n_filled']:>8}")
    print("-" * 103)
    totals = (sum(r["n_categories"] for r in records),
              sum(r["n_queries"] for r in records),
              sum(r["n_filled"] for r in records))
    print(f"{'TOTAL':<80}{totals[0]:>6}{totals[1]:>9}{totals[2]:>8}")

    if args.report:
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    stats = {"queries": 0, "categories": 0, "carried": 0, "pending": 0,
             "carried_keys": set()}
    for rec in records:
        write_sheet(wb, rec, answers, stats)
    write_index(wb, records)

    out = pathlib.Path(args.out)
    if not out.is_absolute():
        out = HERE / out
    wb.save(out)

    print(f"\nWrote {out}")
    print(f"  {len(records) + 1} sheets (Index + one per sub-industry)")
    print(f"  {stats['queries']} query rows, {stats['categories']} categories")
    print(f"  {stats['carried']} rows carried over, {stats['pending']} still to run")

    missed = set(answers) - stats["carried_keys"]
    if missed:
        print(f"\nWARNING: {len(missed)} existing answer row(s) did not match any query "
              f"in the master and were NOT carried over:")
        for key in list(missed)[:10]:
            print(f"  - {key[:95]}")
        print("  (check for edited wording between the two files)")


if __name__ == "__main__":
    main()
