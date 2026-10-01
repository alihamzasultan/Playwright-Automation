"""
Look at the answers while a run is in progress. Read-only, and it never touches the
workbook, so it cannot interfere with a live run or trip Excel's file lock.

    python show.py                    # the last 10 answers
    python show.py --progress         # counts per engine and per sub-industry
    python show.py --follow           # live, as answers arrive
    python show.py --row 38           # every engine's answer for one row, in full
    python show.py --engine grok -n 5 # filter by engine
    python show.py --sources          # show the cited links too

It reads the JSONL log, which is appended the instant each answer arrives and is never
rewritten -- so it is always at least as current as the workbook.
"""

import argparse
import collections
import json
import pathlib
import time

HERE = pathlib.Path(__file__).parent
DEFAULT_JSONL = "LLM_Response_Trackers_answers.jsonl"


def read(path):
    rows = []
    if not path.exists():
        return rows
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if line:
            try:
                rows.append(json.loads(line))
            except Exception:
                pass
    return rows


def one_line(rec, width=64):
    tag = ""
    answer = rec.get("answer") or ""
    if answer.startswith("N/A"):
        tag = " [no overview]"
    elif answer.startswith("[captured from page text"):
        tag = " [text fallback]"
    elif answer.startswith("[AI Overview located by heading"):
        tag = " [by heading]"
    return (f"  {rec['ts'][11:]}  {rec['engine']:<13} r{rec['row']:<4} "
            f"{rec['chars']:>6}ch  src={rec.get('source_count', 0):<3} "
            f"{rec['query'][:width]}{tag}")


def show_progress(rows):
    print(f"\n{len(rows)} answers captured by the runner "
          "(your carried-over manual rows are not in this log)\n")
    per_engine = collections.Counter(r["engine"] for r in rows)
    na = collections.Counter(r["engine"] for r in rows
                             if (r.get("answer") or "").startswith("N/A"))
    print("  per engine:")
    for engine, count in per_engine.most_common():
        extra = f"   ({na[engine]} of them 'no overview')" if na[engine] else ""
        chars = [r["chars"] for r in rows if r["engine"] == engine]
        avg = sum(chars) // max(1, len(chars))
        with_src = sum(1 for r in rows
                       if r["engine"] == engine and r.get("source_count"))
        print(f"    {engine:<14} {count:>5}   avg {avg:>6}ch   "
              f"{with_src}/{count} with sources{extra}")

    print("\n  per sub-industry:")
    for sheet, count in collections.Counter(r["sub_industry"] for r in rows).most_common():
        print(f"    {sheet:<34} {count:>5}")

    print("\n  per tier:")
    for tier, count in collections.Counter(r["tier"] for r in rows).most_common():
        print(f"    {tier[:44]:<45} {count:>5}")

    if len(rows) > 1:
        first = rows[0]["ts"]
        last = rows[-1]["ts"]
        print(f"\n  first answer {first[11:]}   latest {last[11:]}")
        try:
            t0 = time.mktime(time.strptime(first, "%Y-%m-%dT%H:%M:%S"))
            t1 = time.mktime(time.strptime(last, "%Y-%m-%dT%H:%M:%S"))
            mins = max(1e-9, (t1 - t0) / 60)
            print(f"  pace  {len(rows) / mins:.1f} answers/min over "
                  f"{mins:.0f} min of elapsed time (gaps included)")
        except Exception:
            pass
    stale = time.time() - pathlib.Path(DEFAULT_JSONL).stat().st_mtime \
        if pathlib.Path(DEFAULT_JSONL).exists() else None
    if stale is not None:
        state = "a run looks active" if stale < 300 else "no run appears active"
        print(f"  last write {stale:.0f}s ago -- {state}")


def show_row(rows, row_no, width):
    hits = [r for r in rows if r["row"] == row_no]
    if not hits:
        print(f"\nNo logged answers for row {row_no}.")
        return
    print(f"\nRow {row_no} -- {hits[0]['sub_industry']}")
    print(f"tier     : {hits[0]['tier']}")
    print(f"category : {hits[0]['category']}")
    print(f"query    : {hits[0]['query']}\n")
    for rec in hits:
        print("=" * 78)
        print(f"{rec['engine']}  ({rec['chars']:,} chars, "
              f"{rec.get('source_count', 0)} sources, {rec['ts'][11:]})")
        print("=" * 78)
        print(rec["answer"][:width])
        if len(rec["answer"]) > width:
            print(f"... [{len(rec['answer']) - width:,} more characters]")
        if rec.get("sources"):
            print("\n-- sources --")
            print(rec["sources"][:1200])
        print()


def main():
    ap = argparse.ArgumentParser(description="View captured answers (read-only).")
    ap.add_argument("--jsonl", default=DEFAULT_JSONL)
    ap.add_argument("-n", "--count", type=int, default=10, help="how many to list")
    ap.add_argument("--progress", action="store_true", help="counts and pace")
    ap.add_argument("--follow", action="store_true", help="keep printing new answers")
    ap.add_argument("--row", type=int, help="show full answers for one row")
    ap.add_argument("--engine", help="filter: chatgpt, claude, grok, overview")
    ap.add_argument("--sheet", help="filter by sub-industry name")
    ap.add_argument("--sources", action="store_true", help="include sources in the list")
    ap.add_argument("--width", type=int, default=4000, help="chars of answer to print")
    args = ap.parse_args()

    path = pathlib.Path(args.jsonl)
    if not path.is_absolute():
        path = HERE / path
    rows = read(path)
    if not rows:
        print(f"No answers logged yet ({path.name} is missing or empty).")
        return

    if args.engine:
        want = args.engine.lower()
        rows = [r for r in rows if want in r["engine"].lower()]
    if args.sheet:
        want = args.sheet.lower()
        rows = [r for r in rows if want in r["sub_industry"].lower()]

    if args.progress:
        show_progress(rows)
        return
    if args.row:
        show_row(rows, args.row, args.width)
        return

    print(f"\nlast {min(args.count, len(rows))} of {len(rows)} answers "
          f"(from {path.name}):\n")
    for rec in rows[-args.count:]:
        print(one_line(rec))
        if args.sources and rec.get("sources"):
            for line in rec["sources"].splitlines()[1:4]:
                print(f"        {line[:96]}")

    if args.follow:
        print("\nwatching for new answers -- Ctrl+C to stop\n")
        seen = len(read(path))
        try:
            while True:
                time.sleep(5)
                current = read(path)
                for rec in current[seen:]:
                    print(one_line(rec))
                seen = len(current)
        except KeyboardInterrupt:
            print("\nstopped watching.")


if __name__ == "__main__":
    main()
