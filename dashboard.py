"""
A live dashboard for the tracker run. Read-only and local.

    python dashboard.py                 # serve on http://localhost:8787 (auto-picks if busy)
    python dashboard.py --port 9000
    python dashboard.py --no-open       # do not launch a browser

Open the URL and leave it up: the page polls for new data every few seconds, so it
follows a run as it happens.

It reads the append-only JSONL for everything live (that file is never rewritten, so
reading it can never disturb a run) and the JSON snapshot for the denominators. The
workbook itself is never opened -- holding an .xlsx open can make the runner's atomic
save fail, which is exactly what this tool must not do.
"""

import argparse
import datetime
import json
import pathlib
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = pathlib.Path(__file__).parent
JSONL = HERE / "LLM_Response_Trackers_answers.jsonl"
JSON = HERE / "LLM_Response_Trackers.json"
PAGE = HERE / "dashboard.html"

# The JSON is rewritten atomically by the runner, so a read can land mid-swap. Keep the
# last good copy and serve that rather than failing.
_cache = {"json": None, "json_at": 0.0}


def read_jsonl():
    records = []
    if not JSONL.exists():
        return records
    try:
        with open(JSONL, "r", encoding="utf-8", errors="replace") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    records.append(json.loads(line))
                except Exception:
                    pass  # a half-written last line while the runner appends
    except Exception:
        pass
    return records


def read_json(max_age=20.0):
    now = time.time()
    if _cache["json"] is not None and now - _cache["json_at"] < max_age:
        return _cache["json"]
    if JSON.exists():
        try:
            _cache["json"] = json.loads(JSON.read_text(encoding="utf-8"))
            _cache["json_at"] = now
        except Exception:
            pass  # mid-swap; keep the previous copy
    return _cache["json"]


def answer_kind(text):
    text = text or ""
    if text.startswith("N/A"):
        return "none"
    if text.startswith("[captured from page text"):
        return "fallback"
    if text.startswith("[AI Overview located by heading"):
        return "heading"
    return "ok"


def build_state():
    records = read_jsonl()
    snapshot = read_json()

    engines, sheets, tiers = {}, {}, {}
    for rec in records:
        eng = engines.setdefault(rec["engine"], {
            "name": rec["engine"], "count": 0, "chars": 0, "with_sources": 0,
            "sources": 0, "fallback": 0, "none": 0,
        })
        eng["count"] += 1
        eng["chars"] += rec.get("chars") or 0
        if rec.get("source_count"):
            eng["with_sources"] += 1
            eng["sources"] += rec["source_count"]
        kind = answer_kind(rec.get("answer"))
        if kind == "fallback":
            eng["fallback"] += 1
        elif kind == "none":
            eng["none"] += 1
        sheets[rec["sub_industry"]] = sheets.get(rec["sub_industry"], 0) + 1
        tiers[rec["tier"]] = tiers.get(rec["tier"], 0) + 1

    for eng in engines.values():
        eng["avg_chars"] = eng["chars"] // max(1, eng["count"])

    # Pace, from the timestamps in the log.
    def parse(ts):
        try:
            return datetime.datetime.fromisoformat(ts)
        except Exception:
            return None

    stamps = [t for t in (parse(r["ts"]) for r in records) if t]
    now = datetime.datetime.now()
    last_10 = sum(1 for t in stamps if (now - t).total_seconds() <= 600)
    last_60 = sum(1 for t in stamps if (now - t).total_seconds() <= 3600)

    idle = None
    if JSONL.exists():
        idle = time.time() - JSONL.stat().st_mtime

    # Denominators and the authoritative answered counts come from the snapshot, which
    # is built from the workbook and so includes the manually carried-over rows.
    totals = {"queries": 0, "answers": {}, "engines": []}
    sheet_rows = {}
    if snapshot:
        totals["queries"] = snapshot["totals"]["queries"]
        totals["answers"] = snapshot["totals"]["answers"]
        totals["engines"] = snapshot.get("engines", [])
        for sub in snapshot.get("sub_industries", []):
            rows = sum(len(c["queries"]) for t in sub["tiers"] for c in t["categories"])
            answered = 0
            for tier in sub["tiers"]:
                for cat in tier["categories"]:
                    for q in cat["queries"]:
                        if any(v and v.get("output") for v in q["answers"].values()):
                            answered += 1
            sheet_rows[sub["sheet"]] = {
                "sub_industry": sub["sub_industry"], "rows": rows,
                "rows_answered": answered, "logged": sheets.get(sub["sheet"], 0),
            }

    engine_names = totals["engines"] or sorted(engines)
    cells_total = totals["queries"] * max(1, len(engine_names))
    cells_done = sum(totals["answers"].values()) if totals["answers"] else 0

    recent = []
    for rec in records[-30:][::-1]:
        recent.append({
            "ts": rec["ts"][11:], "engine": rec["engine"], "row": rec["row"],
            "sheet": rec["sub_industry"], "category": rec["category"],
            "query": rec["query"], "chars": rec.get("chars") or 0,
            "sources": rec.get("source_count") or 0,
            "kind": answer_kind(rec.get("answer")),
        })

    return {
        "generated": now.isoformat(timespec="seconds"),
        "logged": len(records),
        "cells_done": cells_done,
        "cells_total": cells_total,
        "queries": totals["queries"],
        "engine_names": engine_names,
        "answers_by_engine": totals["answers"],
        "engines": [engines.get(name, {
            "name": name, "count": 0, "avg_chars": 0, "with_sources": 0,
            "sources": 0, "fallback": 0, "none": 0,
        }) for name in engine_names],
        "sheets": sorted(sheet_rows.values(), key=lambda s: -s["rows_answered"]),
        "tiers": [{"tier": k, "count": v} for k, v in
                  sorted(tiers.items(), key=lambda kv: -kv[1])],
        "pace_10m": last_10,
        "pace_60m": last_60,
        "idle_seconds": round(idle) if idle is not None else None,
        "active": bool(idle is not None and idle < 300),
        "recent": recent,
        "has_snapshot": snapshot is not None,
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass  # keep the console quiet; the run's own output matters more

    def _send(self, code, body, content_type):
        payload = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(payload)
        except Exception:
            pass

    def do_GET(self):
        if self.path.startswith("/api/state"):
            try:
                self._send(200, json.dumps(build_state()), "application/json")
            except Exception as exc:
                self._send(500, json.dumps({"error": f"{exc.__class__.__name__}: {exc}"}),
                           "application/json")
            return
        if self.path in ("/", "/index.html", "/dashboard.html"):
            if not PAGE.exists():
                self._send(500, f"{PAGE.name} is missing next to dashboard.py",
                           "text/plain; charset=utf-8")
                return
            self._send(200, PAGE.read_text(encoding="utf-8"),
                       "text/html; charset=utf-8")
            return
        self._send(404, "not found", "text/plain; charset=utf-8")


def main():
    ap = argparse.ArgumentParser(description="Live dashboard for the tracker run.")
    ap.add_argument("--port", type=int, default=8787)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--no-open", action="store_true", help="don't launch a browser")
    ap.add_argument("--once", action="store_true",
                    help="print the state as JSON and exit (no server)")
    args = ap.parse_args()

    if args.once:
        print(json.dumps(build_state(), indent=2)[:4000])
        return

    # Something else may already hold the port -- on Windows that surfaces as
    # WinError 10013 (access forbidden) rather than "address in use" -- so try a few.
    server = None
    for port in range(args.port, args.port + 12):
        try:
            server = ThreadingHTTPServer((args.host, port), Handler)
            if port != args.port:
                print(f"port {args.port} was busy; using {port}")
            args.port = port
            break
        except OSError:
            continue
    if server is None:
        raise SystemExit(f"Could not bind any port from {args.port} to {args.port + 11}. "
                         "Pass --port with a free one.")

    url = f"http://{args.host}:{args.port}/"
    print(f"dashboard : {url}")
    print(f"reading   : {JSONL.name} (live) + {JSON.name} (totals)")
    print("the workbook is never opened, so this cannot interfere with a run")
    print("Ctrl+C to stop\n")
    if not args.no_open:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nstopped.")
        server.server_close()


if __name__ == "__main__":
    main()
