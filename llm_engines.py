"""
The four engines the tracker has columns for, behind one interface.

Each engine says where its new-chat page is, how to find the composer, how to find the
latest answer and how to tell that the answer is still streaming. run_tracker.py then
treats them all the same way.

A warning worth reading before you debug anything here: these are other people's
front-ends, and their markup changes without notice. Every selector is a LIST of
candidates tried in order, and each engine has a probe you can run to see which ones
actually matched today:

    python llm_engines.py --probe claude
    python llm_engines.py --probe all --query "test question"

If an engine stops working, run its probe, then fix the list at the top of this file --
that is the only place selectors live.
"""

import argparse
import os
import pathlib
import random
import re
import sys
import time
import urllib.parse

import chatgpt_query as cgq

# Set LLM_DEBUG_DIR to have a page saved whenever an answer cannot be found. Google
# renames the AI Overview container regularly, and a saved page is the only way to pin
# the current selector without guessing.
DEBUG_DIR = os.environ.get("LLM_DEBUG_DIR") or ""


def dump_page(page, tag):
    """Save the live HTML so a missing selector can be diagnosed afterwards."""
    if not DEBUG_DIR:
        return None
    try:
        folder = pathlib.Path(DEBUG_DIR)
        folder.mkdir(parents=True, exist_ok=True)
        out = folder / f"{tag}-{time.strftime('%H%M%S')}.html"
        out.write_text(page.content(), encoding="utf-8")
        return out
    except Exception:
        return None


class Engine:
    def __init__(self, key, label, url, composer, answer, streaming=(), send=(),
                 login=(), kind="chat", note="", noise=()):
        self.key = key
        self.label = label          # must match the workbook column prefix
        self.url = url              # a fresh conversation every time
        self.composer = list(composer)
        self.answer = list(answer)
        self.streaming = list(streaming)
        self.send = list(send)
        self.login = list(login)
        self.kind = kind            # "chat" types a question; "search" loads a URL
        self.note = note
        # Interface text that sits inside the message element and is not part of the
        # answer: screen-reader labels, "thought for 13s" timers and the like.
        self.noise = [re.compile(p, re.I) for p in noise]

    def __repr__(self):
        return f"<Engine {self.key}>"


ENGINES = {
    "chatgpt": Engine(
        key="chatgpt", label="ChatGPT", url="https://chatgpt.com/",
        composer=["#prompt-textarea", "div[contenteditable='true'][id='prompt-textarea']"],
        answer=["[data-message-author-role='assistant']"],
        streaming=["[data-testid='stop-button']", "button[aria-label*='Stop']"],
        send=["[data-testid='send-button']", "button[aria-label*='Send']"],
        login=["[data-testid='login-button']", "a[href*='/auth/login']"],
    ),
    "claude": Engine(
        key="claude", label="Claude", url="https://claude.ai/new",
        composer=["div[contenteditable='true'].ProseMirror",
                  "[data-testid='chat-input']",
                  "div.ProseMirror[contenteditable='true']",
                  "div[contenteditable='true']"],
        answer=["div[data-is-streaming]", ".font-claude-message",
                "[data-testid='assistant-message']",
                "div[class*='font-claude']"],
        streaming=["[data-is-streaming='true']", "button[aria-label*='Stop']"],
        send=["button[aria-label='Send message']", "button[aria-label*='Send']",
              "button[type='submit']"],
        login=["a[href*='/login']", "button:has-text('Sign in')"],
        # A screen-reader label inside the message element, seen in live output as
        # "Claude responded: For a 40-bed shop, ...".
        noise=[r"^\s*Claude(?:'s)?\s+(?:responded|response)\s*:\s*",
               r"^\s*Claude\s+said\s*:\s*"],
    ),
    "grok": Engine(
        key="grok", label="Grok", url="https://grok.com/",
        # Grok's composer is a ProseMirror contenteditable, not a textarea. The textarea
        # candidates stay last: the page has other textareas, and picking one of those
        # meant the typing went nowhere and the editor stayed empty
        # (<p class="is-empty is-editor-empty">).
        composer=["div.ProseMirror[contenteditable='true']",
                  "div[contenteditable='true'].ProseMirror",
                  ".ProseMirror[contenteditable='true']",
                  "div[contenteditable='true']",
                  "textarea[aria-label*='Ask']", "textarea[placeholder*='Ask']",
                  "textarea"],
        answer=["div.message-bubble", "[data-testid='message']",
                "div[class*='response-content']", "div.prose",
                "div[class*='markdown']", "div[class*='prose']",
                "div[class*='message']"],
        streaming=["button[aria-label*='Stop']", "button[aria-label*='stop']"],
        send=["button[type='submit']", "button[aria-label*='Submit']",
              "button[aria-label*='Send']"],
        login=["a[href*='/sign-in']", "a[href*='x.com/i/oauth']",
               "button:has-text('Sign in with X')", "button:has-text('Sign in')",
               "button:has-text('Log in')"],
        # Grok puts its own timer at the top of the message: "Worked for 13s ...".
        noise=[r"^\s*(?:Worked|Thought|Thinking|Reasoned)\s+for\s+[\d.]+\s*"
               r"(?:s|sec|secs|seconds|m|min|minutes)\b\s*",
               r"^\s*Expert\s*\n"],
    ),
    # Google is not a chat: the query goes in the URL and the AI Overview either appears
    # or it does not. Plenty of queries simply never get one.
    "aio": Engine(
        key="aio", label="AI Overviews",
        url="https://www.google.com/search?hl=en&q={query}",
        composer=[],
        answer=["div[data-attrid='SGEAnswer']", "div[data-subtree='aio']",
                "#eobd", "div[jsname='xXq91c']", "div[aria-label*='AI Overview']"],
        streaming=[],
        send=[],
        login=[],
        kind="search",
        note="No AI Overview is shown for many queries; those rows are recorded as N/A.",
    ),
}
ORDER = ["chatgpt", "claude", "aio", "grok"]
NO_OVERVIEW = "N/A — Google showed no AI Overview for this query."

# Find the AI Overview by its own heading and walk up to the block that holds it. Google
# changes those container attributes often, so the heading text is the stabler anchor.
_AIO_BY_HEADING_JS = """() => {
  const labels = ['ai overview', 'ai-powered overview'];
  const heads = Array.from(document.querySelectorAll('h1,h2,h3,div,span,strong'))
    .filter(el => {
      const t = (el.textContent || '').trim().toLowerCase();
      return t.length < 40 && labels.some(l => t === l || t.startsWith(l));
    });
  for (const head of heads) {
    let node = head;
    for (let up = 0; up < 6 && node; up++) {
      node = node.parentElement;
      if (!node) break;
      const text = (node.innerText || '').trim();
      if (text.length > 200) {
        // Drop the heading itself and anything after the results start.
        return text.split('\\n').filter(line => {
          const l = line.trim().toLowerCase();
          return l && !labels.includes(l);
        }).join('\\n');
      }
    }
  }
  return null;
}"""

CONSENT = ["button:has-text('Accept all')", "button:has-text('I agree')",
           "button[aria-label='Accept all']", "#L2AGLb"]
SHOW_MORE = ["div[aria-label*='Show more']", "button[aria-label*='Show more']",
             "div[jsname='rPRdsc']", "button:has-text('Show more')"]


def settle(page, timeout=25):
    """
    Let redirects finish before touching the page. Several of these sites bounce through
    an auth or Cloudflare hop immediately after load, and anything read during that hop
    dies with "Execution context was destroyed, most likely because of a navigation".
    """
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        try:
            page.wait_for_load_state("domcontentloaded", timeout=6_000)
            url = page.url
            if url == last:  # two reads agree: the bouncing has stopped
                return True
            last = url
        except Exception:
            pass
        time.sleep(0.8)
    return False


def safe(call, default=None, tries=4):
    """Retry a page read that a navigation may have interrupted."""
    for attempt in range(tries):
        try:
            return call()
        except Exception:
            if attempt == tries - 1:
                return default
            time.sleep(0.7)
    return default


def first_present(page, selectors, timeout=0, visible=True):
    """
    Locator for the first selector that matches, else (None, None).

    Visibility matters: pages carry hidden inputs and off-screen duplicates, and matching
    one of those is how typing ends up going nowhere. A visible match is preferred, and a
    hidden one is only used if nothing visible matches at all.
    """
    deadline = time.time() + timeout
    while True:
        hidden_hit = None
        for selector in selectors:
            try:
                locator = page.locator(selector)
                count = locator.count()
                if not count:
                    continue
                if not visible:
                    return locator, selector
                for i in range(min(count, 5)):
                    candidate = locator.nth(i)
                    try:
                        if candidate.is_visible():
                            return candidate, selector
                    except Exception:
                        continue
                hidden_hit = hidden_hit or (locator, selector)
            except Exception:
                continue
        if hidden_hit and time.time() >= deadline:
            return hidden_hit
        if time.time() >= deadline:
            return None, None
        time.sleep(0.4)


def count_answers(page, engine):
    total = 0
    for selector in engine.answer:
        try:
            total = max(total, page.locator(selector).count())
        except Exception:
            pass
    return total


def latest_answer_text(page, engine):
    for selector in engine.answer:
        try:
            locator = page.locator(selector)
            if locator.count():
                return locator.last.inner_text()
        except Exception:
            continue
    return ""


def is_streaming(page, engine):
    for selector in engine.streaming:
        try:
            if page.locator(selector).count():
                return True
        except Exception:
            continue
    return False


def clean(engine, text):
    """Strip interface text that lives inside the message element."""
    out = str(text or "")
    for pattern in engine.noise:
        out = pattern.sub("", out, count=1)
    return re.sub(r"\n{3,}", "\n\n", out).strip()


def page_text(page):
    """All the visible text on the page, for the selector-free fallback."""
    for selector in ("main", "body"):
        try:
            locator = page.locator(selector)
            if locator.count():
                return locator.first.inner_text()
        except Exception:
            continue
    return ""


def new_text_since(base, current):
    """
    Lines present now that were not there before the question was sent. Crude, but it
    does not care what the site's markup looks like, which is the point.
    """
    seen = {line.strip() for line in base.splitlines() if line.strip()}
    fresh = [line for line in current.splitlines()
             if line.strip() and line.strip() not in seen]
    return "\n".join(fresh).strip()


def wait_until_settled(page, engine, before, timeout, quiet=2.5, base_text=""):
    """
    Wait for the answer, then for it to stop changing.

    Two ways of seeing the answer. The engine's selectors are preferred. If none of them
    ever match -- a UI redesign, or selectors that were guessed rather than verified --
    it falls back to watching the page's own text grow, which works on any front-end.
    Fallback answers are labelled, because they can pick up surrounding interface text.
    """
    deadline = time.time() + timeout
    mode = None
    while time.time() < deadline:
        if count_answers(page, engine) > before:
            mode = "selector"
            break
        if base_text and len(new_text_since(base_text, page_text(page))) > 150:
            mode = "text"
            break
        time.sleep(random.uniform(0.3, 0.7))
    if mode is None:
        raise TimeoutError(
            f"{engine.label} never started answering. Its answer selectors "
            f"{engine.answer} did not match and no new page text appeared -- run: "
            f"python llm_engines.py --probe {engine.key}")

    last, stable_since = None, None
    while time.time() < deadline:
        if mode == "selector":
            text = latest_answer_text(page, engine)
            streaming = is_streaming(page, engine)
        else:
            text = new_text_since(base_text, page_text(page))
            streaming = False
        if text and text == last and not streaming:
            if stable_since is None:
                stable_since = time.time()
            elif time.time() - stable_since >= quiet:
                break
        else:
            stable_since = None
            last = text
        time.sleep(random.uniform(0.4, 0.8))

    if not last:
        raise TimeoutError(f"{engine.label}'s answer did not finish in time")
    if mode == "text":
        log_note = (f"[captured from page text: {engine.label}'s answer selectors did "
                    f"not match, so this may include interface text. Run "
                    f"'python llm_engines.py --probe {engine.key}' to fix the selectors.]")
        print(f"    note: using the text fallback for {engine.label}", flush=True)
        return f"{log_note}\n\n{clean(engine, last)}"
    return clean(engine, last)


def ask_chat(engine, page, human, query, timeout, manual_challenge=False):
    page.goto(engine.url, wait_until="domcontentloaded")
    settle(page)
    human.pause(0.8, 2.0)
    human.wander()
    # Pass this engine's own composer selectors: the challenge is only really cleared
    # once the engine's page is actually there.
    if not cgq.clear_challenge(page, human, done=engine.composer,
                               manual=manual_challenge):
        raise RuntimeError(
            f"{engine.label}: Cloudflare challenge not cleared. Signing in to "
            f"{engine.url.split('/')[2]} through 'python chatgpt_query.py --login' "
            f"leaves a cf_clearance cookie in the profile, which usually stops it.")

    composer, _ = first_present(page, engine.composer, timeout=45)
    if composer is None:
        if first_present(page, engine.login)[0] is not None:
            raise RuntimeError(f"{engine.label} is signed out in this profile "
                               f"(run: python chatgpt_query.py --login)")
        raise RuntimeError(f"{engine.label}: no composer matched {engine.composer}. "
                           f"Run: python llm_engines.py --probe {engine.key}")

    before = count_answers(page, engine)
    base_text = page_text(page)  # for the selector-free fallback
    target = composer.first
    human.click(target)
    human.pause(0.2, 0.6)
    page.keyboard.press("Meta+A" if sys.platform == "darwin" else "Control+A")
    page.keyboard.press("Backspace")

    if len(query) > 280 and human.paste(query):
        pass
    else:
        human.type(query)
    human.pause(0.4, 1.2)

    button, _ = first_present(page, engine.send)
    sent = False
    if button is not None:
        try:
            if button.first.is_enabled():
                human.click(button.first)
                sent = True
        except Exception:
            sent = False
    if not sent:
        page.keyboard.press("Enter")

    return wait_until_settled(page, engine, before, timeout, base_text=base_text)


def ask_search(engine, page, human, query, timeout):
    """Google AI Overview: load the results page and read the block, if there is one."""
    url = engine.url.format(query=urllib.parse.quote_plus(query))
    page.goto(url, wait_until="domcontentloaded")
    settle(page)
    human.pause(1.0, 2.2)

    consent, _ = first_present(page, CONSENT)
    if consent is not None:
        try:
            human.click(consent.first, timeout=5_000)
            human.pause(0.8, 1.8)
        except Exception:
            pass
    cgq.clear_challenge(page, human, done=["#search", "#rso", "form[action='/search']"])

    block, selector = first_present(page, engine.answer, timeout=12)
    if block is None:
        # Google renames these containers constantly, so before believing there is no
        # overview, look for the "AI Overview" heading itself and read its container.
        found = safe(lambda: page.evaluate(_AIO_BY_HEADING_JS), None)
        if found and len(found) > 120:
            return f"[AI Overview located by heading text]\n\n{found.strip()}"
        saved = dump_page(page, "aio-no-overview")
        if saved:
            # Is the phrase on the page at all? That separates "Google served no
            # overview" from "served one, but we could not find its container".
            body = (safe(lambda: page_text(page), "") or "").lower()
            mentioned = "ai overview" in body
            print(f"    no AI Overview found; page saved to {saved} "
                  f"(phrase 'AI Overview' present on page: {mentioned})", flush=True)
        return NO_OVERVIEW

    more, _ = first_present(page, SHOW_MORE)
    if more is not None:
        try:
            human.click(more.first, timeout=5_000)
            human.pause(1.0, 2.0)
        except Exception:
            pass

    deadline = time.time() + min(timeout, 40)
    last, stable_since = None, None
    while time.time() < deadline:
        try:
            text = block.first.inner_text()
        except Exception:
            break
        if text and text == last:
            if stable_since is None:
                stable_since = time.time()
            elif time.time() - stable_since >= 1.5:
                break
        else:
            stable_since = None
            last = text
        time.sleep(0.5)

    text = (last or "").strip()
    text = re.sub(r"\n{3,}", "\n\n", text)
    if len(text) < 40:
        return NO_OVERVIEW
    return f"[AI Overview via {selector}]\n\n{text}"


EXCEL_CELL_MAX = 32_767  # a hard Excel limit; a longer string corrupts the workbook
# Hosts that are the engine's own interface rather than a cited source.
OWN_HOSTS = ("chatgpt.com", "openai.com", "claude.ai", "anthropic.com", "grok.com",
             "x.com", "twitter.com", "google.com", "gstatic.com", "googleusercontent.com",
             "accounts.google.com", "policies.google.com", "support.google.com")

_LINKS_JS = """(selector) => {
  const nodes = document.querySelectorAll(selector);
  const el = nodes[nodes.length - 1];
  if (!el) return [];
  return Array.from(el.querySelectorAll('a[href]')).map(a => ({
    href: a.href || '',
    text: (a.innerText || a.getAttribute('aria-label') || '').trim(),
  }));
}"""


def for_cell(text):
    """
    Excel refuses anything over 32,767 characters in one cell, so an over-long answer is
    trimmed for the sheet only. The JSON and JSONL copies always hold the full text.
    """
    text = str(text or "")
    if len(text) <= EXCEL_CELL_MAX:
        return text
    keep = EXCEL_CELL_MAX - 120
    return (text[:keep] +
            f"\n\n[... trimmed for Excel: {len(text):,} characters total. "
            f"The full answer is in the JSON and JSONL files.]")


# Tracking parameters the engines bolt onto cited links. ChatGPT appends
# ?utm_source=chatgpt.com to every citation, which makes the same page look like two
# different URLs when you count domains later.
TRACKING_PARAMS = ("utm_source", "utm_medium", "utm_campaign", "utm_term",
                   "utm_content", "utm_id", "ref", "ref_src", "fbclid", "gclid",
                   "igshid", "si", "mc_cid", "mc_eid")
# Grok's citation pills carry a word-joiner before the title ("⁠Curatrix").
INVISIBLE = dict.fromkeys(map(ord, "​‌‍⁠﻿­"), None)


def tidy_url(href):
    """Drop tracking parameters so the same page always reads as the same URL."""
    try:
        parts = urllib.parse.urlsplit(href)
        if not parts.query:
            return href
        kept = [(k, v) for k, v in urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
                if k.lower() not in TRACKING_PARAMS]
        return urllib.parse.urlunsplit(
            (parts.scheme, parts.netloc, parts.path,
             urllib.parse.urlencode(kept), parts.fragment))
    except Exception:
        return href


def tidy_title(text):
    """Strip invisible characters and collapse whitespace in a link's label."""
    return re.sub(r"\s+", " ", str(text or "").translate(INVISIBLE)).strip()


def tidy_sources_block(text):
    """
    Re-clean an already-saved sources cell, for rows captured before the cleanups above.
    Lines look like 'domain - title - url'; anything that does not parse is left alone.
    """
    lines = str(text or "").splitlines()
    out = []
    for line in lines:
        if " http" not in line:
            out.append(line)
            continue
        head, _, url = line.rpartition(" — ") if " — " in line else line.rpartition(" - ")
        if not url.startswith("http"):
            out.append(line)
            continue
        parts = [tidy_title(p) for p in re.split(r"\s+—\s+|\s+-\s+", head) if p.strip()]
        cleaned = tidy_url(url.strip())
        out.append(" — ".join(parts + [cleaned]) if parts else cleaned)
    return "\n".join(out)


def extract_sources(page, engine, limit=80):
    """
    The links cited inside the latest answer, as 'domain — title — url' lines.

    Read from the answer element itself, so it picks up whatever the engine rendered as
    citations. Links back to the engine's own interface are dropped. Returns "" when the
    answer cited nothing, which is a real and common outcome, not a failure.
    """
    rows = []
    for selector in engine.answer:
        rows = safe(lambda s=selector: page.evaluate(_LINKS_JS, s), []) or []
        if rows:
            break

    seen, out = set(), []
    for row in rows:
        href = (row.get("href") or "").strip()
        if not href or href.startswith(("javascript:", "#", "mailto:")):
            continue
        try:
            host = urllib.parse.urlparse(href).netloc.lower().removeprefix("www.")
        except Exception:
            continue
        if not host or any(host == h or host.endswith("." + h) for h in OWN_HOSTS):
            continue
        # Dedupe on the tidied URL, so the same page cited twice with different
        # tracking parameters collapses to one entry.
        key = tidy_url(href).split("#")[0]
        if key in seen:
            continue
        seen.add(key)
        title = tidy_title(row.get("text"))
        clean_href = tidy_url(href)
        out.append(f"{host} — {title} — {clean_href}" if title
                   else f"{host} — {clean_href}")
        if len(out) >= limit:
            break
    if not out:
        return ""
    return f"[{len(out)} link(s) cited in the answer]\n" + "\n".join(out)


def ask(engine, page, human, query, timeout=240, manual_challenge=False):
    if isinstance(engine, str):
        engine = ENGINES[engine]
    if engine.kind == "search":
        return ask_search(engine, page, human, query, timeout)
    return ask_chat(engine, page, human, query, timeout,
                    manual_challenge=manual_challenge)


def resolve(names):
    """'chatgpt,claude' or 'all' -> [Engine, ...] in a sensible order."""
    if not names or names.strip().lower() == "all":
        return [ENGINES[k] for k in ORDER]
    chosen = []
    for raw in names.split(","):
        key = raw.strip().lower()
        if not key:
            continue
        if key in ENGINES:
            chosen.append(ENGINES[key])
            continue
        hits = [e for e in ENGINES.values()
                if key in e.key or key in e.label.lower().replace(" ", "")]
        if not hits:
            sys.exit(f"Unknown engine {raw!r}. Choose from: "
                     f"{', '.join(ORDER)} (or 'all').")
        chosen.append(hits[0])
    seen, out = set(), []
    for engine in chosen:
        if engine.key not in seen:
            seen.add(engine.key)
            out.append(engine)
    return out


# ----------------------------------------------------------------------------------
# Probe: which selectors match right now
# ----------------------------------------------------------------------------------
def probe(engine, page, human, query, timeout):
    print(f"\n{'=' * 70}\n{engine.label}  ({engine.url})\n{'=' * 70}")
    try:
        if engine.kind == "search":
            answer = ask_search(engine, page, human, query, timeout)
            print(f"\n  ANSWER ({len(answer)} chars):\n  {answer[:400]}")
            return

        page.goto(engine.url, wait_until="domcontentloaded")
        settle(page)  # these sites bounce through auth/Cloudflare on load
        human.pause(1.0, 2.0)

        landed = safe(lambda: page.url, "?")
        print(f"  asked for : {engine.url}")
        print(f"  landed on : {landed}")
        print(f"  title     : {safe(lambda: page.title(), '?')[:70]!r}")

        wanted_host = engine.url.split("/")[2]
        landed_host = landed.split("/")[2] if "//" in landed else "?"
        if landed_host != wanted_host:
            print(f"  >> REDIRECTED to {landed_host} -- almost always a sign-in wall.")
            print(f"  >> Sign in at {engine.url} via: python chatgpt_query.py --login")

        for group, selectors in (("composer", engine.composer),
                                 ("answer", engine.answer),
                                 ("send", engine.send),
                                 ("login", engine.login)):
            for selector in selectors:
                total = safe(lambda s=selector: page.locator(s).count(), "err")
                shown = total
                if isinstance(total, int) and total:
                    vis = safe(
                        lambda s=selector: sum(
                            1 for i in range(min(total, 5))
                            if page.locator(s).nth(i).is_visible()),
                        "?")
                    shown = f"{total} ({vis} visible)"
                flag = "MATCH" if isinstance(total, int) and total else "     "
                print(f"  {flag} {group:<9} {selector:<48} -> {shown}")

        if first_present(page, engine.login)[0] is not None and \
                first_present(page, engine.composer)[0] is None:
            print("  >> looks SIGNED OUT in this profile: "
                  "run python chatgpt_query.py --login")
            return
        if first_present(page, engine.composer)[0] is None:
            print("  >> no visible composer. If the page looks fine in the window, the "
                  "composer selectors need updating in ENGINES.")
            return

        print("\n  sending the test query...")
        answer = ask_chat(engine, page, human, query, timeout)
        labelled = answer.startswith("[captured from page text")
        print(f"\n  ANSWER ({len(answer)} chars, "
              f"{'TEXT FALLBACK - answer selectors did not match' if labelled else 'selector path'}):")
        print(f"  {answer[:400]}")
    except Exception as exc:
        print(f"\n  FAILED: {exc.__class__.__name__}: {exc}")
        print(f"  landed on: {safe(lambda: page.url, '?')}")


def main():
    ap = argparse.ArgumentParser(description="Probe the LLM front-ends' selectors.")
    ap.add_argument("--probe", default="all", help="engine key, comma list, or all")
    ap.add_argument("--query", default="What is 2 + 2? Answer in one short sentence.")
    ap.add_argument("--timeout", type=int, default=180)
    # Browser options, same as chatgpt_query.py
    ap.add_argument("-p", "--profile", default=cgq.DEFAULT_PROFILE)
    ap.add_argument("--mode", choices=("own", "copy", "real"), default="own")
    ap.add_argument("--profile-dir", default=None)
    ap.add_argument("--refresh-copy", action="store_true")
    ap.add_argument("--chrome-path", default=None)
    ap.add_argument("--cdp-port", type=int, default=cgq.CDP_PORT)
    ap.add_argument("--force-close", action="store_true")
    ap.add_argument("--close-browser", action="store_true")
    ap.add_argument("--no-stealth", action="store_true")
    ap.add_argument("--runtime-fix", default=None,
                    choices=("addBinding", "alwaysIsolated", "off"),
                    help="how rebrowser hides the Runtime.enable leak. alwaysIsolated survives pages that redirect mid-load, such as Cloudflare interstitials; addBinding hides more but can die on them.")
    ap.add_argument("--wpm", type=int, default=72)
    ap.add_argument("--jitter", type=float, default=1.0)
    args = ap.parse_args()

    engines = resolve(args.probe)
    with cgq.sync_playwright() as p:
        page, close = cgq.open_browser(p, args)
        try:
            human = cgq.Human(page, wpm=args.wpm, jitter=args.jitter)
            cgq.fit_window(page)
            for engine in engines:
                probe(engine, page, human, args.query, args.timeout)
        finally:
            try:
                close()
            except Exception:
                pass
    print("\nFix any non-matching selectors in the ENGINES table in llm_engines.py.")


if __name__ == "__main__":
    main()
