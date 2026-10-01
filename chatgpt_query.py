"""
Ask ChatGPT in your real, already-logged-in Chrome, copy the answer, save it to a text file.

Made to look like a person using their own browser: your actual Chrome binary, your
actual profile (so the ChatGPT session is already signed in), human mouse/typing/scroll
timing, no automation switches, and Cloudflare Turnstile handling.

Setup
-----
    pip install playwright

Optional, recommended:
    pip install rebrowser-playwright   # patches the CDP Runtime.enable leak
    pip install pyperclip              # clipboard fallback + fast paste of long prompts

Usage
-----
    python chatgpt_query.py --list-profiles              # see your Chrome profiles
    python chatgpt_query.py --login                      # sign in once, unautomated
    python chatgpt_query.py "What is the capital of France?"
    python chatgpt_query.py "hi" -p ai-automation        # pick a profile by email/name
    python chatgpt_query.py "explain quicksort" -o answer.txt
    python chatgpt_query.py "one" "two" "three"          # fresh chat + file each
    python chatgpt_query.py --selftest                   # print what a detector sees

Modes (--mode)
--------------
  copy  (default)  Copies your chosen profile once into its own folder and runs your
                   real Chrome on that. Your ChatGPT login comes along, nothing has to
                   be closed, and your normal Chrome keeps running untouched.
                   --refresh-copy redoes the copy if the session ever goes stale.
  real             Your Chrome on your actual profile folder, over a debugging port.
                   Two catches: every Chrome window must be closed first, and Chrome
                   136+ ignores a debugging port on the default profile folder
                   entirely, so on a current Chrome this cannot work -- it detects that
                   and falls back to copy instead of closing anything.
  own              A dedicated empty profile; you sign in once, by hand, and it sticks.
                   Slowest to start, but it always works.

Signing in
----------
Google will not let you complete a Google sign-in inside an automation-driven browser
("Couldn't sign you in -- this browser or app may not be secure"), and that is an
account-security check, not a bot check. Two ways around it, neither of which fights it:

  * copy mode (default): the ChatGPT session comes over with the copied cookies, so no
    sign-in happens at all. This is the easy path.
  * --login: opens a plain Chrome that this script starts and never attaches to, so
    Google sees an ordinary browser. Sign in there, close it, then use --mode own. The
    session persists, so this is a one-time step.

Notes
-----
* The window is always visible. Headless Chrome is trivially detectable and is the
  single biggest giveaway, so there is no headless mode here on purpose.
* Automating ChatGPT's web UI is against OpenAI's terms of service; the API is the
  supported path. Your account, your call.
"""

import os
import sys


def _runtime_fix_mode():
    """
    How rebrowser-playwright should hide the Runtime.enable leak. It reads this at import
    time, before argparse exists, so the command line is scanned by hand here.

      addBinding      (default) hides the leak best, but has to acquire the main world
                      through a binding. On a page that redirects mid-load -- a Cloudflare
                      interstitial, say -- that can die with
                      "CRSession.__re__getMainWorld ... type: 'closed'".
      alwaysIsolated  steadier on redirecting pages; evaluates in an isolated world,
                      which a determined site can notice.
      off             plain Playwright behaviour, no patch.
    """
    for i, arg in enumerate(sys.argv):
        if arg == "--runtime-fix" and i + 1 < len(sys.argv):
            return sys.argv[i + 1]
        if arg.startswith("--runtime-fix="):
            return arg.split("=", 1)[1]
    return os.environ.get("REBROWSER_PATCHES_RUNTIME_FIX_MODE") or "addBinding"


RUNTIME_FIX = _runtime_fix_mode()
os.environ["REBROWSER_PATCHES_RUNTIME_FIX_MODE"] = (
    "0" if RUNTIME_FIX in ("off", "0", "none", "disabled") else RUNTIME_FIX)
os.environ.setdefault("REBROWSER_PATCHES_UTILITY_WORLD_NAME", "util")

import argparse
import datetime
import json
import math
import pathlib
import random
import re
import shutil
import socket
import subprocess
import tempfile
import threading
import time

try:  # patched build: hides the Runtime.enable leak advanced anti-bots look for
    from rebrowser_playwright.sync_api import TimeoutError as PWTimeout  # type: ignore
    from rebrowser_playwright.sync_api import sync_playwright  # type: ignore

    RUNTIME_PATCHED = True
except ImportError:
    from playwright.sync_api import TimeoutError as PWTimeout
    from playwright.sync_api import sync_playwright

    RUNTIME_PATCHED = False

CHATGPT_URL = "https://chatgpt.com/"
# Whose Chrome profile to drive by default; override with -p. Matches a profile folder
# name, display name or signed-in email.
DEFAULT_PROFILE = "kodexolabsinternal@gmail.com"
OWN_PROFILE = pathlib.Path.home() / ".chatgpt_playwright_profile"   # --login lives here
COPY_PROFILE = pathlib.Path.home() / ".chatgpt_playwright_copy"     # --mode copy target
CDP_PORT = 9222

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    str(pathlib.Path.home() / r"AppData\Local\Google\Chrome\Application\chrome.exe"),
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
]
USER_DATA_DIRS = [
    pathlib.Path.home() / r"AppData\Local\Google\Chrome\User Data",
    pathlib.Path.home() / "Library/Application Support/Google/Chrome",
    pathlib.Path.home() / ".config/google-chrome",
]
# Where the session cookies live inside a copied user-data folder.
COOKIE_REL = "Default/Network/Cookies"
# Caches: big, worthless in a copy, and Chrome rebuilds them.
COPY_SKIP = {
    "Cache", "Code Cache", "GPUCache", "DawnCache", "DawnGraphiteCache",
    "DawnWebGPUCache", "GrShaderCache", "ShaderCache", "Service Worker",
    "component_crx_cache", "extensions_crx_cache", "optimization_guide_model_store",
    "Crashpad", "CrashpadMetrics", "blob_storage", "File System", "Safe Browsing",
    "segmentation_platform", "BrowserMetrics", "Greaselion", "Download Service",
    # Chrome's own account plumbing. Copying it makes the new profile try to
    # re-authenticate your Google *browser* account on startup, which Google rejects in
    # an automated browser ("Couldn't sign you in") in a stray tab. Website cookies --
    # the ChatGPT session -- live in Network/Cookies and are unaffected by leaving these
    # out. "Sessions" is also skipped so the copy does not reopen your real tabs.
    "Sync Data", "Sync Extension Settings", "Sync App Settings", "Account Web Data",
    "Login Data For Account", "Sessions", "GCM Store", "Accounts",
}

# --- ChatGPT DOM (data-testids change now and then; they are all here) -------------
COMPOSER = "#prompt-textarea, div[contenteditable='true'][id='prompt-textarea']"
ASSISTANT_MSG = "[data-message-author-role='assistant']"
SEND_BUTTON = "[data-testid='send-button'], button[aria-label*='Send']"
STOP_BUTTON = "[data-testid='stop-button'], button[aria-label*='Stop']"
COPY_BUTTON = "[data-testid='copy-turn-action-button'], button[aria-label='Copy']"
LOGIN_MARKER = "[data-testid='login-button'], a[href*='/auth/login']"

# --- Cloudflare ------------------------------------------------------------------
CF_MARKERS = [
    "#challenge-running",
    "#cf-challenge-running",
    "div.cf-browser-verification",
    "iframe[src*='challenges.cloudflare.com']",
]
CF_TITLES = ("just a moment", "attention required", "verify you are human", "one more step")

# Real Chrome reports real plugins, a real GPU and a real window.chrome, so faking any
# of that makes things worse. This is only what the automation switches actually leak.
STEALTH_MINIMAL = """
Object.defineProperty(Object.getPrototypeOf(navigator), 'webdriver', {
  get: () => undefined, configurable: true,
});
if (window.navigator.permissions && window.Notification) {
  const q = window.navigator.permissions.query.bind(window.navigator.permissions);
  window.navigator.permissions.query = (p) =>
    p && p.name === 'notifications'
      ? Promise.resolve({ state: Notification.permission, name: 'notifications',
                          onchange: null, addEventListener(){}, removeEventListener(){} })
      : q(p);
}
try { delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array; } catch (e) {}
"""

# Only for Playwright's bundled Chromium: no plugins, no window.chrome and a
# SwiftShader GPU string, each a standalone tell.
STEALTH_CHROMIUM = """
if (!window.chrome) {
  window.chrome = { runtime: {}, app: { isInstalled: false },
                    csi: () => {}, loadTimes: () => {} };
}
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
(() => {
  const make = (name, filename, desc) => ({ name, filename, description: desc, length: 1 });
  const list = [
    make('PDF Viewer', 'internal-pdf-viewer', 'Portable Document Format'),
    make('Chrome PDF Viewer', 'internal-pdf-viewer', 'Portable Document Format'),
    make('Chromium PDF Viewer', 'internal-pdf-viewer', 'Portable Document Format'),
    make('WebKit built-in PDF', 'internal-pdf-viewer', 'Portable Document Format'),
  ];
  Object.defineProperty(navigator, 'plugins', {
    get: () => Object.assign(Object.create(PluginArray.prototype), list,
                             { length: list.length }),
  });
})();
(() => {
  const patch = (proto) => {
    if (!proto) return;
    const orig = proto.getParameter;
    proto.getParameter = function (p) {
      if (p === 37445) return 'Google Inc. (Intel)';
      if (p === 37446) return 'ANGLE (Intel, Intel(R) UHD Graphics 630 Direct3D11 vs_5_0 ps_5_0, D3D11)';
      return orig.apply(this, [p]);
    };
  };
  patch(window.WebGLRenderingContext && WebGLRenderingContext.prototype);
  patch(window.WebGL2RenderingContext && WebGL2RenderingContext.prototype);
})();
"""

SELFTEST_HTML = """<!doctype html><meta charset="utf-8"><title>fingerprint</title>
<pre id="out">running...</pre><script>
const gl = document.createElement('canvas').getContext('webgl');
const dbg = gl && gl.getExtension('WEBGL_debug_renderer_info');
const rows = {
  'navigator.webdriver': navigator.webdriver,
  'window.chrome': typeof window.chrome,
  'chrome.runtime': typeof (window.chrome && window.chrome.runtime),
  'plugins.length': navigator.plugins.length,
  'languages': navigator.languages.join(','),
  'userAgent': navigator.userAgent,
  'headless in UA': /headless/i.test(navigator.userAgent),
  'webgl vendor': dbg ? gl.getParameter(dbg.UNMASKED_VENDOR_WEBGL) : 'n/a',
  'webgl renderer': dbg ? gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) : 'n/a',
  'hardwareConcurrency': navigator.hardwareConcurrency,
  'deviceMemory': navigator.deviceMemory,
  'window': outerWidth + ' x ' + outerHeight,
  'screen': screen.width + ' x ' + screen.height,
  'timezone': Intl.DateTimeFormat().resolvedOptions().timeZone,
  'cdc_ vars': Object.keys(window).filter(k => k.startsWith('cdc_')).length,
  'permissions/Notification match': 'checking',
};
navigator.permissions.query({name:'notifications'}).then(r => {
  rows['permissions/Notification match'] = (r.state === Notification.permission);
}).catch(() => {}).then(() => {
  document.getElementById('out').textContent =
    Object.entries(rows).map(([k, v]) => k.padEnd(30) + ' : ' + v).join('\\n');
  window.__rows = rows;
  document.title = 'fingerprint-ready';
});
</script>"""


def log(msg):
    print("[chatgpt] " + str(msg), flush=True)


def slugify(text, limit=45):
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", text).strip("-").lower()
    return (slug[:limit] or "query").rstrip("-")


# ----------------------------------------------------------------------------------
# Finding Chrome and its profiles.  Only Local State's plain-JSON profile list is read
# here -- profile folder names, display names, signed-in emails.  No cookie or
# credential store is ever opened or decrypted by this script.
# ----------------------------------------------------------------------------------
def find_chrome():
    for path in CHROME_PATHS:
        if pathlib.Path(path).exists():
            return path
    return None


def find_user_data():
    for path in USER_DATA_DIRS:
        if path.exists():
            return path
    return None


def resolve_profile_dir(args, for_login=False):
    """
    Which profile folder this run uses. 'copy' and 'own' must never share one: a copy
    run rebuilds its folder from your Chrome profile, which would wipe the sign-ins that
    --login put in the own profile.
    """
    override = getattr(args, "profile_dir", None)
    if override:
        return pathlib.Path(override).expanduser().resolve()
    if for_login or getattr(args, "mode", "own") != "copy":
        return OWN_PROFILE.expanduser().resolve()
    return COPY_PROFILE.expanduser().resolve()


def chrome_major_version(exe=None):
    """Read the version from the folder Chrome keeps beside chrome.exe. 0 if unknown."""
    exe = exe or find_chrome()
    if not exe:
        return 0
    try:
        for child in pathlib.Path(exe).parent.iterdir():
            m = re.match(r"^(\d+)\.\d+\.\d+", child.name)
            if child.is_dir() and m:
                return int(m.group(1))
    except Exception:
        pass
    return 0


def debug_port_blocked(user_data, exe=None):
    """
    Chrome 136 and later ignore --remote-debugging-port (and --remote-debugging-pipe)
    when --user-data-dir is the default profile directory. It was tightened to stop
    other processes from reading a logged-in profile's cookies over CDP, and it cannot
    be switched off -- the port simply never opens. A separate user-data-dir is fine,
    which is what 'copy' and 'own' modes use.
    """
    major = chrome_major_version(exe)
    if major < 136:
        return False
    default = find_user_data()
    try:
        return bool(default) and pathlib.Path(user_data).resolve() == default.resolve()
    except Exception:
        return True


def list_profiles(user_data):
    """[{dir, name, email, active}] newest-used first."""
    state = user_data / "Local State"
    out, last_used = [], None
    try:
        data = json.loads(state.read_text(encoding="utf-8", errors="ignore"))
        prof = data.get("profile", {})
        last_used = prof.get("last_used")
        for folder, info in prof.get("info_cache", {}).items():
            if (user_data / folder).is_dir():
                out.append({
                    "dir": folder,
                    "name": info.get("name") or folder,
                    "email": info.get("user_name") or "",
                    "active": float(info.get("active_time") or 0),
                })
    except Exception as exc:
        log(f"Could not read the profile list ({exc.__class__.__name__}).")

    if not out:  # no Local State: fall back to folders on disk
        for p in sorted(user_data.glob("Profile *")) + [user_data / "Default"]:
            if p.is_dir():
                out.append({"dir": p.name, "name": p.name, "email": "", "active": 0.0})

    out.sort(key=lambda r: (r["dir"] == last_used, r["active"]), reverse=True)
    return out


def pick_profile(profiles, wanted):
    """Match a folder name, display name or email; case-insensitive substring."""
    if not wanted:
        return profiles[0] if profiles else None
    w = wanted.strip().lower()
    for row in profiles:
        if w == row["dir"].lower():
            return row
    for key in ("email", "name", "dir"):
        for row in profiles:
            if w in str(row[key]).lower():
                return row
    return None


def print_profiles(user_data):
    profiles = list_profiles(user_data)
    print(f"\nChrome profiles in {user_data}\n" + "-" * 66)
    for i, row in enumerate(profiles):
        mark = "  <- default pick" if i == 0 else ""
        print(f"  {row['dir']:<12} {row['name']:<22} {row['email']}{mark}")
    print("\nUse: python chatgpt_query.py \"your query\" -p \"<folder, name or email>\"\n")


def chrome_running():
    try:
        if sys.platform == "win32":
            out = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
                capture_output=True, text=True, timeout=20,
            ).stdout
            return "chrome.exe" in out.lower()
        out = subprocess.run(["pgrep", "-x", "chrome"], capture_output=True, timeout=20)
        return out.returncode == 0
    except Exception:
        return False


def close_chrome(force=False):
    if not force:
        print("\nChrome is running. It has to close for this mode -- your tabs come back\n"
              "when Chrome reopens (or use --mode copy to leave Chrome alone).")
        if input("Close Chrome now? [y/N] ").strip().lower() not in ("y", "yes"):
            return False
    log("Closing Chrome...")
    try:
        if sys.platform == "win32":
            subprocess.run(["taskkill", "/IM", "chrome.exe", "/F"],
                           capture_output=True, timeout=30)
        else:
            subprocess.run(["pkill", "-x", "chrome"], capture_output=True, timeout=30)
    except Exception as exc:
        log(f"Could not close Chrome: {exc}")
        return False
    time.sleep(2.5)
    return not chrome_running()


def _copy_tree(src, dst, failures):
    """
    Copy a tree file by file. shutil.copytree gives up on a whole directory when one
    file in it is locked, which silently loses Network/Cookies -- the one file that
    decides whether the copy is signed in -- because Chrome holds it open.
    """
    dst.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        if item.name in COPY_SKIP or item.name.endswith(("-journal", ".log")):
            continue
        out = dst / item.name
        try:
            if item.is_dir():
                _copy_tree(item, out, failures)
            else:
                shutil.copy2(item, out)
        except Exception as exc:
            failures.append((item.name, exc.__class__.__name__))


def rewrite_local_state(dest, profile_dir):
    """
    A copied Local State still lists every profile from the original Chrome and points
    last_used at a folder ("Profile 17") that does not exist inside the copy, so Chrome
    opens the profile picker and waits for a choice instead of just opening the profile.
    Rewrite it to describe exactly one profile, Default, and keep everything else --
    notably os_crypt, which the copied cookies need in order to decrypt -- untouched.
    """
    path = pathlib.Path(dest) / "Local State"
    if not path.is_file():
        return False
    try:
        data = json.loads(path.read_text(encoding="utf-8", errors="ignore"))
        profile = data.setdefault("profile", {})
        info = profile.get("info_cache") or {}
        mine = dict(info.get(profile_dir) or next(iter(info.values()), {}) or {})
        mine.pop("is_ephemeral", None)
        profile["info_cache"] = {"Default": mine}
        profile["last_used"] = "Default"
        profile["profiles_order"] = ["Default"]
        profile["last_active_profiles"] = ["Default"]
        data.setdefault("browser", {})["show_profile_picker_on_startup"] = False
        path.write_text(json.dumps(data), encoding="utf-8")
        return True
    except Exception as exc:
        log(f"Could not rewrite Local State ({exc.__class__.__name__}); Chrome may show "
            "the profile picker.")
        return False


def strip_browser_signin(dest):
    """
    Take Chrome's own Google account out of the copied profile, and stop it restoring the
    original profile's tabs. Without this, Chrome opens a second tab trying to re-auth
    the browser account and Google answers "This browser or app may not be secure",
    because the browser is automation-driven. None of this touches website cookies, so
    the ChatGPT session is unaffected -- only Chrome's own sign-in goes away.
    """
    prefs_path = pathlib.Path(dest) / "Default" / "Preferences"
    try:
        prefs = json.loads(prefs_path.read_text(encoding="utf-8", errors="ignore"))
        for key in ("account_info", "signin", "sync", "gaia_cookie",
                    "account_id_migration_state", "account_tracker_service_last_update",
                    "token_service"):
            prefs.pop(key, None)
        session = prefs.setdefault("session", {})
        session["restore_on_startup"] = 5  # open a new tab, not the old session
        session["startup_urls"] = []
        prefs_path.write_text(json.dumps(prefs), encoding="utf-8")
    except Exception as exc:
        log(f"Could not strip the browser sign-in from Preferences "
            f"({exc.__class__.__name__}).")

    # The same account is mirrored in Local State's profile entry.
    state_path = pathlib.Path(dest) / "Local State"
    try:
        data = json.loads(state_path.read_text(encoding="utf-8", errors="ignore"))
        entry = data.get("profile", {}).get("info_cache", {}).get("Default")
        if entry:
            for key in list(entry):
                if key.startswith("gaia") or key in (
                    "user_name", "hosted_domain", "is_consented_primary_account",
                    "signin_required", "managed_user_id",
                ):
                    entry.pop(key, None)
            state_path.write_text(json.dumps(data), encoding="utf-8")
    except Exception:
        pass


def cookies_present(dest):
    """Did the session cookies actually make it into the copy?"""
    db = pathlib.Path(dest) / COOKIE_REL
    try:
        return db.is_file() and db.stat().st_size > 4096
    except Exception:
        return False


def start_chrome_normally(exe=None):
    """Put the user's own Chrome back, exactly as they would launch it."""
    exe = exe or find_chrome()
    if not exe:
        return False
    try:
        subprocess.Popen([exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


def copy_profile(user_data, profile_dir, dest, refresh=False):
    """
    Copy one profile into its own user-data folder so this can run beside a live Chrome.
    Login survives because the cookie encryption key in Local State is copied with it and
    decrypts under the same Windows account.
    """
    # Absolute, always: Chrome resolves a relative --user-data-dir against its own
    # working directory, decides the dir is one it already has open, and quietly hands
    # the launch to the running browser ("Opening in existing browser session").
    dest = pathlib.Path(dest).expanduser().resolve()
    marker = dest / ".copied-from"
    if refresh and dest.exists():
        log(f"Refreshing the profile copy at {dest}")
        shutil.rmtree(dest, ignore_errors=True)
    if marker.exists() and marker.read_text(encoding="utf-8").strip() == profile_dir:
        log(f"Reusing the existing copy of {profile_dir} at {dest}")
        rewrite_local_state(dest, profile_dir)  # also repairs copies made earlier
        strip_browser_signin(dest)
        return dest

    src = user_data / profile_dir
    if not src.is_dir():
        sys.exit(f"Profile folder not found: {src}")
    if chrome_running():
        log("Chrome is running; some files are locked and will be skipped. If the copy "
            "ends up signed out, close Chrome and delete " + str(dest) + " to redo it.")

    log(f"Copying {profile_dir} -> {dest} (one time, skipping caches)...")
    dest.mkdir(parents=True, exist_ok=True)
    for name in ("Local State", "First Run"):
        try:
            if (user_data / name).exists():
                shutil.copy2(user_data / name, dest / name)
        except Exception:
            pass

    target = dest / "Default"  # Chrome uses Default inside a fresh user-data dir
    failures = []
    _copy_tree(src, target, failures)
    # A copied profile looks to Chrome like one that crashed, which makes it open with a
    # "Restore pages?" bubble. Clear that: no normal session starts with one.
    prefs = target / "Preferences"
    try:
        data = json.loads(prefs.read_text(encoding="utf-8", errors="ignore"))
        data.setdefault("profile", {})
        data["profile"]["exit_type"] = "Normal"
        data["profile"]["exited_cleanly"] = True
        prefs.write_text(json.dumps(data), encoding="utf-8")
    except Exception:
        pass

    rewrite_local_state(dest, profile_dir)
    strip_browser_signin(dest)

    marker.write_text(profile_dir, encoding="utf-8")
    if failures:
        names = ", ".join(sorted({n for n, _ in failures})[:6])
        log(f"{len(failures)} file(s) were locked and skipped: {names}")
    if not cookies_present(dest):
        log("The session cookies were NOT copied, so this copy is signed out.")
    return dest


def interactive_login(args):
    """
    Google refuses OAuth sign-in in a browser it considers automated ("Couldn't sign you
    in -- this browser or app may not be secure"), because CDP is attached to it. Rather
    than fight an account-security check, do the sign-in in a plain Chrome that this
    script starts and never attaches to: no debugging port, no automation flags, nothing
    to detect, because it genuinely is not automated. The cookies it leaves in the
    profile are what the automated run uses afterwards, and they last for months.
    """
    exe = args.chrome_path or find_chrome()
    if not exe:
        sys.exit("Chrome not found; pass --chrome-path.")
    profile = resolve_profile_dir(args, for_login=True)
    profile.mkdir(parents=True, exist_ok=True)

    sites = ["https://chatgpt.com/", "https://claude.ai/", "https://grok.com/",
             "https://www.google.com/"]
    print("\n" + "=" * 70)
    print("Opening a normal Chrome window on the automation profile:")
    print(f"  {profile}")
    print("\n  1. Sign in to each site in the tabs that open -- ChatGPT, Claude and\n"
          "     Grok. Google buttons work here: this window is a plain browser, so\n"
          "     Google has nothing to object to.\n"
          "  2. Once each one is open and working, CLOSE that Chrome window.\n"
          "  3. Come back here and press Enter.")
    print("\n  Only the engines you sign in to can be used; the rest are skipped.")
    print("=" * 70 + "\n")

    subprocess.Popen(
        [exe, f"--user-data-dir={profile}", "--no-first-run",
         "--no-default-browser-check"] + sites,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    input("Press Enter once you have signed in and closed that window... ")

    if cookies_present(profile):
        log("Profile is signed in and ready.")
        print("\nCheck each engine answers, then start filling the tracker:\n"
              "  python llm_engines.py --probe all --mode own\n"
              "  python run_tracker.py --engines all --sheet Hospitals --limit 4 "
              "--mode own\n")
    else:
        log("No cookies landed in that profile. Did the window open on the path above, "
            "and did you close it before pressing Enter?")


def recopy_with_chrome_closed(user_data, chosen, args):
    """
    Chrome keeps the cookie database open while it runs, so a copy taken alongside it
    comes out signed out. Closing Chrome for the two or three seconds the copy takes
    fixes that once; after it, the copy lives in its own user-data folder, so Chrome
    comes straight back and the two run side by side from then on.
    """
    fallback = resolve_profile_dir(args)
    print("\nChrome has the cookie file open, so the copy carries no ChatGPT session.\n"
          "Closing Chrome for ~3 seconds fixes it for good -- I reopen it immediately\n"
          "after, and from then on this runs beside your normal browsing.")
    if not args.force_close:
        answer = input("Close Chrome briefly and copy again? [y/N] ").strip().lower()
        if answer not in ("y", "yes"):
            log("Keeping the signed-out copy: sign in in the window that opens (it will "
                "stick), or use --mode own.")
            return fallback

    if not close_chrome(force=True):
        log("Could not close Chrome; continuing with the signed-out copy.")
        return fallback

    dest = copy_profile(user_data, chosen["dir"], resolve_profile_dir(args), refresh=True)
    if start_chrome_normally():
        log("Your Chrome is reopening now.")
    time.sleep(1.5)
    if cookies_present(dest):
        log("Session cookies copied; this profile is signed in.")
    return dest


# ----------------------------------------------------------------------------------
# Human-ish input.  Anti-bot scoring leans on event timing as much as on JS properties:
# instant clicks at exact element centres, zero-delay keystrokes and teleporting
# cursors are what get flagged.
# ----------------------------------------------------------------------------------
class Human:
    def __init__(self, page, wpm=72, jitter=1.0):
        self.page = page
        self.wpm = wpm
        self.jitter = jitter
        self.x, self.y = random.uniform(200, 700), random.uniform(200, 500)

    def pause(self, low, high=None):
        high = low * 1.8 if high is None else high
        time.sleep(random.uniform(low, high) * self.jitter)

    @staticmethod
    def _ease(t):
        return 3 * t * t - 2 * t * t * t  # smoothstep: slow start, slow stop

    def move(self, tx, ty, steps=None):
        """Cubic Bezier glide with a slight overshoot, like a hand on a mouse."""
        sx, sy = self.x, self.y
        dist = math.hypot(tx - sx, ty - sy)
        steps = steps or max(12, min(45, int(dist / 12) + random.randint(4, 12)))
        bow = min(120, dist * random.uniform(0.08, 0.28))
        c1 = (sx + (tx - sx) * 0.3 + random.uniform(-bow, bow),
              sy + (ty - sy) * 0.3 + random.uniform(-bow, bow))
        c2 = (sx + (tx - sx) * 0.7 + random.uniform(-bow, bow),
              sy + (ty - sy) * 0.7 + random.uniform(-bow, bow))
        over = (tx + random.uniform(-6, 6), ty + random.uniform(-5, 5))

        for i in range(1, steps + 1):
            t = self._ease(i / steps)
            u = 1 - t
            px = u**3 * sx + 3 * u**2 * t * c1[0] + 3 * u * t**2 * c2[0] + t**3 * over[0]
            py = u**3 * sy + 3 * u**2 * t * c1[1] + 3 * u * t**2 * c2[1] + t**3 * over[1]
            self.page.mouse.move(px, py)
            time.sleep(random.uniform(0.004, 0.016) * self.jitter)

        if math.hypot(over[0] - tx, over[1] - ty) > 2:  # settle onto the target
            self.page.mouse.move(tx, ty)
            time.sleep(random.uniform(0.02, 0.07))
        self.x, self.y = tx, ty

    def click(self, locator, timeout=30_000):
        locator.wait_for(state="visible", timeout=timeout)
        locator.scroll_into_view_if_needed()
        box = locator.bounding_box()
        if not box:
            locator.click()
            return
        # inside the element, biased to the middle but never dead centre
        tx = box["x"] + box["width"] * random.triangular(0.25, 0.75, 0.5)
        ty = box["y"] + box["height"] * random.triangular(0.25, 0.75, 0.5)
        self.move(tx, ty)
        self.pause(0.06, 0.22)
        self.page.mouse.down()
        time.sleep(random.uniform(0.045, 0.14))
        self.page.mouse.up()

    def _char_delay(self, ch, prev):
        base = 60.0 / (self.wpm * 5)  # seconds per character
        d = random.gauss(base, base * 0.42)
        if ch == " ":
            d *= 1.25
        if prev and prev in ".,!?;:":
            d *= 2.1
        if ch.isupper() or ch in "()[]{}\"'/\\@#$%^&*_|<>~`":
            d *= 1.5  # reaching for shift
        return max(0.022, d) * self.jitter

    def type(self, text):
        kb = self.page.keyboard
        prev = ""
        for i, ch in enumerate(text):
            if random.random() < 0.012 and ch.isalpha():  # typo, noticed, fixed
                kb.type(random.choice("asdfghjkl"))
                time.sleep(random.uniform(0.09, 0.3))
                kb.press("Backspace")
                time.sleep(random.uniform(0.05, 0.18))
            kb.type(ch)
            time.sleep(self._char_delay(ch, prev))
            prev = ch
            if ch == " " and random.random() < 0.05:  # mid-sentence think
                time.sleep(random.uniform(0.25, 1.1) * self.jitter)
            if i and i % random.randint(45, 90) == 0:
                time.sleep(random.uniform(0.2, 0.7) * self.jitter)

    def paste(self, text):
        """What a person does with a long prompt: paste it."""
        try:
            import pyperclip

            pyperclip.copy(text)
        except Exception:
            return False
        self.pause(0.2, 0.6)
        self.page.keyboard.press("Meta+V" if sys.platform == "darwin" else "Control+V")
        self.pause(0.3, 0.8)
        return True

    def scroll(self, amount, bursts=None):
        bursts = bursts or random.randint(2, 5)
        per = amount / bursts
        for _ in range(bursts):
            self.page.mouse.wheel(0, per * random.uniform(0.7, 1.3))
            time.sleep(random.uniform(0.12, 0.45) * self.jitter)

    def wander(self):
        """Idle cursor drift while a page settles."""
        for _ in range(random.randint(1, 3)):
            self.move(random.uniform(120, 1100), random.uniform(120, 700))
            self.pause(0.15, 0.6)

    def read(self, text):
        """Dwell for roughly as long as reading the answer would take."""
        words = max(1, len(text.split()))
        seconds = min(9.0, 0.8 + words / random.uniform(380, 700))
        end = time.time() + seconds
        while time.time() < end:
            self.scroll(random.uniform(180, 520), bursts=random.randint(1, 3))
            time.sleep(random.uniform(0.3, 0.9))


# ----------------------------------------------------------------------------------
# Browser
# ----------------------------------------------------------------------------------
def launch_args():
    # A plausible window on a normal desktop, slightly different every run.
    w = random.choice([1280, 1366, 1440, 1512, 1600])
    h = random.choice([760, 800, 864, 900])
    # --disable-blink-features=AutomationControlled is deliberately NOT passed. Current
    # Chrome answers it with a visible "unsupported command-line flag" infobar, which is
    # a louder tell than the thing it used to hide -- and it hides nothing here anyway,
    # since navigator.webdriver is already handled by dropping --enable-automation and by
    # the init script.
    return [
        "--profile-directory=Default",  # never stop at the profile picker

        f"--window-size={w},{h}",
        f"--window-position={random.randint(0, 120)},{random.randint(0, 80)}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-features=Translate",
        "--disable-sync",  # never try to re-auth Chrome's own Google account
    ]


# Playwright's own flags that announce automation or differ from a stock browser.
IGNORE_ARGS = [
    "--enable-automation",
    "--disable-extensions",
    "--disable-component-extensions-with-background-pages",
    "--disable-default-apps",
    "--disable-popup-blocking",
    "--disable-back-forward-cache",
    "--disable-client-side-phishing-detection",
    "--enable-blink-features=IdleDetection",
    "--no-sandbox",
]


def safe_url(page):
    try:
        return page.url or ""
    except Exception:
        return ""

def call_bounded(fn, timeout=8.0, default=None):
    """
    Run a Playwright call on a worker thread and give up when it overruns.

    Anything that needs an execution context -- evaluate(), title(), locator.count() --
    takes no timeout argument, so a context that never resolves blocks the whole run
    forever with no traceback and no CPU burn. That is what --runtime-fix alwaysIsolated
    does on about:blank and on Cloudflare interstitials, where there is no settled
    document to build an isolated world in. The thread is left to finish on its own; it
    is a daemon, so it cannot hold the process open.
    """
    result = {}

    def work():
        try:
            result["value"] = fn()
        except Exception as exc:
            result["error"] = exc

    thread = threading.Thread(target=work, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive() or "value" not in result:
        return default
    return result["value"]


def evaluate_bounded(page, script, timeout=8.0, default=None):
    """page.evaluate() with a deadline. See call_bounded."""
    return call_bounded(lambda: page.evaluate(script), timeout, default)


def fit_window(page):
    """
    Keep the window inside the real desktop. Chrome honours --window-size even when it
    is bigger than the screen, and a window wider than the desktop it sits on is the
    kind of small inconsistency fingerprinters collect.
    """
    try:
        if str(safe_url(page)).startswith("about:"):
            return  # nothing to fit yet, and isolated worlds hang on about:blank
        d = evaluate_bounded(page, "({aw: screen.availWidth, ah: screen.availHeight,"
                                   "  ow: outerWidth, oh: outerHeight})")
        if not d:
            return
        aw, ah = int(d.get("aw") or 0), int(d.get("ah") or 0)
        ow, oh = int(d.get("ow") or 0), int(d.get("oh") or 0)
        if not (aw and ah) or (ow and oh and ow <= aw and oh <= ah):
            return  # already sane, or the window is not reporting yet
        w = max(900, min(ow or aw, aw) - random.randint(0, 60))
        h = max(620, min(oh or ah, ah) - random.randint(0, 40))
        cdp = page.context.new_cdp_session(page)
        window_id = cdp.send("Browser.getWindowForTarget")["windowId"]
        cdp.send("Browser.setWindowBounds", {
            "windowId": window_id,
            "bounds": {
                "left": random.randint(0, max(0, aw - w)),
                "top": random.randint(0, max(0, ah - h)),
                "width": w, "height": h,
            },
        })
        cdp.detach()
    except Exception:
        pass  # cosmetic; never worth failing a run over


def close_stray_tabs(ctx, keep):
    """
    Drop any tab the profile opened by itself -- a restored page, a sign-in promo, or
    Chrome's rejected account re-auth -- so the run starts on one clean tab.
    """
    for page in list(ctx.pages):
        if page is keep:
            continue
        try:
            url = page.url or ""
            log(f"Closing a tab Chrome opened on its own: {url[:70]}")
            page.close()
        except Exception:
            pass


def port_open(port, host="127.0.0.1"):
    with socket.socket() as s:
        s.settimeout(0.6)
        return s.connect_ex((host, port)) == 0


def open_browser(p, args):
    """Returns (page, close_fn). page is a new tab, signed in where possible."""
    if args.mode == "real":
        return open_real_chrome(p, args)
    return open_persistent(p, args)


def open_real_chrome(p, args):
    """
    Your own Chrome, your own profile, driven over a debugging port. The fingerprint is a
    genuine browser -- real GPU, real extensions, real history -- and the ChatGPT session
    is whatever you are already signed into.
    """
    exe = find_chrome()
    user_data = find_user_data()
    if not exe:
        sys.exit("Chrome not found. Use --mode own, or pass --chrome-path.")
    if not user_data:
        sys.exit("Chrome's User Data folder not found. Use --mode own.")

    if port_open(args.cdp_port):
        log(f"Found Chrome already listening on {args.cdp_port}; attaching to it.")
    elif debug_port_blocked(user_data, exe):
        # Do not close anyone's browser over this: the port would never open anyway.
        log(f"Chrome {chrome_major_version(exe)} refuses a debugging port on the default "
            "profile folder (a security change in Chrome 136), so 'real' mode cannot "
            "work here.")
        log("Switching to 'copy' mode: same profile, its own folder, your Chrome stays "
            "open. Pass --mode copy to skip this notice.")
        args.mode = "copy"
        return open_persistent(p, args)
    else:
        profiles = list_profiles(user_data)
        chosen = pick_profile(profiles, args.profile)
        if not chosen:
            print_profiles(user_data)
            sys.exit(f"No profile matched {args.profile!r}.")
        label = f"{chosen['dir']} ({chosen['email'] or chosen['name']})"

        if chrome_running():
            log("Chrome is already running, so it would ignore the debugging port.")
            if not close_chrome(force=args.force_close):
                sys.exit("Chrome is still running. Close it, or use --mode copy to run "
                         "alongside it.")

        log(f"Starting your Chrome with profile {label}")
        cmd = [
            exe,
            f"--remote-debugging-port={args.cdp_port}",
            f"--user-data-dir={user_data}",
            f"--profile-directory={chosen['dir']}",
            "--no-first-run",
            "--no-default-browser-check",
            "about:blank",
        ]
        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        deadline = time.time() + 40
        while time.time() < deadline and not port_open(args.cdp_port):
            time.sleep(0.5)
        if not port_open(args.cdp_port):
            sys.exit("Chrome never opened the debugging port. Close every Chrome window "
                     "and try again, or use --mode copy.")

    browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{args.cdp_port}")
    ctx = browser.contexts[0] if browser.contexts else browser.new_context()
    if not args.no_stealth:
        try:
            ctx.add_init_script(STEALTH_MINIMAL)
        except Exception:
            pass
    page = ctx.new_page()  # a genuinely new tab in your own window
    log("Driving your real Chrome"
        + (" (Runtime.enable patched)" if RUNTIME_PATCHED else ""))

    def close():
        # Leave the browser open -- it is yours. Just drop our tab.
        try:
            if args.close_browser:
                browser.close()
            else:
                page.close()
        except Exception:
            pass

    return page, close


def chrome_holding(profile_dir):
    """
    PIDs of Chrome processes already running on this user-data-dir.

    Chrome allows one instance per profile folder: a second launch sees the singleton
    lock, hands its URL to the first, and exits. Playwright then loses the process it
    started and reports "Target page, context or browser has been closed" -- a traceback
    that says nothing about the real cause, which is simply a browser left open.
    """
    pids = []
    needle = str(profile_dir).replace("/", "\\").lower()
    try:
        if sys.platform == "win32":
            out = subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
                 "ForEach-Object { \"$($_.ProcessId)|$($_.CommandLine)\" }"],
                capture_output=True, text=True, timeout=30).stdout
        else:
            out = subprocess.run(["ps", "-eo", "pid=,args="],
                                 capture_output=True, text=True, timeout=30).stdout
        for line in out.splitlines():
            line = line.strip()
            if not line or needle not in line.replace("/", "\\").lower():
                continue
            head = line.split("|", 1)[0] if "|" in line else line.split(None, 1)[0]
            if head.isdigit():
                pids.append(int(head))
    except Exception:
        pass
    return pids


def free_profile(profile_dir, force=True):
    """
    Close any Chrome squatting on the profile, so the launch below is not handed off.

    A window the person is using is worth asking about; an orphan with no window is not,
    and orphans are what a killed run leaves behind every time.
    """
    pids = chrome_holding(profile_dir)
    if not pids:
        return True
    log(f"Chrome is already open on {profile_dir} ({len(pids)} process(es)); "
        "closing it, or the launch would be handed to it and fail.")
    for pid in pids:
        try:
            if sys.platform == "win32":
                subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                               capture_output=True, timeout=20)
            else:
                os.kill(pid, 9)
        except Exception:
            pass
    time.sleep(2.5)
    left = chrome_holding(profile_dir)
    if left:
        log(f"  {len(left)} Chrome process(es) still hold it; close that window by hand.")
        return False
    log("  profile is free.")
    return True


def open_persistent(p, args):
    """--mode copy (a copy of your profile) or --mode own (sign in once by hand)."""
    exe = find_chrome() if args.chrome_path is None else args.chrome_path
    real_chrome = bool(exe)

    if args.mode == "copy":
        user_data = find_user_data()
        if not user_data:
            sys.exit("Chrome's User Data folder not found; use --mode own.")
        chosen = pick_profile(list_profiles(user_data), args.profile)
        if not chosen:
            print_profiles(user_data)
            sys.exit(f"No profile matched {args.profile!r}.")
        profile_dir = copy_profile(user_data, chosen["dir"], resolve_profile_dir(args),
                                   refresh=args.refresh_copy)
        if not cookies_present(profile_dir) and chrome_running():
            profile_dir = recopy_with_chrome_closed(user_data, chosen, args)
        log(f"Using a copy of {chosen['dir']} ({chosen['email'] or chosen['name']})")
    else:
        profile_dir = resolve_profile_dir(args)

    launch = dict(
        user_data_dir=str(profile_dir),
        headless=False,  # deliberately not an option
        args=launch_args(),
        ignore_default_args=IGNORE_ARGS,
        no_viewport=True,
    )
    if real_chrome:
        launch["executable_path"] = exe
    else:
        log("Real Chrome not found; using bundled Chromium (weaker fingerprint).")

    free_profile(profile_dir)
    ctx = p.chromium.launch_persistent_context(**launch)
    if not args.no_stealth:
        ctx.add_init_script(
            STEALTH_MINIMAL if real_chrome else STEALTH_MINIMAL + STEALTH_CHROMIUM
        )
    try:
        ctx.grant_permissions(["clipboard-read", "clipboard-write"],
                              origin="https://chatgpt.com")
    except Exception:
        pass

    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    close_stray_tabs(ctx, page)
    log(("Launched your Chrome" if real_chrome else "Launched Chromium")
        + f" with {profile_dir}"
        + (" (Runtime.enable patched)" if RUNTIME_PATCHED else ""))
    return page, ctx.close


# ----------------------------------------------------------------------------------
# Cloudflare
# ----------------------------------------------------------------------------------
def on_challenge(page):
    """
    Is a Cloudflare interstitial on screen?

    Every call in here needs an execution context, and on a challenge page that context
    is exactly what may never settle -- so each one gets a deadline. Treating a timeout
    as "no challenge" is the safe default: the caller then proceeds and fails loudly on
    the real page instead of freezing with no output.
    """
    try:
        title = call_bounded(lambda: page.title(), timeout=6.0, default="") or ""
        if any(t in title.lower() for t in CF_TITLES):
            return True
        for selector in CF_MARKERS:
            if call_bounded(lambda sel=selector: page.locator(sel).count(),
                            timeout=4.0, default=0):
                return True
        return False
    except Exception:
        return False


def click_turnstile(page, human):
    """
    Click the 'Verify you are human' box. The widget lives in a cross-origin iframe, so
    the locator route often fails and the reliable fallback is a real mouse click at the
    checkbox's usual position inside the iframe's box.
    """
    frame_el = page.locator("iframe[src*='challenges.cloudflare.com']").first
    try:
        if not frame_el.count():
            # Some interstitials render the checkbox straight into the page.
            inline = page.locator("input[type='checkbox']").first
            if inline.count():
                human.click(inline, timeout=6_000)
                return True
            return False

        box = frame_el.bounding_box()
        if not box or box["width"] < 20:
            return False

        human.wander()
        human.pause(0.8, 2.0)  # people do not click the instant it appears
        try:
            frame = page.frame_locator("iframe[src*='challenges.cloudflare.com']").first
            human.click(frame.locator("input[type='checkbox'], label").first, timeout=5_000)
            return True
        except Exception:
            human.move(box["x"] + min(34, box["width"] * 0.08),
                       box["y"] + box["height"] * 0.5)
            page.mouse.down()
            time.sleep(random.uniform(0.05, 0.12))
            page.mouse.up()
            return True
    except Exception:
        return False


def clear_challenge(page, human, timeout=150, done=None, attempts=3, manual=False):
    """
    Sit through a Cloudflare interstitial and click Turnstile if it asks, then confirm we
    actually got through. Returns True if clear, False if it never cleared -- it does not
    hang and it never spams reloads.

    `done` is the selector list that means "we are on the real page now"; it must come
    from the engine being used, since waiting for ChatGPT's composer on claude.ai would
    simply never finish. `manual=True` stops and lets a person solve it by hand.
    """
    if not on_challenge(page):
        return True

    done = list(done or [COMPOSER])
    log(f"Cloudflare check on {page.url.split('/')[2] if '//' in page.url else page.url}; "
        "waiting it out...")
    deadline = time.time() + timeout
    clicks = 0

    while time.time() < deadline:
        human.pause(1.0, 2.4)
        if not on_challenge(page):
            log("Cloudflare check passed.")
            human.pause(0.6, 1.5)
            return True
        for selector in done:  # sometimes the page is through before the marker clears
            try:
                if page.locator(selector).count():
                    return True
            except Exception:
                pass

        if clicks < attempts and click_turnstile(page, human):
            clicks += 1
            log(f"Clicked the verification checkbox ({clicks}/{attempts}).")
            human.pause(3.0, 6.0)  # give the redirect a chance
        else:
            human.wander()

    if manual and sys.stdin and sys.stdin.isatty():
        log("Could not clear it. Solve the check in the browser window, then press Enter.")
        try:
            input("  [Enter] once the page has loaded... ")
        except Exception:
            pass
        return not on_challenge(page)

    log("Cloudflare did not let us through. Leaving this query for a later run.")
    return False


# ----------------------------------------------------------------------------------
# ChatGPT flow
# ----------------------------------------------------------------------------------
def visit(page, human, url=CHATGPT_URL):
    page.goto(url, wait_until="domcontentloaded")
    human.pause(0.7, 2.0)  # land, look at the page
    human.wander()
    clear_challenge(page, human)


def wait_for_login(page, human, timeout=300):
    try:
        page.wait_for_selector(COMPOSER, timeout=20_000)
        human.pause(0.5, 1.4)
        return
    except PWTimeout:
        pass

    if page.locator(LOGIN_MARKER).count():
        log("This profile is signed out of ChatGPT. Sign in in the open window; I'll wait.")
        log("If Google answers \"this browser or app may not be secure\", stop here and "
            "run: python chatgpt_query.py --login")
    else:
        log("Waiting for the composer (sign in or finish a check if prompted)...")
    page.wait_for_selector(COMPOSER, timeout=timeout * 1000)
    log("Signed in.")
    human.pause(1.0, 2.5)


def send_query(page, human, query):
    composer = page.locator(COMPOSER).first
    human.click(composer)
    human.pause(0.25, 0.7)

    page.keyboard.press("Meta+A" if sys.platform == "darwin" else "Control+A")
    page.keyboard.press("Backspace")
    human.pause(0.1, 0.35)

    if len(query) > 280 and human.paste(query):
        log(f"Pasted a {len(query)}-char prompt.")
    else:
        human.type(query)

    human.pause(0.4, 1.3)  # re-read before sending
    send = page.locator(SEND_BUTTON).first
    if send.count() and send.is_enabled():
        human.click(send)
    else:
        page.keyboard.press("Enter")
    log("Sent: " + query[:90] + ("..." if len(query) > 90 else ""))


def wait_for_answer(page, answers_before, quiet_seconds=2.5, timeout=240):
    """A new assistant turn appears, then its text stops changing."""
    deadline = time.time() + timeout

    while time.time() < deadline:
        if page.locator(ASSISTANT_MSG).count() > answers_before:
            break
        time.sleep(random.uniform(0.3, 0.7))
    else:
        raise TimeoutError("ChatGPT never started answering.")

    last_text, stable_since = None, None
    while time.time() < deadline:
        text = page.locator(ASSISTANT_MSG).last.inner_text()
        streaming = page.locator(STOP_BUTTON).count() > 0

        if text and text == last_text and not streaming:
            if stable_since is None:
                stable_since = time.time()
            elif time.time() - stable_since >= quiet_seconds:
                return text
        else:
            stable_since = None
            last_text = text
        time.sleep(random.uniform(0.4, 0.8))

    raise TimeoutError("ChatGPT's answer did not finish in time.")


def copy_to_clipboard(page, human, text):
    """ChatGPT's own Copy button, so the answer lands on the real clipboard."""
    try:
        turn = page.locator(ASSISTANT_MSG).last
        turn.scroll_into_view_if_needed()
        box = turn.bounding_box()
        if box:
            human.move(box["x"] + box["width"] * 0.4, box["y"] + box["height"] - 12)
            human.pause(0.2, 0.6)
        btn = page.locator(COPY_BUTTON).last
        if btn.count():
            human.click(btn, timeout=6_000)
            log("Answer copied to clipboard (Copy button).")
            return True
    except Exception as exc:
        log("Copy button failed (" + exc.__class__.__name__ + "); trying pyperclip.")

    try:
        import pyperclip

        pyperclip.copy(text)
        log("Answer copied to clipboard (pyperclip).")
        return True
    except Exception:
        log("Could not copy to clipboard (pip install pyperclip for the fallback).")
        return False


def save(query, answer, out_path, url):
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    header = (
        "Query:  " + query + "\n"
        "Asked:  " + stamp + "\n"
        "Chat:   " + url + "\n" + "-" * 60 + "\n\n"
    )
    out_path.write_text(header + answer + "\n", encoding="utf-8")
    log("Saved to " + str(out_path))


def selftest(p, args):
    """Load a local page and print what a detection script would see."""
    page, close = open_browser(p, args)
    try:
        tmp = pathlib.Path(tempfile.gettempdir()) / "cgq_fingerprint.html"
        tmp.write_text(SELFTEST_HTML, encoding="utf-8")
        page.goto(tmp.as_uri(), wait_until="load")
        page.wait_for_function("document.title === 'fingerprint-ready'", timeout=10_000)
        fit_window(page)
        time.sleep(1.2)  # a brand-new window reports 0x0 until it is actually shown

        rows = page.evaluate("window.__rows")
        live = page.evaluate("({win: outerWidth + ' x ' + outerHeight,"
                             "  inner: innerWidth + ' x ' + innerHeight})")
        rows["window"] = live["win"]
        rows["inner window"] = live["inner"]

        print("\n--- what the page sees " + "-" * 40)
        for key, value in rows.items():
            print(f"  {key:<30} : {value}")
        print("-" * 63)

        tells = []
        if rows.get("navigator.webdriver") not in (None, False):
            tells.append("navigator.webdriver is exposed")
        if rows.get("headless in UA"):
            tells.append("UA says headless")
        if not rows.get("plugins.length"):
            tells.append("no plugins")
        if rows.get("cdc_ vars"):
            tells.append("cdc_ variables present")
        if "swiftshader" in str(rows.get("webgl renderer", "")).lower():
            tells.append("software renderer (bundled-Chromium tell)")
        if rows.get("permissions/Notification match") is False:
            tells.append("permissions/Notification mismatch")
        win = str(rows.get("window", ""))
        if win.startswith("0 ") or win == "0 x 0":
            tells.append("window reports 0x0")
        print("PASS: no obvious automation tells." if not tells
              else "TELLS: " + "; ".join(tells))
        print("Cross-check in the same window at https://bot.sannysoft.com or "
              "https://abrahamjuliot.github.io/creepjs/\n")
        Human(page, args.wpm, args.jitter).pause(1.0, 2.0)
    finally:
        close()


def main():
    ap = argparse.ArgumentParser(
        description="Ask ChatGPT in your real logged-in Chrome and save the answer.")
    ap.add_argument("query", nargs="*", help="one or more queries (each gets its own file)")
    ap.add_argument("-o", "--output", help="output file (single query only)")
    ap.add_argument("-d", "--outdir", default="chatgpt_answers", help="folder for auto-named files")
    ap.add_argument("-p", "--profile", default=DEFAULT_PROFILE,
                    help=f"Chrome profile: folder, display name or email "
                         f"(default: {DEFAULT_PROFILE})")
    ap.add_argument("--list-profiles", action="store_true", help="show your Chrome profiles")
    ap.add_argument("--login", action="store_true",
                    help="sign in once in a plain, unautomated Chrome (use this if "
                         "Google says the browser may not be secure)")
    ap.add_argument("--mode", choices=("own", "copy", "real"), default="own",
                    help="own (default): the dedicated profile you signed into with "
                         "--login; copy: your Chrome profile copied to its own folder, "
                         "runs beside your open Chrome; real: your Chrome and profile "
                         "directly, needs Chrome closed and only works below Chrome 136")
    ap.add_argument("--profile-dir", default=None,
                    help=f"profile folder (default: {OWN_PROFILE} for own and --login, "
                         f"{COPY_PROFILE} for copy)")
    ap.add_argument("--refresh-copy", action="store_true",
                    help="redo the profile copy (use if the copy is signed out)")
    ap.add_argument("--chrome-path", default=None, help="path to chrome.exe")
    ap.add_argument("--cdp-port", type=int, default=CDP_PORT)
    ap.add_argument("--force-close", action="store_true",
                    help="close a running Chrome without asking (real mode)")
    ap.add_argument("--close-browser", action="store_true",
                    help="also quit Chrome at the end (real mode leaves it open)")
    ap.add_argument("--wpm", type=int, default=72, help="typing speed (default 72)")
    ap.add_argument("--jitter", type=float, default=1.0,
                    help="timing multiplier; >1 is slower and more cautious")
    ap.add_argument("--timeout", type=int, default=240, help="seconds to wait per answer")
    ap.add_argument("--no-stealth", action="store_true", help="skip the patches (debugging)")
    ap.add_argument("--runtime-fix", default=None,
                    choices=("addBinding", "alwaysIsolated", "off"),
                    help="how rebrowser hides the Runtime.enable leak. alwaysIsolated survives pages that redirect mid-load, such as Cloudflare interstitials; addBinding hides more but can die on them.")
    ap.add_argument("--keep-open", action="store_true", help="pause before closing the tab")
    ap.add_argument("--selftest", action="store_true", help="print the fingerprint and exit")
    args = ap.parse_args()

    if args.list_profiles:
        user_data = find_user_data()
        if not user_data:
            sys.exit("Chrome's User Data folder not found.")
        print_profiles(user_data)
        return

    if args.login:
        interactive_login(args)
        return

    if args.selftest:
        with sync_playwright() as p:
            selftest(p, args)
        return

    queries = [q for q in args.query if q.strip()]
    if not queries:
        entered = input("Query for ChatGPT: ").strip()
        if not entered:
            sys.exit("No query given.")
        queries = [entered]
    if args.output and len(queries) > 1:
        sys.exit("--output takes a single query; use --outdir for several.")

    if not RUNTIME_PATCHED:
        log("Tip: pip install rebrowser-playwright closes the CDP Runtime.enable leak.")

    written = []
    with sync_playwright() as p:
        page, close = open_browser(p, args)
        try:
            human = Human(page, wpm=args.wpm, jitter=args.jitter)
            fit_window(page)
            visit(page, human)
            wait_for_login(page, human)

            for i, query in enumerate(queries, 1):
                if i > 1:
                    human.pause(2.0, 6.0)  # nobody fires queries back to back
                    visit(page, human)
                    page.wait_for_selector(COMPOSER, timeout=60_000)
                    human.pause(0.8, 2.0)

                before = page.locator(ASSISTANT_MSG).count()
                send_query(page, human, query)
                log("Waiting for the answer...")
                answer = wait_for_answer(page, before, timeout=args.timeout)

                human.read(answer)
                copy_to_clipboard(page, human, answer)

                if args.output:
                    out = pathlib.Path(args.output)
                else:
                    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
                    out = pathlib.Path(args.outdir) / (stamp + "-" + slugify(query) + ".txt")
                save(query, answer, out, page.url)
                written.append(out)

            if args.keep_open:
                input("Press Enter to finish...")
        finally:
            try:
                close()
            except Exception:
                pass

    if written:
        print("\nDone. Files written:")
        for path in written:
            print("  " + str(path.resolve()))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nInterrupted.")
