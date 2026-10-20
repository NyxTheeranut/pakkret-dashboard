#!/usr/bin/env python3
"""
Fetches the latest version of the shared Excel files by asking YOUR OWN BROWSER to download them, then hands
them to intake_sources.py, which files them into Source Data. No download clicks, no copying files around.

Why the browser: it is already signed in to the company account. This script never sees a password, token or
cookie, and needs no IT approval. It only opens a download link and then watches your Downloads folder.

Where the links come from (links are read locally and never printed):
  * OneDrive "shortcut to a file" items (*.xlsx.url) found in the folders listed in Config/sources.json
    ("watch"; by default Downloads and every synced OneDrive folder). Adding such a shortcut in OneDrive is all
    that is needed to add a file.
  * an optional "links" list in Config/sources.json: [{"url": "...", "name": "optional file name"}]
Only files whose name matches one of the "rules" in Config/sources.json are fetched (same rules as the intake).

What it does:
  1. Builds a download link for each file (direct SharePoint path + ?download=1) and opens it as a BACKGROUND TAB
     in your existing browser window (Arc, Chrome, Edge, Brave, Safari on macOS), a couple of seconds apart: no new
     window, and the tab you are looking at does not change. Other browsers/systems: the link is just opened.
  2. Waits (up to "fetch.timeout" seconds, default 90) for the files to appear in Downloads.
  3. Closes the tabs this run opened if they are still open (only those; and only when every file arrived, so
     you can still see what went wrong otherwise). Some browsers, Arc included, close the tab themselves.
  4. Runs intake_sources.py, which copies new/changed files into Source Data/<RTR|7-11> (old copy archived,
     identical files left alone) and then deletes the temporary downloads this run made.

If anything does not arrive (not signed in, offline, downloads blocked) it says so and carries on with what is
already in Source Data. Usage:
    python3 sync/fetch_links.py        # fetch, then file them
    python3 sync/fetch_links.py --dry-run # only list which files it would fetch
Test options: --root, --watch (repeatable), --downloads <dir>, --opener <command> (replaces the browser).
"""
import argparse
import fnmatch
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from urllib.parse import unquote, urlparse, urlunparse

import intake_sources as intake

PARTIAL = (".crdownload", ".part", ".tmp", ".download")


def say(msg=""):
    print(msg, flush=True)


def read_shortcut(path):
    try:
        text = Path(path).read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    m = re.search(r"^URL=(.+)$", text, re.M)
    return m.group(1).strip() if m else None


def download_link(url):
    """Direct SharePoint/OneDrive file path + ?download=1 (what the browser needs to save instead of preview)."""
    u = urlparse(url)
    if not u.scheme.startswith("http"):
        return None
    if u.path.lower().endswith(".xlsx"):
        return urlunparse((u.scheme, u.netloc, u.path, "", "download=1", ""))
    # sharing-style link without a file path: best effort
    return url + ("&" if u.query else "?") + "download=1"


def expected_name(url, fallback):
    last = unquote(urlparse(url).path.rstrip("/").split("/")[-1])
    return intake.norm_name(last if last.lower().endswith(".xlsx") else fallback)


def collect(cfg, watch):
    """-> list of {"name", "url"} (deduplicated by file name) for files that match the rules."""
    rules = cfg["rules"]
    found = {}
    skipped = 0
    for w in watch:
        folder = Path(w["path"]).expanduser()
        if not folder.is_dir():
            continue
        for p in intake.scan(folder, int(w.get("depth", 1)), Path("/nonexistent")):
            if not p.name.lower().endswith(".xlsx.url"):
                continue
            url = read_shortcut(p)
            if not url:
                continue
            name = expected_name(url, p.name[:-4])
            if any(fnmatch.fnmatch(name.casefold(), r["pattern"].casefold()) for r in rules):
                found.setdefault(name.casefold(), {"name": name, "url": url})
            else:
                skipped += 1
    for item in cfg.get("links", []):
        url = item.get("url")
        if not url:
            continue
        name = intake.norm_name(item.get("name") or expected_name(url, "file.xlsx"))
        found.setdefault(name.casefold(), {"name": name, "url": url})
    return sorted(found.values(), key=lambda x: x["name"].casefold()), skipped


def open_in_browser(link, opener, browser=None):
    try:
        if browser and not opener and open_as_background_tab(browser, link):
            return True
        if opener:
            subprocess.run(opener.split() + [link], check=False)
        elif sys.platform == "darwin":
            subprocess.run(["open", "-g", link], check=False)
        elif sys.platform.startswith("win"):
            os.startfile(link)                                    # type: ignore[attr-defined]
        else:
            import webbrowser
            webbrowser.open(link)
        return True
    except Exception:
        return False



# ---- closing the tabs this run opened (macOS: Arc, Chrome/Edge/Brave, Safari) -------------------------------
APPS = {                                   # default-browser bundle id -> (app name, scripting family)
    "company.thebrowser.browser": ("Arc", "chromium"),
    "com.google.chrome": ("Google Chrome", "chromium"),
    "com.microsoft.edgemac": ("Microsoft Edge", "chromium"),
    "com.brave.browser": ("Brave Browser", "chromium"),
    "com.apple.safari": ("Safari", "safari"),
}
LIST_CHROMIUM = """on run argv
  set out to ""
  tell application "%APP%"
    set nw to count of windows
    repeat with i from 1 to nw
      set wid to (id of window i) as text
      set nt to count of tabs of window i
      repeat with j from 1 to nt
        set u to ""
        try
          set u to URL of tab j of window i
        end try
        set tid to ""
        try
          set tid to (id of tab j of window i) as text
        end try
        set out to out & wid & "|" & tid & "|" & u & linefeed
      end repeat
    end repeat
  end tell
  return out
end run"""
LIST_SAFARI = """on run argv
  set out to ""
  tell application "Safari"
    repeat with w in windows
      set wid to (id of w) as text
      set i to 0
      repeat with t in tabs of w
        set i to i + 1
        set u to ""
        try
          set u to URL of t
        end try
        set out to out & wid & "|" & i & "|" & u & linefeed
      end repeat
    end repeat
  end tell
  return out
end run"""
CLOSE_CHROMIUM = """on run argv
  tell application "%APP%"
    repeat with k from 2 to (count of argv)
      set AppleScript's text item delimiters to "|"
      set wid to text item 1 of (item k of argv)
      set tid to text item 2 of (item k of argv)
      try
        set nw to count of windows
        repeat with i from 1 to nw
          if ((id of window i) as text) is wid then
            set nt to count of tabs of window i
            repeat with j from 1 to nt
              if ((id of tab j of window i) as text) is tid then
                close tab j of window i
                exit repeat
              end if
            end repeat
          end if
        end repeat
      end try
    end repeat
  end tell
end run"""
CLOSE_SAFARI = """on run argv
  tell application "Safari"
    repeat with k from 2 to (count of argv)
      set AppleScript's text item delimiters to "|"
      set wid to text item 1 of (item k of argv)
      set idx to (text item 2 of (item k of argv)) as integer
      try
        repeat with w in windows
          if ((id of w) as text) is wid then close tab idx of w
        end repeat
      end try
    end repeat
  end tell
end run"""


def default_browser():
    """(app name, family) of the default browser on macOS, or None where tabs cannot be controlled."""
    if sys.platform != "darwin":
        return None
    try:
        out = subprocess.run(["plutil", "-convert", "json", "-o", "-", str(Path.home() / "Library/Preferences/com.apple.LaunchServices/com.apple.launchservices.secure.plist")],
                             capture_output=True, text=True, timeout=10).stdout
        for h in json.loads(out).get("LSHandlers", []):
            if h.get("LSHandlerURLScheme") == "https":
                return APPS.get((h.get("LSHandlerRoleAll") or "").lower())
    except Exception:
        pass
    return None


def run_osa(script, args, timeout=30):
    try:
        r = subprocess.run(["osascript", "-e", script, *args], capture_output=True, text=True, timeout=timeout)
        return (r.returncode == 0), r.stdout
    except Exception:
        return False, ""



# ---- opening a link as a BACKGROUND TAB in the existing window (no new window, your view does not change) -----
OPEN_ARC = """on run argv
  tell application "Arc"
    if (count of windows) is 0 then error "no window"
    tell window 1 to make new tab with properties {URL:(item 1 of argv)}
  end tell
end run"""
OPEN_CHROMIUM = """on run argv
  tell application "%APP%"
    if (count of windows) is 0 then error "no window"
    set w to front window
    set prevIdx to active tab index of w
    make new tab at end of tabs of w with properties {URL:(item 1 of argv)}
    set active tab index of w to prevIdx
  end tell
end run"""
OPEN_SAFARI = """on run argv
  tell application "Safari"
    if (count of windows) is 0 then error "no window"
    set w to front window
    set prev to current tab of w
    tell w to make new tab at end of tabs with properties {URL:(item 1 of argv)}
    set current tab of w to prev
  end tell
end run"""


def open_as_background_tab(browser, link):
    app, family = browser
    if app == "Arc":
        script = OPEN_ARC
    elif family == "safari":
        script = OPEN_SAFARI
    else:
        script = OPEN_CHROMIUM.replace("%APP%", app)
    ok, _ = run_osa(script, [link])
    return ok


def list_tabs(browser):
    """-> [(window id, tab id or index, url)] or None when the browser cannot be read (e.g. no permission yet)."""
    app, family = browser
    ok, out = run_osa(LIST_SAFARI if family == "safari" else LIST_CHROMIUM.replace("%APP%", app), [])
    if not ok:
        return None
    rows = []
    for line in out.splitlines():
        parts = line.split("|", 2)
        if len(parts) == 3:
            rows.append((parts[0], parts[1], parts[2]))
    return rows


def tabs_to_close(before, after, hosts):
    """Only tabs that did not exist before this run AND are blank or on one of our download hosts."""
    known = {(w, t) for w, t, _ in before}
    out = []
    for w, t, url in after:
        if (w, t) in known:
            continue
        low = url.strip().lower()
        blank = (low == "" or low.startswith(("about:", "chrome:", "arc:", "edge:", "favorites://", "topsites://")))
        host = (urlparse(url).hostname or "").lower()
        if blank or host in hosts:
            out.append((w, t, url))
    return out


def close_tabs(browser, tabs):
    app, family = browser
    specs = [f"{w}|{t}" for w, t, _ in tabs]
    if family == "safari":                                   # tab indexes shift: close from the end
        specs.sort(key=lambda sp: int(sp.split("|")[1]), reverse=True)
    ok, _ = run_osa(CLOSE_SAFARI if family == "safari" else CLOSE_CHROMIUM.replace("%APP%", app), ["-", *specs])
    return ok


def downloads_snapshot(folder):
    return {p: p.stat().st_mtime for p in folder.glob("*") if p.is_file()}


def wait_for(folder, before, expected, timeout):
    """Poll until every expected file has arrived and stopped growing; return (arrived paths, missing names)."""
    want = {n.casefold(): n for n in expected}
    sizes, arrived = {}, {}
    t0 = time.time()
    while time.time() - t0 < timeout:
        time.sleep(1.5)
        for p in folder.glob("*"):
            if not p.is_file() or p.name.lower().endswith(PARTIAL):
                continue
            key = intake.norm_name(p.name).casefold()
            if key not in want:
                continue
            try:
                st = p.stat()
            except OSError:
                continue
            if p in before and st.st_mtime <= before[p] + 1:
                continue                                           # an old file that was already there
            if sizes.get(p) == st.st_size and st.st_size > 0:
                arrived[p] = key
            sizes[p] = st.st_size
        if {k for k in arrived.values()} >= set(want):
            break
    got = {k for k in arrived.values()}
    return sorted(arrived), sorted(want[k] for k in want if k not in got)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--root", default=str(intake.DEFAULT_ROOT))
    ap.add_argument("--watch", action="append", default=[])
    ap.add_argument("--downloads", default=str(Path.home() / "Downloads"))
    ap.add_argument("--opener", default=None, help="test hook: command used instead of the browser")
    args = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    root = Path(args.root).resolve()
    cfg = intake.load_config(root, args.dry_run)
    fcfg = {"enabled": True, "timeout": 90, "stagger": 2.5, **cfg.get("fetch", {})}
    watch = [{"path": w, "depth": 2} for w in args.watch] if args.watch else cfg["watch"]
    dl_dir = Path(args.downloads)
    passthrough = (["--dry-run"] if args.dry_run else []) + ["--quiet-links", "--root", str(root)] + [x for w in args.watch for x in ("--watch", w)]

    say("▶ Fetching the latest shared files…" + ("  (dry run: nothing will be opened)" if args.dry_run else ""))
    if not fcfg["enabled"]:
        say("  (fetching is switched off in Config/sources.json)")
        intake.main(passthrough)
        return
    items, skipped = collect(cfg, watch)
    if skipped:
        say(f"  ({skipped} shortcut(s) ignored: the file name matches none of the rules)")
    if not items:
        say("  no link shortcuts found, checking the folders for files instead")
        intake.main(passthrough)
        return
    for it in items:
        say(f"  · {it['name']}")

    manifest = None
    if not args.dry_run:
        before = downloads_snapshot(dl_dir) if dl_dir.is_dir() else {}
        browser = None if args.opener else default_browser()
        tabs_before = list_tabs(browser) if browser else None
        if browser and tabs_before is None:
            say(f"  (cannot read {browser[0]}'s tabs, so they will be left open. To allow closing them: System Settings > Privacy & Security > Automation, allow Terminal to control {browser[0]})")
            browser = None
        opened = 0
        for i, it in enumerate(items):
            link = download_link(it["url"])
            if link and open_in_browser(link, args.opener, browser):
                opened += 1
            if i < len(items) - 1:
                time.sleep(float(fcfg["stagger"]))
        say(f"  asked the browser to download {opened} file(s); waiting for them…")
        arrived, missing = wait_for(dl_dir, before, [it["name"] for it in items], float(fcfg["timeout"]))
        say(f"  received {len({intake.norm_name(p.name).casefold() for p in arrived})} of {len(items)} file(s)")
        if missing:
            say("  ! not received: " + ", ".join(missing))
            say("    (signed out of the company account in your browser? offline? downloads blocked?)")
            if browser:
                say("    the tabs were left open so you can see why")
        elif browser and tabs_before is not None:
            time.sleep(1.5)
            tabs_after = list_tabs(browser)
            if tabs_after is not None:
                hosts = {(urlparse(download_link(it["url"]) or "").hostname or "").lower() for it in items}
                mine = tabs_to_close(tabs_before, tabs_after, hosts)
                if mine and close_tabs(browser, mine):
                    say(f"  closed the {len(mine)} browser tab(s) this run opened")
        if arrived:
            fd, manifest = tempfile.mkstemp(prefix="fetched_", suffix=".json")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump([str(p) for p in arrived], f)
        say()
    intake.main(passthrough + (["--cleanup-manifest", manifest] if manifest else []))


if __name__ == "__main__":
    main()
