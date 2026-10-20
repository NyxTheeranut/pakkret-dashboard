#!/usr/bin/env python3
"""
Picks up newly downloaded (or OneDrive-synced) Excel files and files them into Source Data, so the
update only needs one double-click instead of download -> move -> rerun.

What it does, in order:
  1. Looks in the folders listed in Config/sources.json ("watch"): by default Downloads plus any
     OneDrive-synced folder it can find. Only a few levels deep, only *.xlsx.
  2. Keeps files whose name matches one of the "rules" (e.g. "AE Monitoring*" -> 7-11). Browser copies
     such as "name (1).xlsx" count as the same file as "name.xlsx"; the newest one wins.
  3. Compares with the copy already in Source Data/<RTR|7-11>/:
       new        no copy yet                      -> copied in
       updated    the found file is newer           -> old copy archived, new one copied in
       unchanged  same file                         -> nothing to do
       older      the found file is older           -> kept as is
  4. Never moves or deletes the original. Replaced copies go to Source Data/_archive/ (last few kept).

Skipped on purpose: lock files (~$...), half-downloaded files (.crdownload/.part/.tmp, or modified in the
last few seconds), empty files and anything that isn't a valid xlsx (zip) file.

Reads no secrets. Usage:
    python3 sync/intake_sources.py         # do it
    python3 sync/intake_sources.py --dry-run # show what it would do, change nothing
Test/override options: --root <dir> (a Dashboard folder), --watch <dir> (repeatable, replaces the config).
"""
import argparse
import fnmatch
import hashlib
import json
import os
import re
import shutil
import sys
import time
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_ROOT = HERE.parent.parent  # this script lives in Territory Explorer/sync/

DEFAULT_RULES = [
    {"pattern": "Commit Target GA*", "dest": "RTR"},
    {"pattern": "RTR_Commit Routing*", "dest": "RTR"},
    {"pattern": "7-Eleven Store Database*", "dest": "7-11"},
    {"pattern": "AE Monitoring*", "dest": "7-11"},
    {"pattern": "Data Synergy*", "dest": "7-11"},
    {"pattern": "LE by store*", "dest": "7-11"},
    {"pattern": "Target 7-eleven*", "dest": "7-11"},
]
PARTIAL_SUFFIXES = (".crdownload", ".part", ".tmp", ".download")
MAX_FILES_SCANNED = 30000


def say(msg=""):
    print(msg, flush=True)


def norm_name(name):
    """NFC-normalised name with a browser's " (1)" copy suffix removed."""
    n = unicodedata.normalize("NFC", name)
    return re.sub(r"\s\(\d+\)(?=\.xlsx$)", "", n, flags=re.I)


def default_watch():
    home = Path.home()
    watch = [{"path": str(home / "Downloads"), "depth": 1}]
    cloud = home / "Library" / "CloudStorage"          # macOS OneDrive client
    if cloud.is_dir():
        for p in sorted(cloud.glob("OneDrive-*")):
            watch.append({"path": str(p), "depth": 3})
    for var in ("OneDriveCommercial", "OneDrive"):      # Windows OneDrive client
        v = os.environ.get(var)
        if v and all(w["path"] != v for w in watch):
            watch.append({"path": v, "depth": 3})
    return watch


def load_config(root, dry):
    cfg_path = root / "Config" / "sources.json"
    if cfg_path.exists():
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    else:
        cfg = {"watch": default_watch(), "rules": DEFAULT_RULES, "keep_archive": 2, "min_age_seconds": 5}
        if not dry:
            cfg_path.parent.mkdir(parents=True, exist_ok=True)
            cfg_path.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
            say(f"  created {cfg_path} (edit it to add folders or file name rules)")
    cfg.setdefault("rules", DEFAULT_RULES)
    cfg.setdefault("keep_archive", 2)
    cfg.setdefault("min_age_seconds", 5)
    return cfg


def scan(folder, depth, skip_under):
    """Yield *.xlsx files under folder, at most `depth` levels down, bounded in total."""
    seen = 0
    stack = [(Path(folder), 0)]
    while stack:
        d, level = stack.pop()
        try:
            entries = list(os.scandir(d))
        except OSError:
            continue
        for e in entries:
            seen += 1
            if seen > MAX_FILES_SCANNED:
                return
            if e.name.startswith("."):
                continue
            p = Path(e.path)
            try:
                if e.is_dir(follow_symlinks=False):
                    if level < depth and skip_under not in p.parents and p != skip_under:
                        stack.append((p, level + 1))
                elif e.name.lower().endswith((".xlsx", ".xlsx.url")) and not e.name.startswith("~$"):
                    yield p
            except OSError:
                continue


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def human(nbytes):
    return f"{nbytes / 1e6:.1f} MB" if nbytes >= 1e6 else f"{nbytes / 1e3:.0f} KB"


def ago(seconds):
    if seconds < 90:
        return "just now"
    if seconds < 5400:
        return f"{seconds / 60:.0f} min ago"
    if seconds < 172800:
        return f"{seconds / 3600:.0f} h ago"
    return f"{seconds / 86400:.0f} days ago"


def prune_archive(archive_dir, base, keep):
    olds = sorted(archive_dir.glob(base + ".bak-*"), key=lambda p: p.name, reverse=True)
    for p in olds[keep:]:
        try:
            p.unlink()
        except OSError:
            pass


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--root", default=str(DEFAULT_ROOT))
    ap.add_argument("--watch", action="append", default=[])
    ap.add_argument("--quiet-links", action="store_true", help="do not explain .url shortcuts (fetch_links.py handles them)")
    ap.add_argument("--cleanup-manifest", default=None,
                    help="JSON list of downloads made by fetch_links.py; they are deleted after a successful run")
    args = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    root = Path(args.root).resolve()
    source_dir = root / "Territory Explorer" / "Source Data"
    archive_dir = source_dir / "_archive"
    cfg = load_config(root, args.dry_run)
    watch = [{"path": w, "depth": 2} for w in args.watch] if args.watch else cfg["watch"]
    rules = cfg["rules"]
    now = time.time()

    say("▶ Checking for new files…" + ("  (dry run: nothing will be changed)" if args.dry_run else ""))
    for w in watch:
        say(f"  looking in: {w['path']}")

    # 1) collect candidates, newest per target name
    best = {}                      # (dest, name) -> (mtime, size, path)
    skipped = []
    link_only = []
    for w in watch:
        folder = Path(w["path"]).expanduser()
        if not folder.is_dir():
            continue
        for p in scan(folder, int(w.get("depth", 1)), source_dir):
            if p.name.lower().endswith(".url"):
                # a OneDrive "shortcut to a file" is only a web link (a few hundred bytes), not the Excel file
                if any(fnmatch.fnmatch(norm_name(p.name[:-4]).casefold(), r["pattern"].casefold()) for r in rules):
                    link_only.append(p.name[:-4])
                continue
            name = norm_name(p.name)
            rule = next((r for r in rules if fnmatch.fnmatch(name.casefold(), r["pattern"].casefold())), None)
            if not rule:
                continue
            try:
                st = p.stat()
            except OSError:
                continue
            if p.name.lower().endswith(PARTIAL_SUFFIXES) or st.st_size == 0:
                skipped.append((p.name, "empty or still downloading"))
                continue
            if now - st.st_mtime < cfg["min_age_seconds"]:
                skipped.append((p.name, "still being written"))
                continue
            if not zipfile.is_zipfile(p):
                hint = "not a valid xlsx file"
                if "OneDrive" in str(p):
                    hint += " (online-only OneDrive file? right-click it > Always keep on this device)"
                skipped.append((p.name, hint))
                continue
            key = (rule["dest"], name)
            cur = best.get(key)
            if cur is None or (st.st_mtime, st.st_size) > (cur[0], cur[1]):
                best[key] = (st.st_mtime, st.st_size, p)

    # 2) decide + act
    counts = {"new": 0, "updated": 0, "unchanged": 0, "older": 0}
    failed_names = set()
    for (dest, name), (mtime, size, src) in sorted(best.items()):
        # an existing copy may carry a browser suffix, e.g. "name (1).xlsx": keep using that file instead of
        # adding a second copy next to it (the update would then read both)
        variants = [q for q in (source_dir / dest).glob("*.xlsx") if norm_name(q.name) == name]
        target = max(variants, key=lambda q: q.stat().st_mtime) if variants else source_dir / dest / name
        where = src.parent.name or str(src.parent)
        if not target.exists():
            action = "new"
        else:
            tst = target.stat()
            if mtime <= tst.st_mtime + 2:
                action = "unchanged" if abs(size - tst.st_size) == 0 else "older"
            elif size == tst.st_size and sha256(src) == sha256(target):
                action = "unchanged"
            else:
                action = "updated"
        counts[action] += 1
        icon = {"new": "+", "updated": "✓", "unchanged": "=", "older": "·"}[action]
        extra = {"new": f"{human(size)}, from {where}", "updated": f"{human(size)}, newer ({ago(now - mtime)}), from {where}",
                 "unchanged": "same as the copy already in Source Data", "older": "the copy in Source Data is newer, kept"}[action]
        shown = target.name if target.exists() else name
        say(f"  {icon} {action:<9} {dest:<5} {shown}  ({extra})")
        if args.dry_run or action not in ("new", "updated"):
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_name(target.name + ".part")
        shutil.copy2(src, tmp)
        if tmp.stat().st_size != size or not zipfile.is_zipfile(tmp):
            tmp.unlink(missing_ok=True)
            say(f"    ✗ copy of {name} did not verify, left the old file in place")
            counts[action] -= 1
            failed_names.add(name.casefold())
            continue
        if target.exists():
            archive_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.fromtimestamp(target.stat().st_mtime).strftime("%Y%m%d%H%M%S")
            shutil.move(str(target), str(archive_dir / f"{target.name}.bak-{stamp}"))
            prune_archive(archive_dir, target.name, int(cfg["keep_archive"]))
        os.replace(tmp, target)

    for n, why in skipped:
        say(f"  ! skipped   {n}  ({why})")
    if link_only and not args.quiet_links:
        say("  ! link only: " + ", ".join(sorted(set(link_only))))
        say("    these OneDrive shortcuts are web links, not the Excel files, so they cannot be read.")
        say("    Ask the owner to share the FOLDER that holds the real files, then Add shortcut to My files on that folder.")
    if not best:
        say("  (no matching files found in the watched folders)")
    if args.cleanup_manifest and not args.dry_run:
        removed = 0
        try:
            made = json.loads(Path(args.cleanup_manifest).read_text(encoding="utf-8"))
        except Exception:
            made = []
        for q in made:
            qp = Path(q)
            if norm_name(qp.name).casefold() in failed_names:
                continue                                   # its copy did not verify: keep the download
            try:
                if qp.exists():
                    qp.unlink()
                    removed += 1
            except OSError:
                pass
        try:
            Path(args.cleanup_manifest).unlink()
        except OSError:
            pass
        if removed:
            say(f"  tidied {removed} temporary download(s) from Downloads")
    say(f"  done: {counts['new']} new, {counts['updated']} updated, {counts['unchanged']} unchanged"
        + (f", {counts['older']} older" if counts["older"] else "") + (f", {len(skipped)} skipped" if skipped else ""))
    say()
    return counts


if __name__ == "__main__":
    main()
