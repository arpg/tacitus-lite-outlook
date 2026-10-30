#!/usr/bin/env python3
"""Copy your Outlook mail into `mail/` as one Markdown file per message.

    py outlook_dump.py              # new mail since the last run (Inbox, Sent Items)
    py outlook_dump.py --full       # everything, from the start
    py outlook_dump.py --folders "Inbox" "Inbox/Projects" "Sent Items"
    py outlook_dump.py status       # what is on disk, no Outlook needed
    py outlook_dump.py selftest     # prove Outlook access on this machine

Exit codes: 0 ok, 1 a step failed, 2 Outlook is not reachable.
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
import tempfile
from pathlib import Path

from outlook_lite import __version__
from outlook_lite.dump import OutlookUnavailable, connect, dump, iter_messages, resolve_folder
from outlook_lite.render import MAIL_DIR, State

ROOT = Path(__file__).resolve().parent
DEFAULT_FOLDERS = ["Inbox", "Sent Items"]


def say(msg: str) -> None:
    print(msg, flush=True)


def cmd_dump(args) -> int:
    try:
        app, ns = connect()
    except OutlookUnavailable as e:
        say(f"Outlook not available: {e}")
        return 2
    say(f"outlook_dump {__version__}: {'full' if args.full else 'incremental'} run into {ROOT / MAIL_DIR}")
    counts = dump(ns, ROOT, args.folders, full=args.full, addresses=not args.no_addresses,
                  limit=args.limit, log=say)
    for name, n in counts.items():
        say(f"  {name}: {n} new")
    say(f"done at {dt.datetime.now():%Y-%m-%d %H:%M}")
    return 0


def cmd_status(args) -> int:
    state = State(ROOT)
    if not state.folders:
        say(f"no mail on disk yet ({ROOT / MAIL_DIR} is empty); run `py outlook_dump.py`")
        return 0
    for name, f in state.folders.items():
        say(f"{name}: {len(f['seen'])} messages, newest {f.get('watermark') or 'unknown'}")
    return 0


def cmd_selftest(args) -> int:
    """Each step prints PASS or FAIL; the first failure stops the run."""
    def step(label, fn):
        try:
            out = fn()
        except Exception as e:
            say(f"FAIL  {label}: {type(e).__name__}: {e}")
            raise
        say(f"PASS  {label}{': ' + str(out) if out is not None else ''}")
        return out

    try:
        app, ns = step("connect to classic Outlook", connect)
    except Exception:
        say("Outlook not available. Classic Outlook must be installed, open and signed in.")
        return 2
    try:
        step("Outlook version", lambda: app.Version)
        step("signed-in account", lambda: ns.Accounts.Item(1).SmtpAddress)
        for name in args.folders:
            folder = step(f"open folder {name!r}", lambda n=name: resolve_folder(ns, n))
            step(f"count items in {name!r}", lambda f=folder: f.Items.Count)
            msgs = step(f"read 3 newest messages in {name!r} (addresses on)",
                        lambda f=folder, n=name: list(iter_messages(f, n, limit=3)))
            for m in msgs:
                say(f"      {m.received:%Y-%m-%d %H:%M}  {m.sender_name!r}  <{m.sender_email}>  {m.subject[:50]!r}")
            say("      If Outlook showed a security prompt just now, add --no-addresses to the scheduled task.")
        with tempfile.TemporaryDirectory() as tmp:
            counts = step("dump 5 messages per folder to a temporary folder",
                          lambda: dump(ns, Path(tmp), args.folders, full=True, limit=5))
            files = sorted(p.relative_to(tmp) for p in Path(tmp).rglob("*.md"))
            for p in files:
                say(f"      {p}")
            if not files:
                say("FAIL  no files were written")
                return 1
        if args.draft:
            from outlook_lite.draft import create_draft
            me = ns.Accounts.Item(1).SmtpAddress
            eid = step("save a draft addressed to yourself",
                       lambda: create_draft(app, me, "outlook_dump selftest",
                                            "This draft was created by outlook_dump.py selftest. Delete it."))
            say("      Open Outlook > Drafts: the draft is there. Nothing was sent.")
    except Exception:
        return 1
    say("ALL PASS")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", nargs="?", default="dump", choices=["dump", "status", "selftest"])
    ap.add_argument("--folders", nargs="+", default=DEFAULT_FOLDERS,
                    help="default folders by name, or paths like Inbox/Projects")
    ap.add_argument("--full", action="store_true", help="read every message, not only new ones")
    ap.add_argument("--limit", type=int, default=None, help="stop after N messages per folder")
    ap.add_argument("--no-addresses", action="store_true",
                    help="do not read sender email addresses (avoids the Outlook security prompt)")
    ap.add_argument("--draft", action="store_true", help="selftest: also save a draft to yourself")
    args = ap.parse_args(argv)
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    return {"dump": cmd_dump, "status": cmd_status, "selftest": cmd_selftest}[args.command](args)


if __name__ == "__main__":
    sys.exit(main())
