#!/usr/bin/env python3
"""One command for everything; the installed program is this file frozen.

    tacitus-lite dump [--full] [--folders ...] [--no-addresses]
    tacitus-lite status
    tacitus-lite selftest [--draft]
    tacitus-lite draft --reply-to <file> --body-file <txt> [--sender-only]
    tacitus-lite draft --to <addr> --subject <s> --body-file <txt>
    tacitus-lite home            # print the mail folder

`--pause` on any command waits for Enter before the window closes (the
Start menu shortcuts use it). Run as a windowless program (the
scheduled task), output goes to `mail/dump.log` in the home folder.
From a source checkout, `py tacitus_lite.py ...` is the same command.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

from outlook_lite import __version__
from outlook_lite.paths import home

USAGE = __doc__
COMMANDS = ("dump", "status", "selftest", "draft", "home", "version")


def _log_file() -> Path:
    p = home() / "mail" / "dump.log"
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    pause = "--pause" in argv
    argv = [a for a in argv if a != "--pause"]
    if sys.stdout is None or sys.stderr is None:
        # A windowless build has no console; keep a log instead.
        fh = open(_log_file(), "a", encoding="utf-8")
        sys.stdout = sys.stderr = fh
        print(f"--- {dt.datetime.now():%Y-%m-%d %H:%M:%S} tacitus-lite {' '.join(argv)}")
    elif sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    try:
        return _dispatch(argv)
    finally:
        sys.stdout.flush()
        if pause:
            try:
                input("\nPress Enter to close this window.")
            except EOFError:
                pass


def _dispatch(argv: list[str]) -> int:
    if not argv or argv[0] in ("-h", "--help", "help"):
        print(USAGE)
        return 0
    cmd, rest = argv[0], argv[1:]
    if cmd == "version":
        print(f"tacitus-lite {__version__}")
        return 0
    if cmd == "home":
        print(home())
        return 0
    if cmd == "draft":
        import make_draft
        return make_draft.main(rest)
    if cmd in ("dump", "status", "selftest"):
        import outlook_dump
        return outlook_dump.main([cmd, *rest])
    print(f"unknown command {cmd!r}; one of {', '.join(COMMANDS)}")
    print(USAGE)
    return 1


if __name__ == "__main__":
    sys.exit(main())
