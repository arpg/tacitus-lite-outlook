#!/usr/bin/env python3
"""Put a draft in Outlook's Drafts folder. Never sends.

    py make_draft.py --reply-to "mail/Inbox/2026/<file>.md" --body-file drafts/reply.txt
    py make_draft.py --reply-to <file> --body-file <txt> --sender-only
    py make_draft.py --to someone@example.edu --subject "Visitor parking" --body-file drafts/new.txt

The reply keeps the quoted thread below the new text, the way Outlook's
own Reply All does. Exit codes: 0 ok, 1 bad input, 2 Outlook not reachable.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from outlook_lite.draft import create_draft, entry_id_of, reply_draft
from outlook_lite.dump import OutlookUnavailable, connect


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--body-file", required=True, help="plain-text file with the draft body")
    ap.add_argument("--reply-to", help="a dumped message file (mail/.../*.md) to reply to")
    ap.add_argument("--sender-only", action="store_true", help="Reply instead of Reply All")
    ap.add_argument("--to", help="new message: recipients, semicolon separated")
    ap.add_argument("--cc", default="")
    ap.add_argument("--subject", default="")
    args = ap.parse_args(argv)

    body = Path(args.body_file).read_text(encoding="utf-8").strip()
    if not body:
        print(f"{args.body_file} is empty; a draft needs a body")
        return 1
    if not args.reply_to and not args.to:
        print("give --reply-to <dumped message file> or --to <address>")
        return 1
    try:
        app, ns = connect()
    except OutlookUnavailable as e:
        print(f"Outlook not available: {e}")
        return 2
    if args.reply_to:
        eid = entry_id_of(Path(args.reply_to))
        reply_draft(ns, eid, body, reply_all=not args.sender_only)
        print(f"Draft saved in Outlook > Drafts: {'reply' if args.sender_only else 'reply all'} to {Path(args.reply_to).name}")
    else:
        create_draft(app, args.to, args.subject, body, cc=args.cc)
        print(f"Draft saved in Outlook > Drafts: to {args.to} re: {args.subject!r}")
    print("Open Outlook, read it, then click Send yourself.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
