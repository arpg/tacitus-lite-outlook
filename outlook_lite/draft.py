"""Drafts in Outlook's Drafts folder. There is no send function in this
package, on purpose: the person reads the draft in Outlook and clicks
Send there.
"""
from __future__ import annotations

import html
from pathlib import Path

from .render import read_headers

OL_MAIL_ITEM = 0  # OlItemType.olMailItem for CreateItem


def body_html(text: str) -> str:
    """Plain text as HTML paragraphs, for prepending to a reply's quoted thread."""
    paras = [p.strip() for p in text.replace("\r\n", "\n").split("\n\n")]
    out = []
    for p in paras:
        if not p:
            continue
        out.append("<p>" + html.escape(p).replace("\n", "<br>") + "</p>")
    return "\n".join(out) + "\n"


def create_draft(app, to: str, subject: str, body: str, cc: str = "") -> str:
    mail = app.CreateItem(OL_MAIL_ITEM)
    mail.To = to
    if cc:
        mail.CC = cc
    mail.Subject = subject
    mail.Body = body
    mail.Save()
    return str(mail.EntryID)


def reply_draft(ns, entry_id: str, body: str, reply_all: bool = True) -> str:
    """A reply to the message with `entry_id`, saved to Drafts with the
    quoted thread below the new text."""
    item = ns.GetItemFromID(entry_id)
    reply = item.ReplyAll() if reply_all else item.Reply()
    reply.HTMLBody = body_html(body) + (reply.HTMLBody or "")
    reply.Save()
    return str(reply.EntryID)


def entry_id_of(markdown_file: Path) -> str:
    """The EntryID header of a dumped message file."""
    headers = read_headers(Path(markdown_file))
    eid = headers.get("EntryID", "")
    if not eid:
        raise ValueError(f"{markdown_file} has no EntryID header")
    return eid
