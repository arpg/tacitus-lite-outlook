"""The COM side of the dump: walk Outlook folders, read mail items.

Every function takes the Outlook objects as arguments. `connect()` is
the only place that touches `win32com`, and it is called only by the
command line, so the tests pass `tests/fake_outlook.py` objects instead.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Callable, Iterable, Iterator

from .render import Message, State, message_path, render_markdown

OL_MAIL_ITEM = 43           # MailItem.Class
OL_FOLDER_INBOX = 6         # OlDefaultFolders
OL_FOLDER_SENT = 5
OL_FOLDER_DRAFTS = 16
DEFAULT_FOLDERS = {"Inbox": OL_FOLDER_INBOX, "Sent Items": OL_FOLDER_SENT}
# PR_SENDER_SMTP_ADDRESS: the SMTP form of an Exchange sender.
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
# How far below the watermark a run re-reads, so a message that arrived
# late with an older ReceivedTime is still picked up.
OVERLAP = dt.timedelta(days=2)


class OutlookUnavailable(RuntimeError):
    """Outlook could not be reached: not installed, not classic, or not signed in."""


def connect():
    """The running (or newly started) classic Outlook, and its MAPI namespace."""
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as e:
        raise OutlookUnavailable(
            "pywin32 is not installed in this Python: run `py -m pip install pywin32`") from e
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
    except Exception as e:  # pywintypes.com_error
        raise OutlookUnavailable(
            "could not start Outlook.Application: classic Outlook must be installed and "
            "signed in (the New Outlook has no COM object model; switch the toggle off)") from e
    return app, ns


def resolve_folder(ns, name: str):
    """A default folder by its English name, or a path like `Inbox/Projects`.

    A path is walked from the store that holds the Inbox, so it is the
    same tree the folder pane shows.
    """
    if name in DEFAULT_FOLDERS:
        return ns.GetDefaultFolder(DEFAULT_FOLDERS[name])
    parts = [p for p in name.replace("\\", "/").split("/") if p]
    root = ns.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    node = root
    for part in parts:
        if part in DEFAULT_FOLDERS:
            node = ns.GetDefaultFolder(DEFAULT_FOLDERS[part])
            continue
        node = node.Folders.Item(part)
    return node


def _get(item, attr, default=""):
    try:
        v = getattr(item, attr)
    except Exception:
        return default
    return default if v is None else v


def _to_datetime(value) -> dt.datetime | None:
    """A pywintypes datetime (or anything with the fields) as a naive local datetime."""
    if value is None:
        return None
    try:
        if value.year < 1971:   # Outlook's "no date" is 4501-01-01, but guard both ends
            return None
        if value.year > 4000:
            return None
        return dt.datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    except Exception:
        return None


def sender_email(item, addresses: bool) -> str:
    if not addresses:
        return ""
    try:
        v = item.PropertyAccessor.GetProperty(PR_SENDER_SMTP)
        if v:
            return str(v)
    except Exception:
        pass
    v = _get(item, "SenderEmailAddress")
    return str(v) if v and "@" in str(v) else ""


def read_item(item, folder: str, addresses: bool = True) -> Message:
    attachments: list[str] = []
    try:
        atts = item.Attachments
        for i in range(1, int(atts.Count) + 1):
            attachments.append(str(atts.Item(i).FileName))
    except Exception:
        pass
    return Message(
        entry_id=str(_get(item, "EntryID")),
        folder=folder,
        subject=str(_get(item, "Subject")),
        sender_name=str(_get(item, "SenderName")),
        sender_email=sender_email(item, addresses),
        to=str(_get(item, "To")),
        cc=str(_get(item, "CC")),
        received=_to_datetime(_get(item, "ReceivedTime", None)),
        body=str(_get(item, "Body")),
        conversation_id=str(_get(item, "ConversationID")),
        attachments=attachments,
    )


def iter_messages(folder, name: str, after: dt.datetime | None = None,
                  addresses: bool = True, limit: int | None = None) -> Iterator[Message]:
    """Mail items of `folder`, newest first, stopping below `after`."""
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    n = 0
    for item in items:
        if limit is not None and n >= limit:
            return
        try:
            if int(item.Class) != OL_MAIL_ITEM:
                continue
        except Exception:
            continue
        msg = read_item(item, name, addresses)
        if after and msg.received and msg.received < after:
            return
        n += 1
        yield msg


def dump_folder(ns, root: Path, state: State, name: str, *, full: bool = False,
                addresses: bool = True, limit: int | None = None,
                log: Callable[[str], None] = lambda s: None) -> int:
    """Write every new message of one folder; return how many were written."""
    folder = resolve_folder(ns, name)
    after = None if full else state.watermark(name)
    if after:
        after = after - OVERLAP
    written = 0
    for msg in iter_messages(folder, name, after=after, addresses=addresses, limit=limit):
        if not msg.entry_id or state.has(name, msg.entry_id):
            continue
        path = message_path(root, msg)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(render_markdown(msg))
        state.record(msg, path)
        written += 1
        if written % 200 == 0:
            state.save()
            log(f"  {name}: {written} written so far")
    state.save()
    return written


def dump(ns, root: Path, folders: Iterable[str], **kw) -> dict[str, int]:
    root = Path(root)
    state = State(root)
    return {name: dump_folder(ns, root, state, name, **kw) for name in folders}
