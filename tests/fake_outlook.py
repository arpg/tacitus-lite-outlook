"""A stand-in for the parts of the Outlook COM object model the code uses.

Attribute names and shapes follow the real model (1-based `Item`,
`Class == 43` for mail, `Items.Sort`, `ReplyAll().HTMLBody` carrying the
quoted thread, `Save()` landing in Drafts), so a test that passes here
exercises the same calls the Windows run makes. What it cannot prove is
that pywin32 marshals them; `outlook_dump.py selftest` does that.
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

OL_MAIL = 43
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"


class ComError(Exception):
    pass


class FakeAttachment:
    def __init__(self, name):
        self.FileName = name


class FakeCollection:
    """1-based `Item(i)`, `Count`, iteration: the COM collection shape."""

    def __init__(self, items=None):
        self._items = list(items or [])

    @property
    def Count(self):
        return len(self._items)

    def Item(self, key):
        if isinstance(key, int):
            return self._items[key - 1]
        for it in self._items:
            if getattr(it, "Name", None) == key:
                return it
        raise ComError(f"no item named {key!r}")

    def __iter__(self):
        return iter(list(self._items))

    def append(self, item):
        self._items.append(item)


class FakeItems(FakeCollection):
    def Sort(self, prop, descending=False):
        assert prop == "[ReceivedTime]"
        self._items.sort(key=lambda m: m.ReceivedTime or dt.datetime.min, reverse=bool(descending))


class FakePropertyAccessor:
    def __init__(self, mail, guard=False):
        self._mail, self._guard = mail, guard

    def GetProperty(self, tag):
        if self._guard:
            raise ComError("access denied by the Object Model Guard")
        if tag == PR_SENDER_SMTP:
            return self._mail._smtp
        raise ComError(f"unknown property {tag}")


class FakeMail:
    Class = OL_MAIL

    def __init__(self, entry_id, subject, sender_name, smtp, to, received, body,
                 cc="", attachments=(), conversation_id="", guard=False, html_body=None):
        self.EntryID = entry_id
        self.Subject = subject
        self.SenderName = sender_name
        self._smtp = smtp
        # Exchange gives the X500 address here; only the SMTP property is readable.
        self.SenderEmailAddress = f"/O=EXCHANGE/OU=GROUP/CN=RECIPIENTS/CN={entry_id}"
        self.To = to
        self.CC = cc
        self.ReceivedTime = received
        self.Body = body
        self.HTMLBody = html_body if html_body is not None else "<p>" + body.replace("\n", "<br>") + "</p>"
        self.ConversationID = conversation_id
        self.Attachments = FakeCollection(FakeAttachment(a) for a in attachments)
        self.PropertyAccessor = FakePropertyAccessor(self, guard)
        self.saved = False
        self._ns = None

    def _reply(self, all_):
        to = self._smtp if not all_ else "; ".join(x for x in [self._smtp, self.To, self.CC] if x)
        r = FakeMail(f"reply-{self.EntryID}", "RE: " + self.Subject, "Me", "me@example.edu", to,
                     dt.datetime.now(), "", html_body="<hr>" + self.HTMLBody)
        r._ns = self._ns
        return r

    def Reply(self):
        return self._reply(False)

    def ReplyAll(self):
        return self._reply(True)

    def Save(self):
        self.saved = True
        if self._ns is not None:
            self._ns.drafts.Items.append(self)


class FakeNonMail:
    """A meeting request or a report item: `Class` is not 43."""
    Class = 53

    def __init__(self, received):
        self.ReceivedTime = received
        self.EntryID = "nonmail"


class FakeFolder:
    def __init__(self, name, items=(), parent=None):
        self.Name = name
        self.Items = FakeItems(items)
        self.Folders = FakeCollection()
        self.Parent = parent

    def add_subfolder(self, folder):
        folder.Parent = self
        self.Folders.append(folder)
        return folder


class FakeAccount:
    def __init__(self, smtp):
        self.SmtpAddress = smtp


class FakeNamespace:
    def __init__(self, inbox, sent, drafts, smtp="me@example.edu"):
        self.store = FakeFolder("Mailbox")
        for f in (inbox, sent, drafts):
            self.store.add_subfolder(f)
        self.inbox, self.sent, self.drafts = inbox, sent, drafts
        self.Accounts = FakeCollection([FakeAccount(smtp)])
        self.Folders = FakeCollection([self.store])
        self._index = {}
        for f in (inbox, sent, drafts):
            self._register(f)

    def _register(self, folder):
        for it in folder.Items:
            self._index[it.EntryID] = it
            it._ns = self
        for sub in folder.Folders:
            self._register(sub)

    def GetDefaultFolder(self, const):
        return {6: self.inbox, 5: self.sent, 16: self.drafts}[const]

    def GetItemFromID(self, entry_id):
        try:
            return self._index[entry_id]
        except KeyError:
            raise ComError("The item could not be found") from None


class FakeApp:
    Version = "16.0.fake"

    def __init__(self, ns):
        self._ns = ns

    def GetNamespace(self, kind):
        assert kind == "MAPI"
        return self._ns

    def CreateItem(self, kind):
        assert kind == 0
        m = FakeMail("new-" + str(self._ns.drafts.Items.Count + 1), "", "Me", "me@example.edu", "",
                     dt.datetime.now(), "")
        m._ns = self._ns
        return m


def mailbox_from_json(path: Path, guard: bool = False) -> tuple[FakeApp, FakeNamespace]:
    """The fixture mailbox: `tests/fixtures/mailbox.json`."""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    folders = {"Inbox": [], "Sent Items": []}
    for i, m in enumerate(data["messages"]):
        folders[m.get("folder", "Inbox")].append(FakeMail(
            entry_id=m.get("entry_id", f"EID{i:04d}"),
            subject=m["subject"], sender_name=m["from_name"], smtp=m["from"],
            to=m.get("to", "me@example.edu"), cc=m.get("cc", ""),
            received=dt.datetime.fromisoformat(m["date"]), body=m["body"],
            attachments=m.get("attachments", ()), conversation_id=m.get("conversation_id", ""),
            guard=guard))
    ns = FakeNamespace(FakeFolder("Inbox", folders["Inbox"]), FakeFolder("Sent Items", folders["Sent Items"]),
                       FakeFolder("Drafts"))
    return FakeApp(ns), ns
