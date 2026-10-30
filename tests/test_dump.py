import datetime as dt
import json
from pathlib import Path

import pytest
from fake_outlook import FakeFolder, FakeMail, FakeNonMail

from outlook_lite.dump import OutlookUnavailable, dump, iter_messages, resolve_folder
from outlook_lite.render import State, read_headers

FIXTURE = Path(__file__).parent / "fixtures" / "mailbox.json"
N_INBOX = 10
N_SENT = 2


def _files(root):
    return sorted(p for p in Path(root).rglob("*.md"))


def test_full_dump_writes_one_file_per_message(tmp_path, mailbox):
    app, ns = mailbox
    counts = dump(ns, tmp_path, ["Inbox", "Sent Items"], full=True)
    assert counts == {"Inbox": N_INBOX, "Sent Items": N_SENT}
    files = _files(tmp_path)
    assert len(files) == N_INBOX + N_SENT
    assert all(f.parts[-3] in ("Inbox", "Sent-Items") for f in files)
    # One file to spot-check end to end.
    portal = [f for f in files if "visitor-parking-validations" in f.name]
    assert len(portal) == 1
    h = read_headers(portal[0])
    assert h["From"] == "Finance Office <finance@example.edu>"
    assert h["Attachments"] == "parking-portal-guide.pdf"
    text = portal[0].read_text(encoding="utf-8")
    assert "Parking Services portal" in text
    assert "\r" not in text


def test_second_run_is_a_no_op_and_new_mail_is_picked_up(tmp_path, mailbox):
    app, ns = mailbox
    dump(ns, tmp_path, ["Inbox"], full=True)
    assert dump(ns, tmp_path, ["Inbox"]) == {"Inbox": 0}
    new = FakeMail("EID-NEW", "Brand new", "Someone", "someone@example.edu", "me@example.edu",
                   dt.datetime(2026, 10, 1, 8), "fresh")
    ns.inbox.Items.append(new)
    assert dump(ns, tmp_path, ["Inbox"]) == {"Inbox": 1}
    assert State(tmp_path).has("Inbox", "EID-NEW")


def test_late_arrival_inside_overlap_is_picked_up(tmp_path, mailbox):
    """A message whose ReceivedTime is a day older than the watermark still lands."""
    app, ns = mailbox
    dump(ns, tmp_path, ["Inbox"], full=True)
    wm = State(tmp_path).watermark("Inbox")
    late = FakeMail("EID-LATE", "Late", "Someone", "someone@example.edu", "me@example.edu",
                    wm - dt.timedelta(days=1), "late")
    ns.inbox.Items.append(late)
    assert dump(ns, tmp_path, ["Inbox"]) == {"Inbox": 1}
    old = FakeMail("EID-OLD", "Old", "Someone", "someone@example.edu", "me@example.edu",
                   wm - dt.timedelta(days=30), "old")
    ns.inbox.Items.append(old)
    assert dump(ns, tmp_path, ["Inbox"]) == {"Inbox": 0}, "below the overlap window: skipped"


def test_non_mail_items_are_skipped(tmp_path, mailbox):
    app, ns = mailbox
    ns.inbox.Items.append(FakeNonMail(dt.datetime(2026, 12, 1)))
    counts = dump(ns, tmp_path, ["Inbox"], full=True)
    assert counts == {"Inbox": N_INBOX}


def test_limit_takes_the_newest(tmp_path, mailbox):
    app, ns = mailbox
    msgs = list(iter_messages(ns.inbox, "Inbox", limit=3))
    assert [m.entry_id for m in msgs] == ["EID-LUNCH", "EID-BUDGET-ODD-SUBJECT", "EID-NEWSLETTER"]


def test_guarded_addresses_fall_back_without_crashing(tmp_path, guarded_mailbox):
    app, ns = guarded_mailbox
    msgs = list(iter_messages(ns.inbox, "Inbox", limit=1))
    assert msgs[0].sender_email == "", "the X500 address is not an email and is dropped"
    assert msgs[0].sender_name == "Dana Whitfield"
    off = list(iter_messages(ns.inbox, "Inbox", limit=1, addresses=False))
    assert off[0].sender_email == ""


def test_odd_subject_makes_a_safe_file_name(tmp_path, mailbox):
    app, ns = mailbox
    dump(ns, tmp_path, ["Inbox"], full=True)
    odd = [f for f in _files(tmp_path) if "Q3-budget" in f.name]
    assert len(odd) == 1
    assert not set('<>:"/\\|?*') & set(odd[0].name)


def test_resolve_folder_walks_a_path(mailbox):
    app, ns = mailbox
    projects = ns.inbox.add_subfolder(FakeFolder("Projects"))
    assert resolve_folder(ns, "Inbox") is ns.inbox
    assert resolve_folder(ns, "Sent Items") is ns.sent
    assert resolve_folder(ns, "Inbox/Projects") is projects
    assert resolve_folder(ns, "Inbox\\Projects") is projects
    with pytest.raises(Exception):
        resolve_folder(ns, "Inbox/Nope")


def test_partial_state_is_saved_during_a_long_run(tmp_path, mailbox):
    app, ns = mailbox
    for i in range(250):
        ns.inbox.Items.append(FakeMail(f"BULK{i}", f"bulk {i}", "Bulk", "b@example.edu",
                                       "me@example.edu", dt.datetime(2024, 1, 1) + dt.timedelta(minutes=i), "x"))
    log = []
    dump(ns, tmp_path, ["Inbox"], full=True, log=log.append)
    assert any("200 written" in line for line in log)
    assert State(tmp_path).counts() == {"Inbox": N_INBOX + 250}


def test_connect_reports_a_missing_outlook():
    """On Linux pywin32 is absent; on a Windows runner Outlook is. Both are exit-2 material."""
    import sys
    if sys.platform == "win32":
        pytest.importorskip("win32com.client")
    with pytest.raises(OutlookUnavailable):
        from outlook_lite.dump import connect
        connect()
