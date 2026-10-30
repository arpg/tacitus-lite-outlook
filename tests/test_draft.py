from pathlib import Path

import pytest

import make_draft
import outlook_dump
from outlook_lite.draft import body_html, create_draft, entry_id_of, reply_draft
from outlook_lite.dump import OutlookUnavailable, dump


def test_body_html_escapes_and_paragraphs():
    h = body_html("Hi Dana,\n\nUse the portal & SpeedType <1234567>.\nThanks")
    assert h == "<p>Hi Dana,</p>\n<p>Use the portal &amp; SpeedType &lt;1234567&gt;.<br>Thanks</p>\n"


def test_reply_all_draft_lands_in_drafts_with_thread_quoted(mailbox):
    app, ns = mailbox
    eid = reply_draft(ns, "EID-COLLEAGUE-SPEAKER", "Use the portal.")
    assert ns.drafts.Items.Count == 1
    d = ns.drafts.Items.Item(1)
    assert d.EntryID == eid and d.saved
    assert d.Subject == "RE: Parking for Friday's seminar speaker?"
    assert d.HTMLBody.startswith("<p>Use the portal.</p>\n<hr>")
    assert "seminar speaker drives in on Friday" in d.HTMLBody
    assert "dana.whitfield@example.edu" in d.To and "me@example.edu" in d.To


def test_sender_only_reply(mailbox):
    app, ns = mailbox
    reply_draft(ns, "EID-BUDGET-ODD-SUBJECT", "Thanks.", reply_all=False)
    d = ns.drafts.Items.Item(1)
    assert d.To == "priya.nair@example.edu"


def test_create_draft(mailbox):
    app, ns = mailbox
    create_draft(app, "a@example.edu", "Hello", "Body", cc="b@example.edu")
    d = ns.drafts.Items.Item(1)
    assert (d.To, d.CC, d.Subject, d.Body, d.saved) == ("a@example.edu", "b@example.edu", "Hello", "Body", True)


def test_entry_id_comes_from_the_dumped_file(tmp_path, mailbox):
    app, ns = mailbox
    dump(ns, tmp_path, ["Inbox"], full=True)
    f = next(p for p in tmp_path.rglob("*.md") if "seminar-speaker" in p.name)
    assert entry_id_of(f) == "EID-COLLEAGUE-SPEAKER"
    bad = tmp_path / "x.md"
    bad.write_text("# nothing\n", encoding="utf-8")
    with pytest.raises(ValueError):
        entry_id_of(bad)


def test_make_draft_cli_reply(tmp_path, mailbox, monkeypatch, capsys):
    app, ns = mailbox
    dump(ns, tmp_path, ["Inbox"], full=True)
    f = next(p for p in tmp_path.rglob("*.md") if "seminar-speaker" in p.name)
    body = tmp_path / "reply.txt"
    body.write_text("Portal, SpeedType 1234567.\n", encoding="utf-8")
    monkeypatch.setattr(make_draft, "connect", lambda: (app, ns))
    assert make_draft.main(["--reply-to", str(f), "--body-file", str(body)]) == 0
    out = capsys.readouterr().out
    assert "Draft saved in Outlook > Drafts: reply all" in out
    assert "click Send yourself" in out
    assert ns.drafts.Items.Count == 1


def test_make_draft_cli_refuses_empty_body_and_no_target(tmp_path, capsys):
    empty = tmp_path / "e.txt"
    empty.write_text("  \n", encoding="utf-8")
    assert make_draft.main(["--to", "a@example.edu", "--body-file", str(empty)]) == 1
    full = tmp_path / "f.txt"
    full.write_text("hi", encoding="utf-8")
    assert make_draft.main(["--body-file", str(full)]) == 1


def test_make_draft_cli_without_outlook_is_exit_2(tmp_path, monkeypatch):
    full = tmp_path / "f.txt"
    full.write_text("hi", encoding="utf-8")

    def boom():
        raise OutlookUnavailable("nope")
    monkeypatch.setattr(make_draft, "connect", boom)
    assert make_draft.main(["--to", "a@example.edu", "--subject", "s", "--body-file", str(full)]) == 2


def test_selftest_passes_against_the_fake(tmp_path, mailbox, monkeypatch, capsys):
    app, ns = mailbox
    monkeypatch.setattr(outlook_dump, "connect", lambda: (app, ns))
    assert outlook_dump.main(["selftest", "--draft"]) == 0
    out = capsys.readouterr().out
    assert "ALL PASS" in out
    assert out.count("FAIL") == 0
    assert "PASS  signed-in account: me@example.edu" in out
    assert ns.drafts.Items.Count == 1


def test_dump_and_status_cli_against_the_fake(tmp_path, mailbox, monkeypatch, capsys):
    app, ns = mailbox
    monkeypatch.setattr(outlook_dump, "connect", lambda: (app, ns))
    monkeypatch.setenv("TACITUS_LITE_HOME", str(tmp_path))
    assert outlook_dump.main(["dump", "--full"]) == 0
    assert outlook_dump.main(["status"]) == 0
    out = capsys.readouterr().out
    assert "Inbox: 10 new" in out and "Sent Items: 2 new" in out
    assert "Inbox: 10 messages, newest 2026-09-23T09:02:00" in out


def test_dump_cli_without_outlook_is_exit_2(monkeypatch, capsys):
    def boom():
        raise OutlookUnavailable("nope")
    monkeypatch.setattr(outlook_dump, "connect", boom)
    assert outlook_dump.main(["dump"]) == 2
    assert outlook_dump.main(["selftest"]) == 2
