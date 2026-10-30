import datetime as dt
from pathlib import Path

from outlook_lite.render import (Message, State, clean_body, message_path, read_headers,
                                 render_markdown, short_id, slug)


def _msg(**kw):
    base = dict(entry_id="EID1", folder="Inbox", subject="Hello", sender_name="Finance Office",
                sender_email="finance@example.edu", to="me@example.edu", cc="",
                received=dt.datetime(2026, 9, 15, 14, 3), body="Hi\r\n\r\n\r\n\r\nBye  ")
    base.update(kw)
    return Message(**base)


def test_slug_strips_windows_reserved_characters():
    s = slug('Re: Q3 budget?? "final" / draft: v2 <please review>')
    assert not set('<>:"/\\|?*') & set(s)
    assert s == "Re-Q3-budget-final-draft-v2-please-revie"
    assert slug("") == "untitled"
    assert slug("   ...  ") == "untitled"
    assert len(slug("x" * 200)) == 40


def test_message_path_is_relative_and_dated(tmp_path):
    p = message_path(tmp_path, _msg())
    rel = p.relative_to(tmp_path)
    assert rel.parts[:3] == ("mail", "Inbox", "2026")
    assert rel.name.startswith("2026-09-15_1403_Finance-Office_Hello_")
    assert rel.name.endswith(short_id("EID1") + ".md")
    assert len(rel.name) < 120


def test_message_path_without_date_does_not_crash(tmp_path):
    p = message_path(tmp_path, _msg(received=None))
    assert "1970" in str(p)


def test_clean_body_caps_blank_lines_and_trailing_space():
    assert clean_body("Hi\r\n\r\n\r\n\r\nBye  ") == "Hi\n\n\nBye\n"
    assert clean_body("") == "\n"


def test_render_and_read_headers_round_trip(tmp_path):
    m = _msg(cc="chair@example.edu", attachments=["a.pdf", "b.xlsx"], conversation_id="CONV1")
    text = render_markdown(m)
    assert text.startswith("# Hello\n")
    assert "- From: Finance Office <finance@example.edu>" in text
    assert "- Attachments: a.pdf, b.xlsx" in text
    assert text.endswith("Hi\n\n\nBye\n")
    f = tmp_path / "m.md"
    f.write_text(text, encoding="utf-8")
    h = read_headers(f)
    assert h["EntryID"] == "EID1"
    assert h["ConversationID"] == "CONV1"
    assert h["Date"] == "2026-09-15 14:03"
    assert h["CC"] == "chair@example.edu"


def test_render_without_address_keeps_name_only():
    text = render_markdown(_msg(sender_email=""))
    assert "- From: Finance Office\n" in text


def test_state_records_watermark_and_survives_reload(tmp_path):
    st = State(tmp_path)
    m1 = _msg(entry_id="A", received=dt.datetime(2026, 1, 1, 9))
    m2 = _msg(entry_id="B", received=dt.datetime(2026, 2, 1, 9))
    st.record(m1, message_path(tmp_path, m1))
    st.record(m2, message_path(tmp_path, m2))
    st.record(m1, message_path(tmp_path, m1))      # older again: watermark stays
    st.save()
    again = State(tmp_path)
    assert again.watermark("Inbox") == dt.datetime(2026, 2, 1, 9)
    assert again.has("Inbox", "A") and again.has("Inbox", "B") and not again.has("Inbox", "C")
    assert again.counts() == {"Inbox": 2}
    assert "\\" not in next(iter(again.folder("Inbox")["seen"].values()))
    assert not (tmp_path / "mail" / "state.tmp").exists()
