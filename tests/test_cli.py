import io
import sys
from pathlib import Path

import make_draft
import outlook_dump
import tacitus_lite
from outlook_lite import paths


def test_home_from_env_then_source_root(tmp_path, monkeypatch):
    monkeypatch.delenv(paths.HOME_ENV, raising=False)
    assert paths.home() == Path(__file__).resolve().parent.parent
    monkeypatch.setenv(paths.HOME_ENV, str(tmp_path))
    assert paths.home() == tmp_path.resolve()


def test_home_when_frozen_is_documents_outlookmail(monkeypatch):
    monkeypatch.delenv(paths.HOME_ENV, raising=False)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    assert paths.home() == paths.documents() / "OutlookMail"
    assert paths.home().name == "OutlookMail"


def test_dispatch_routes_to_the_two_scripts(monkeypatch, capsys):
    seen = []
    monkeypatch.setattr(outlook_dump, "main", lambda argv: seen.append(("dump", argv)) or 0)
    monkeypatch.setattr(make_draft, "main", lambda argv: seen.append(("draft", argv)) or 0)
    assert tacitus_lite.main(["dump", "--full"]) == 0
    assert tacitus_lite.main(["status"]) == 0
    assert tacitus_lite.main(["selftest", "--draft"]) == 0
    assert tacitus_lite.main(["draft", "--to", "a@example.edu", "--body-file", "x"]) == 0
    assert seen == [("dump", ["dump", "--full"]), ("dump", ["status"]), ("dump", ["selftest", "--draft"]),
                    ("draft", ["--to", "a@example.edu", "--body-file", "x"])]


def test_help_version_home_and_unknown(tmp_path, monkeypatch, capsys):
    monkeypatch.setenv(paths.HOME_ENV, str(tmp_path))
    assert tacitus_lite.main([]) == 0
    assert tacitus_lite.main(["version"]) == 0
    assert tacitus_lite.main(["home"]) == 0
    assert tacitus_lite.main(["bogus"]) == 1
    out = capsys.readouterr().out
    assert "tacitus-lite dump" in out and "tacitus-lite 0." in out and str(tmp_path.resolve()) in out
    assert "unknown command 'bogus'" in out


def test_pause_waits_for_enter(monkeypatch, capsys):
    monkeypatch.setattr(outlook_dump, "main", lambda argv: 0)
    monkeypatch.setattr("builtins.input", lambda prompt="": "")
    assert tacitus_lite.main(["status", "--pause"]) == 0


def test_windowless_run_logs_to_dump_log(tmp_path, monkeypatch):
    """The scheduled task runs the windowless build: no console, so a log file."""
    monkeypatch.setenv(paths.HOME_ENV, str(tmp_path))
    monkeypatch.setattr(outlook_dump, "main", lambda argv: print("Outlook not available: test") or 2)
    monkeypatch.setattr(sys, "stdout", None)
    monkeypatch.setattr(sys, "stderr", None)
    try:
        assert tacitus_lite.main(["dump"]) == 2
    finally:
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
    log = (tmp_path / "mail" / "dump.log").read_text(encoding="utf-8")
    assert "tacitus-lite dump" in log and "Outlook not available: test" in log
