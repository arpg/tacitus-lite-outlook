#!/usr/bin/env python3
"""Ask the Codex CLI the fixture questions and grade the answers.

    python3 tests/eval_codex.py            # prints each answer and PASS/FAIL
    python3 tests/eval_codex.py --keep     # leave the corpus folder behind

The corpus is built from `tests/fixtures/mailbox.json` through the real
dump code (with the fake Outlook), so the files are byte for byte what
`outlook_dump.py` writes. `AGENTS.md` is copied beside them. Codex runs
in its read-only sandbox with `-C` on that folder. Needs `codex login`
on this machine; the run spends that account's quota.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from fake_outlook import mailbox_from_json  # noqa: E402
from outlook_lite.dump import dump  # noqa: E402

CODEX = os.environ.get("CODEX_BIN", "codex")

# (question, substrings the answer must contain, substrings it must not contain).
# A "must" entry with "|" is satisfied by any one of its alternatives.
QUESTIONS = [
    ("What is the right way to pay for visitor parking? Can I use the p-card or the t-card?",
     ["portal", "SpeedType", "visitor-parking-validations"],
     []),
    ("Did I answer Dana about the seminar speaker's parking? What did I say?",
     ["1234567", "Sent-Items"],
     []),
    ("When are T-card receipts due?",
     ["30 days", "T-card-receipts"],
     []),
    ("What did the dean say about the travel budget for next year?",
     ["not in the mail|couldn't find|could not find|no message|nothing in|no email|did not find"],
     ["$"]),
]


def build_corpus(dst: Path) -> None:
    app, ns = mailbox_from_json(ROOT / "tests" / "fixtures" / "mailbox.json")
    dump(ns, dst, ["Inbox", "Sent Items"], full=True)
    shutil.copy(ROOT / "AGENTS.md", dst / "AGENTS.md")
    (dst / "notes").mkdir()
    (dst / "drafts").mkdir()


def ask(cwd: Path, question: str, timeout: int = 300) -> str:
    out = cwd / ".last-answer.txt"
    cmd = [CODEX, "exec", "--skip-git-repo-check", "--color", "never", "--sandbox", "read-only",
           "-C", str(cwd), "-o", str(out), question]
    subprocess.run(cmd, check=False, timeout=timeout, capture_output=True, text=True)
    return out.read_text(encoding="utf-8") if out.exists() else ""


def grade(answer: str, must: list[str], must_not: list[str]) -> list[str]:
    low = answer.lower()
    problems = [f"missing {s!r}" for s in must
                if not any(alt.lower() in low for alt in s.split("|"))]
    problems += [f"contains {s!r}" for s in must_not if s.lower() in low]
    return problems


def main(argv=None) -> int:
    keep = "--keep" in (argv or sys.argv[1:])
    if shutil.which(CODEX) is None:
        print(f"{CODEX} is not on PATH")
        return 2
    tmp = Path(tempfile.mkdtemp(prefix="outlook-lite-eval-"))
    build_corpus(tmp)
    print(f"corpus: {tmp}")
    failures = 0
    for q, must, must_not in QUESTIONS:
        answer = ask(tmp, q)
        problems = grade(answer, must, must_not)
        print("=" * 72)
        print(f"Q: {q}")
        print(answer.strip() or "(no answer)")
        print("PASS" if not problems else "FAIL: " + "; ".join(problems))
        failures += bool(problems)
    if not keep:
        shutil.rmtree(tmp, ignore_errors=True)
    print("=" * 72)
    print(f"{len(QUESTIONS) - failures}/{len(QUESTIONS)} passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
