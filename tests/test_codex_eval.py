"""The Codex eval as a test. Opt in with RUN_CODEX_EVAL=1: it needs the
Codex CLI, a login, and spends quota."""
import os
import shutil

import pytest

import eval_codex


@pytest.mark.skipif(os.environ.get("RUN_CODEX_EVAL") != "1", reason="set RUN_CODEX_EVAL=1")
def test_codex_answers_from_the_corpus(tmp_path):
    if shutil.which(eval_codex.CODEX) is None:
        pytest.skip("codex not on PATH")
    eval_codex.build_corpus(tmp_path)
    failures = []
    for q, must, must_not in eval_codex.QUESTIONS:
        answer = eval_codex.ask(tmp_path, q)
        problems = eval_codex.grade(answer, must, must_not)
        if problems:
            failures.append(f"{q}\n  -> {'; '.join(problems)}\n  answer: {answer[:500]!r}")
    assert not failures, "\n".join(failures)


def test_grade():
    assert eval_codex.grade("Use the Portal with the speedtype.", ["portal", "SpeedType"], ["p-card"]) == []
    assert eval_codex.grade("Use the p-card.", ["portal"], ["p-card"]) == ["missing 'portal'", "contains 'p-card'"]
    assert eval_codex.grade("I could not find it.", ["not in the mail|could not find"], []) == []
