import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from fake_outlook import mailbox_from_json  # noqa: E402

FIXTURE = ROOT / "tests" / "fixtures" / "mailbox.json"


@pytest.fixture
def mailbox():
    """(app, namespace) for the fixture mailbox."""
    return mailbox_from_json(FIXTURE)


@pytest.fixture
def guarded_mailbox():
    """The same mailbox with the Object Model Guard refusing address reads."""
    return mailbox_from_json(FIXTURE, guard=True)
