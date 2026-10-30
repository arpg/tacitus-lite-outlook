"""Pure logic: a message record, its file name, its Markdown, the state file.

Nothing here imports COM, so every function is tested on Linux.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path

MAIL_DIR = "mail"
STATE_FILE = "state.json"
# Characters Windows refuses in a file name, plus a few that confuse shells.
_BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f\'`$&;]+')
_WS = re.compile(r"\s+")


@dataclass
class Message:
    entry_id: str
    folder: str
    subject: str
    sender_name: str
    sender_email: str
    to: str
    cc: str
    received: dt.datetime | None
    body: str
    conversation_id: str = ""
    attachments: list[str] = field(default_factory=list)

    def to_json(self) -> dict:
        d = asdict(self)
        d["received"] = self.received.isoformat(timespec="seconds") if self.received else None
        return d


def slug(text: str, limit: int = 40) -> str:
    """A short, file-name-safe form of a subject or a name."""
    text = _BAD.sub(" ", text or "")
    text = _WS.sub("-", text.strip()).strip("-.")
    if not text:
        text = "untitled"
    return text[:limit].rstrip("-.")


def short_id(entry_id: str) -> str:
    """Eight hex characters that identify a message after a subject clash."""
    return hashlib.sha1(entry_id.encode("utf-8")).hexdigest()[:8]


def message_path(root: Path, msg: Message) -> Path:
    """mail/<folder>/<year>/<date_time>_<sender>_<subject>_<id>.md"""
    when = msg.received or dt.datetime(1970, 1, 1)
    stamp = when.strftime("%Y-%m-%d_%H%M")
    name = f"{stamp}_{slug(msg.sender_name, 24)}_{slug(msg.subject)}_{short_id(msg.entry_id)}.md"
    return root / MAIL_DIR / slug(msg.folder, 24) / f"{when.year:04d}" / name


def clean_body(body: str) -> str:
    """Normalise line endings, trim trailing spaces, cap blank runs at two."""
    body = (body or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in body.split("\n")]
    out: list[str] = []
    blank = 0
    for ln in lines:
        blank = blank + 1 if not ln else 0
        if blank <= 2:
            out.append(ln)
    return "\n".join(out).strip() + "\n"


def render_markdown(msg: Message) -> str:
    sender = msg.sender_name or msg.sender_email or "unknown"
    if msg.sender_email and msg.sender_email.lower() not in sender.lower():
        sender = f"{sender} <{msg.sender_email}>"
    date = msg.received.strftime("%Y-%m-%d %H:%M") if msg.received else "unknown"
    lines = [
        f"# {msg.subject or '(no subject)'}",
        "",
        f"- From: {sender}",
        f"- To: {msg.to or ''}",
    ]
    if msg.cc:
        lines.append(f"- CC: {msg.cc}")
    lines += [
        f"- Date: {date}",
        f"- Folder: {msg.folder}",
    ]
    if msg.attachments:
        lines.append(f"- Attachments: {', '.join(msg.attachments)}")
    lines += [
        f"- EntryID: {msg.entry_id}",
    ]
    if msg.conversation_id:
        lines.append(f"- ConversationID: {msg.conversation_id}")
    lines += ["", "---", "", clean_body(msg.body)]
    return "\n".join(lines)


_HEADER = re.compile(r"^- (\w+): ?(.*)$")


def read_headers(path: Path) -> dict[str, str]:
    """The `- Key: value` lines above the first `---` of a dumped message."""
    out: dict[str, str] = {}
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.rstrip("\n")
            if line.strip() == "---":
                break
            m = _HEADER.match(line)
            if m:
                out[m.group(1)] = m.group(2)
    return out


class State:
    """Which EntryIDs are on disk, per folder, and the newest time seen.

    `seen` maps EntryID to the file's path relative to the root, so a
    second run skips what it has and a rebuild can find every file.
    """

    def __init__(self, root: Path):
        self.root = Path(root)
        self.path = self.root / MAIL_DIR / STATE_FILE
        self.folders: dict[str, dict] = {}
        if self.path.exists():
            with open(self.path, encoding="utf-8") as fh:
                data = json.load(fh)
            self.folders = data.get("folders", {})

    def folder(self, name: str) -> dict:
        return self.folders.setdefault(name, {"watermark": None, "seen": {}})

    def watermark(self, name: str) -> dt.datetime | None:
        wm = self.folder(name).get("watermark")
        return dt.datetime.fromisoformat(wm) if wm else None

    def has(self, name: str, entry_id: str) -> bool:
        return entry_id in self.folder(name)["seen"]

    def record(self, msg: Message, path: Path) -> None:
        f = self.folder(msg.folder)
        f["seen"][msg.entry_id] = str(path.relative_to(self.root)).replace("\\", "/")
        if msg.received:
            wm = self.watermark(msg.folder)
            if wm is None or msg.received > wm:
                f["watermark"] = msg.received.isoformat(timespec="seconds")

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump({"version": 1, "folders": self.folders}, fh, indent=1)
        tmp.replace(self.path)

    def counts(self) -> dict[str, int]:
        return {name: len(f["seen"]) for name, f in self.folders.items()}
