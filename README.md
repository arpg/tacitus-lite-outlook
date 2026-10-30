# tacitus-lite-outlook

Ask a coding agent (Codex, Claude Code, or any tool that reads a
folder) questions about your own Outlook mail, on a Windows PC, with
nothing installed on a server and no connector your IT group has to
approve.

Classic Outlook already holds your mailbox on disk. `outlook_dump.py`
reads it through Outlook's own automation interface (COM) and writes
one Markdown file per message into `mail/`. The agent searches those
files. `make_draft.py` puts a reply into Outlook's **Drafts** folder;
you read it in Outlook and click Send yourself. Nothing here can send,
move or delete mail.

Questions it answers well: "What is the right way to pay for visitor
parking with a p-card?", "Did I ever reply to the catering quote?",
"What did finance say the receipt deadline was?", "Draft a reply to
Dana about Friday's speaker".

## What you need

- Windows 10 or 11 with **classic Outlook**, signed in to your account.
  The **New Outlook** (the toggle at the top right of the Outlook window)
  has no automation interface; turn the toggle off. If your organization
  has removed classic Outlook, this tool cannot read your mail.
- Python 3.10 or newer from [python.org](https://www.python.org/downloads/windows/)
  (tick "Add python.exe to PATH" in the installer).
- A coding agent that works in a folder: the
  [Codex CLI](https://github.com/openai/codex) (`npm install -g @openai/codex`,
  needs Node.js) or its VS Code extension with a ChatGPT account, or
  Claude Code with a Claude account.
- `rg` (ripgrep) makes searching fast. The Codex CLI brings its own; for
  anything else, `winget install BurntSushi.ripgrep.MSVC`.

## Install (10 minutes)

Open **PowerShell** (not as administrator) and run, one line at a time:

```powershell
cd $HOME\Documents
git clone https://github.com/arpg/tacitus-lite-outlook.git
cd tacitus-lite-outlook
py -m pip install -r requirements.txt
py outlook_dump.py selftest --draft
```

If you have no `git`, download the ZIP from GitHub's green **Code**
button and unzip it into `Documents` instead of the `git clone` line.

`selftest` prints a `PASS` line for each step: it connects to Outlook,
names your account, opens the Inbox and Sent Items, reads three
messages from each, copies five to a temporary folder, and saves a
draft addressed to you (delete it from Drafts afterwards). The last line
is `ALL PASS`. If a step says `FAIL`, the line names the step; the
usual causes are the New Outlook toggle and Outlook not being open.

If Outlook showed a window saying "A program is trying to access e-mail
address information" during the selftest, click Allow, and add
`--no-addresses` to the scheduled task below: the files then carry
sender names without addresses, and nothing waits on a prompt.

Then copy your mail:

```powershell
py outlook_dump.py --full
```

The first run reads everything and takes a few minutes for a large
mailbox (about a thousand messages a minute). It prints a count per
folder when it finishes. To include a folder of your own, name it with
its path: `py outlook_dump.py --full --folders Inbox "Sent Items" "Inbox/Projects"`.

### Keep it fresh

Run this once; it registers a Windows scheduled task that copies new
mail every 30 minutes while you are logged in:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup-task.ps1
```

`.\setup-task.ps1 -Remove` deletes the task. `py outlook_dump.py status`
shows what is on disk.

## Ask questions

In PowerShell, in the same folder:

```powershell
codex
```

(or open the folder in VS Code with the Codex extension, or run
`claude`). `AGENTS.md` in this folder tells the agent how the files are
laid out, to search before answering, to prefer the newest message from
the office that owns a rule, and to cite the file it answered from. Ask
in plain words:

> What is the right way to pay for visitor parking using a p-card or a t-card?

> Did I answer Dana about the seminar speaker's parking?

> Draft a reply to the catering quote saying yes to 40 guests.

A draft goes to Outlook > Drafts. Read it there and press Send.

## Privacy

Your mail stays on your PC. The agent sends the parts it reads to its
provider (OpenAI for Codex, Anthropic for Claude Code) to answer your
question, the same as pasting that email into the chat would. Check that
this is acceptable under your organization's rules for the content of
your mailbox before the first run; if an Outlook connector for your
agent was declined, ask why before using this.

`mail/`, `notes/` and `drafts/` are listed in `.gitignore`, so a
`git pull` for an update never touches them and nothing in them can be
pushed by accident.

## How it is tested

The code that reads Outlook is a thin layer; everything else is plain
Python with a test suite that runs on Linux, macOS and Windows.

| What | Where | Proves |
|---|---|---|
| `tests/test_render.py`, `test_dump.py`, `test_draft.py` | any OS, `py -m pytest` | file names, Markdown, incremental runs, drafts, exit codes, against a fake Outlook object model (`tests/fake_outlook.py`) that mirrors the real one's shapes |
| GitHub Actions, `windows-latest` | CI | the same suite on Windows, pywin32 installs and imports, and the scripts answer exit 2 cleanly when no Outlook is present |
| `tests/eval_codex.py` | a machine with the Codex CLI | the agent, given the fixture mailbox and `AGENTS.md`, answers the parking question with the newer rule and cites the right file |
| `py outlook_dump.py selftest --draft` | the PC with Outlook | the real COM calls: connect, account, folders, read, dump, draft |

Only the last row needs Outlook. Run it on any Windows machine with
classic Outlook signed in to a mailbox; the fixture suite and CI cover
the rest before the code gets there.

## License

MIT. Fictional names and policies in `tests/fixtures/`.
