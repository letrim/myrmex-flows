# MYRMEX flows — how to write, test and ship them

This folder holds **Python flows** for Epsilon HyperM (ΦΜΥ, ΑΠΔ, …) that the
MYRMEX engine (`myrmex.exe`) runs. Here you write and test flows; the engine
itself is not here and is never changed from here. Anything the engine lacks is
a request to Lefteris, not code in a flow.

## Setup: where `myrmex.exe` comes from

Use the **installed agent**, not a copy in git:

- Install it once with the signed installer: MYRMEX cloud → Admin →
  «Εκδόσεις agent» → download the latest `MyrmexSetup-<version>.exe`. It
  installs to `C:\Myrmex\` (`myrmex.exe` = the command line, `MyrmexAgent.exe`
  = the tray). Later versions: tray → «Check for updates…».
- Every office PC already has the same `C:\Myrmex\myrmex.exe` (the agent is
  installed there), so a flow can be tested on any of them **without git, Python
  or anything else**: copy the `.py` (and a scenario) to the PC, run it.
- The tray shows the version (`Check for updates… (v0.3.2)`). Each run's log
  starts with `Flow started: … (Python, Flow SDK x.y.z)`.

This folder is the git repo **myrmex-flows** (only text; flows, docs, scenarios):

```
C:\myrmex-flows\
  CLAUDE.md                 this file
  FLOW_SDK_REFERENCE.md     every f.* command — comes with each agent release
  YAML_REFERENCE.md         what each parameter does (YAML examples, same keys)
  flows\ΑΠΔ.py, flows\ΦΜΥ.py
  scenarios\<flow>\*.json   dry-run scenarios, one per branch
  examples\*.py             small probes (dialog_text_probe, checkbox_probe)
```

Runs write `myrmex.log`, `logs\<run>\` and `screenshots\` into the folder they
are started from: keep them out of git.

### Which SDK features may a flow use?

`MIN_SDK_VERSION` must be what the agent on the target PCs has. A flow that
needs more is refused with «needs Flow SDK x; this agent has y. Update the agent».

| Flow SDK | Agent | Adds |
|---|---|---|
| 1.0 | 0.3.2 | everything in FLOW_SDK_REFERENCE.md of 0.3.2 |
| 1.1 | 0.3.3 | `Param(..., secret=True)`; `f.click(..., hover=, hold=)`; primary mouse button; `relative_coordinates` from the connected window |
| 1.2 | 0.4.0 | `expect_hint` (toolbar tooltips), `f.downloads_folder()` |
| 1.3 | 0.4.1 | `f.fail(..., retryable=False)`, `f.delete_file`, `f.dialog_text` |
| 1.4 | 0.4.2 | `f.attach` (the receipt to the cloud) |
| 1.5 | 0.4.4 | `f.save_dialog` (the save dialog's file name field — Epsilon Extra), `f.set_checkbox` / `f.checkbox_state` of a `TwwCheckBox` from its picture |

The cloud admin sets «Ελάχιστη έκδοση agent» per flow version: an older agent never gets it.
Until 0.4.4 is installed everywhere: `MIN_SDK_VERSION = "1.4"` and none of the 1.5 items.

## What a flow is

```python
from myrmex_agent.sdk import Param

FLOW_VERSION = "1"                  # the cloud version it will be uploaded as
MIN_SDK_VERSION = "1.0"
RESET_PROCESSES = ["Acrobat.exe", "AcroRd32.exe"]   # killed before/after every client (+ HyperM, firefox)
PARAMS = {
    "period.year": int,
    "period.month": Param(int, choices=range(1, 13)),
    "client.afm": Param(str, pattern=r"\d{9}"),
    "client.name": str,
    "output.folder": str,
    "test": Param(str, required=False, choices=("yes", "no")),
}

def run(f, ctx):
    with f.phase("open_company"):          # phases: log grouping + --from target
        f.navigate_menu("Δημιουργία - διαχείριση εταιριών")
        ...
```

The engine loads the file, checks PARAMS, and calls `run(f, ctx)` once per
client. Batch mode, logs, `summary.json`, failure screenshots, closing
HyperM/Firefox between clients and the cloud report are the engine's job.

- Every action is `f.<action>(…)`: the full list is in `FLOW_SDK_REFERENCE.md`,
  what each parameter does in `YAML_REFERENCE.md`. **Don't invent parameters**:
  an unknown one is an error.
- `ctx["client.afm"]`, `ctx.get("test")`. PARAMS values come converted
  (`period.month` is an int).
- Control flow is plain Python (`if`, `for`, functions, f-strings).
- Outcome of a client:
  - `return f.skip("…")` → skipped
  - `return f.done("0 ΕΓΓΡΑΦΕΣ — …")` → success with a message (status word first)
  - `f.fail("ΑΣΥΜΦΩΝΙΑ …")` → failed
  - `f.check(cond, "msg")` → fails unless `cond` is true
  - `return f.halt("…")` → stop for a person: nothing closed, the run stops
- `f.retry(fn, attempts=3, delay=10)` where `fn(f)` = «do the step if not done
  yet, then verify it».
- `f.window_exists(cls, title, timeout=…)` → bool, also inside Epsilon. Reads
  return values: `f.get_control_text(…)`, `f.excel_read(…)`, `f.browser_page()`,
  `f.dialog_buttons(…)`, `f.control_exists(…)`, `f.uia_page(…)`.
- A password: `Param(str, required=False, secret=True)` (SDK 1.1) — `***` in every
  log, result and trace. Type it with `f.fill(..., fill_method="message")`; never
  `f.log` / `f.set_result` it.

## Hard rules

1. **No `win32gui`, `pywinauto`, `pyautogui`, `ctypes`, `os`, `subprocess`,
   `time`, `open()`, `eval`, or `__dunder__` in a flow.** Allowed imports:
   `myrmex_agent.sdk`, `re`, `datetime`, `decimal`, `math`, `string`,
   `itertools`, `functools`, `collections`, `collections.abc`, `dataclasses`,
   `enum`, `typing`, `json`, `unicodedata`, and
   `from pathlib import PurePath, PureWindowsPath`.
   Anything else is refused at load time. Pauses: `f.wait(s)`.
2. **If something is missing from `f.`, don't work around it.** Say which command
   is needed and why, as a request for Lefteris (the engine gets it, every flow
   can use it).
3. **Submission in `ΦΜΥ.py`:** the cloud sends `submit=yes|no` with **every**
   job — only `yes` submits, any other value is a test (stops right before
   ΜΕΤΑΦΟΡΤΩΣΗ). Without `submit` (a local run) it submits; `test=yes` always
   wins. **Never remove the `submit` param or fail on it**: every cloud job would
   fail. A flow must never turn a missing or unexpected value into a submission.
4. **Clicks that can't be undone** (ΜΕΤΑΦΟΡΤΩΣΗ, Οριστικοποίηση, final Υποβολή)
   **never inside `f.retry`**: one click whose `timeout` covers the attempts,
   inside `with f.time_limit(…):`.
5. **Right after the final Υποβολή: `f.committed("υποβλήθηκε στην ΑΑΔΕ")`.**
6. **Never answer a dialog the flow didn't expect.** Answer only dialogs
   identified by class and title; anything else → photo and `f.fail`.
7. **Keep the existing comments and reasons** (why by message, why typed,
   dates of the runs that showed it). They come from real offices.

## Works on EVERY office PC — not just yours

Epsilon looks different from PC to PC and version to version. What worked on one
PC proves nothing about the next. Seen on real PCs (03–08/11/2026):

| Seen | So |
|---|---|
| Epsilon's left menu 52 px wider on one PC: every screen position shifted right | **No absolute screen coordinates.** Find controls by class / text / `control_id` / `within` / `prefer`, click by message where possible |
| In insert mode a bar of icons pushed the rights grids 22 px down | When a painted thing must be clicked at a position: `method="relative_coordinates"` from the **nearest real control** (the grid, the form), never from the screen or the main window |
| «Μισθοδοτική Κατάσταση» in the menu search: 3, 4 or 5 rows depending on PC and user | `f.navigate_menu(item, rows=N, select=…)`: the exact number of rows, else nothing is picked; check the window that opened |
| A new user's column layout: Επωνυμία first on one PC, Α.Φ.Μ. first on another; «Τύπος» the 6th row on one, the last on another | **Never assume a default state** (a box ticked or not, an order, a selection). Read it first; act only if needed; read it again after |
| A checkbox whose state can't be read (painted lists such as the grid's columns popup) | **Don't click it blindly** — a click on a ticked box unticks it. Take a photo and leave it to a person, or ask for an engine command that reads it |
| «Επιλογές Δήλωσης» of the ΦΜΥ are InfoPower `TwwCheckBox`: BM_GETCHECK always 0 | SDK 1.5: `f.set_checkbox(caption, class_name="TwwCheckBox")` reads them from their picture and confirms after the click; `examples\checkbox_probe.py` checks the reading on a PC first |
| Epsilon Extra: the name filled into «the first Edit» of «Αποθήκευση» didn't take (a bare «.pdf») | SDK 1.5: `f.save_dialog(path, title="Αποθήκευση")` — the file name field by id; before it, find the newest file written and rename it |
| One office: mouse buttons swapped — a «left» click was a right click (fixed in the engine, 0.3.3) | Don't reason from «the pointer is on the button»: check the RESULT of every click (the window opened, the field emptied) |
| A click missed and the typing went into the admin's record | Before typing into a form, check it is the form you expect (e.g. it emptied after «Εισαγωγή») |
| Epsilon's «Epsilon Connector» balloon is a `#32770`; a Chrome tab had «Σφάλμα» in its title | Look only at Epsilon's own messages: `TMessageForm`, or `#32770` **with a title** |
| A minimized Epsilon hides `TfMain` | Don't minimize Epsilon; `f.maximize_window` after `f.focus_window` on agent 0.3.2 |
| Firefox of Epsilon: one PC showed no UI Automation elements | Check what you rely on (e.g. `f.browser_focus_uia`) and stop *before* anything irreversible if it's missing |
| Different screen sizes / scaling | If a flow must click positions, `f.check(f.screen_size() == [1920, 1080], "…")` first — and prefer not to need it |

Also: a step's speed differs (server load, PC). Wait for a **condition**
(`f.wait_for`, `f.window_exists(…, timeout=…)`, `f.wait_ready`), not for a fixed
number of seconds.

## Test without signing

Local runs never need a signature (signed or not, a local `.py` runs). From this
folder:

```bat
:: 1) without Epsilon: lint, import check and the whole flow against a mock
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --dry-run
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --dry-run --scenario scenarios\ΦΜΥ\no_employees.json

:: 2) on a PC with Epsilon (Pause the tray agent first). test=yes stops before ΜΕΤΑΦΟΡΤΩΣΗ
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --var test=yes --var period.year=2026 --var period.month=9 --var client.afm=800611083 --var client.name="…" --var output.folder=C:\Myrmex\exports\…

:: batch, resume from a phase
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --clients clients.csv --output-root C:\Myrmex\exports --var period.year=2026 --var period.month=9
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --from export_excel --var …
```

**`--dry-run` must pass for every branch before Epsilon.** For each branch (no
employees, KAD question, mismatch, already submitted, submit with tax due / zero,
a message nobody expected) write a scenario in `scenarios\<flow>\`:

```json
{"description": "no employees this period",
 "vars": {"period.year": "2026", "period.month": "9", "client.afm": "123456789",
          "client.name": "ΤΕΣΤ", "output.folder": "C:\\exports\\x"},
 "windows": [{"class_name": "TMessageForm", "title": "Πληροφορία", "results": [true]}],
 "store":   {"fmy_amount": "834,49", "payroll_rows": [["…"]]},
 "fail":    [{"action": "wait_for_file", "error": "no PDF"}]}
```

`windows` answers `window_exists` (`results` in turn, the last one repeats; no
rule = False; a rule without `title` matches any title). `store` gives what
reads return, by their `store_as` — give each read its own `store_as` so a
scenario can answer it. `fail` makes a matching step fail.

Then on a real PC: first with `test=yes` (or a flow-specific «stop before the
irreversible step» that ends with `f.halt`, leaving Epsilon open to look at).
A failure leaves `logs\<run>\` (the client's log, `summary.json`, a screenshot
and the list of open windows); `f.screenshot()` adds your own photos.
**Every check that can fail should say what it saw** (the field's text, the
buttons of the message, the photo's name) so one run is enough to know why.

## Shipping a flow to the cloud

Cloud agents run a Python flow **only if it is signed** for exactly that name,
version and content. **Lefteris signs** (agreed 25/11/2026): push the flow here
with a new `FLOW_VERSION`, say what changed and the dry-runs / real runs that
passed; he signs and uploads it. A version never changes once uploaded: any
change, even one character, is a new `FLOW_VERSION`. A change that needs the
engine (a new `f.` command) is a request to Lefteris first — no workaround in the
flow. **Never put a key file or its passphrase in git or in a flow.**

## Current work (25/11/2026)

- `flows\ΑΠΔ.py` = cloud «ΑΠΔ» **v5**, `flows\ΦΜΥ.py` = cloud «ΦΜΥ» **v10** (MIN_SDK 1.4,
  agent 0.4.2). The next version starts from these files.
- The cloud's per-company buttons reach the flow as job params, `yes` / `no` like
  `submit`: ΑΠΔ `review.accept`, ΦΜΥ `fmy.amended`, `fmy.total_pause`, `fmy.temp_absence`.
  The cloud splits a job so every job has one value of each. **Never rename these params**
  without telling Lefteris: the cloud's flow settings use the same names.
- Next, once agent 0.4.4 is everywhere (MIN_SDK 1.5): `save_as()` → `f.save_dialog`,
  `declaration_options()` → `f.set_checkbox(…, class_name="TwwCheckBox")` (after
  `checkbox_probe.py` on a real Epsilon).
