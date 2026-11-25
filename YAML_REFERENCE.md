# MYRMEX Agent — Action Reference

What every action does and every key it takes. **Flows are Python only**
(Flow SDK, `def run(f, ctx)`): YAML flows don't run since agent 0.3.1, locally
or from the cloud. The examples below are still written as the old YAML steps;
`- action: X` with keys `k: v` is `f.X(k=v)` in a flow (see
`docs/FLOW_SDK_REFERENCE.md` for the exact signatures).

## Running a workflow

**Single client:**
```
myrmex.exe flows\ΦΜΥ.py
myrmex.exe flows\ΦΜΥ.py params.txt
myrmex.exe flows\ΦΜΥ.py --var client.afm=123456789 --var period.year=2026
```

**Batch (multiple clients from CSV):**
```
myrmex.exe flows\ΦΜΥ.py --clients clients.csv --output-root C:\exports --var period.year=2026 --var period.month=3
```

See the [Batch execution](#batch-execution) section for full details.

From source:
```
python -m myrmex_agent flows\ΦΜΥ.py
python -m myrmex_agent flows\ΦΜΥ.py --var key=value
```

---

## CLI Reference

### Params file

If no params file is passed, the agent looks for `params.txt` next to the executable
(or in the current directory when running from source).

Format — `KEY = VALUE`, `#` comments, dot-notation for nested keys:
```
# params.txt
client.afm    = 123456789
period.year   = 2026
period.month  = 3
output.folder = C:\exports\2026
```

`--var` overrides always win over the params file.

---

### `--dry-run`

Validate the flow and print every action without executing anything.
No windows are opened, no files are written.

```
myrmex.exe flows\ΦΜΥ.py --dry-run
```

**What it checks:**
- YAML is structurally valid
- Every `action:` value is a known, registered action
- Every step has the keys its action needs (e.g. `set_var` needs `key` and `value`;
  `if` needs `condition` or `window_exists`; `method: coordinates` needs `x` and `y`).
  A key left blank (`key:` with nothing after it) counts as missing.
- All nested step blocks (`then`, `else`, `steps`, `catch`) are valid recursively

**Output:**

```
Flow: basic_fmy.yaml  (67 top-level steps)

[0]  set_keyboard_layout
[1]  lookup
[2]  open_app  (C:\...\HyperM.exe)
...
[66] if  (window_exists: TMessageForm, "Πληροφορία")
     then:
       [0] connect_window  ('Πληροφορία')
       [1] click
       ...
     else:
       [0] wait
       ...

────────────────────────────────────────────────────────
Validation OK — 273 actions total, 0 errors.
```

A failed validation looks like:
```
[FAIL] steps[12]: unknown action 'conect_window'
[FAIL] steps[45].then[3]: missing 'action' key
[FAIL] steps[52].else[44].then[0] (set_var): missing required key(s): key
```

---

### `--from PHASE`

Start at a phase of the flow (`with f.phase("…")`): every step before it is
skipped. Variables that skipped steps would have read can be given with `--var`.

```
myrmex.exe flows\ΦΜΥ.py --from export_excel --var client.afm=123456789 --var period.year=2026 --var period.month=3 --var output.folder=C:\exports
```

---

## Template variables

Use `{{ variable }}` anywhere in string values. Dot-notation accesses nested fields:
```yaml
value: "{{ output.folder }}\\{{ output.filename }}.xlsx"
```

Variables are set by `--var` on the CLI, by `set_var` steps, or as outputs of other actions.

Built in: `{{ now }}` — current time as `YYYYMMDD_HHMMSS` (for file names);
`{{ now_epoch }}` — seconds since 1970 (for `find_file newer_than`).

---

## Keys every step accepts

| Key | Meaning |
|---|---|
| `on_error: continue` | a failure of this step is logged as a warning and the flow goes on (default `fail`) |
| `max_seconds: N` | the engine stops waiting for this step after N seconds and **fails the client** (ΑΛΛΑΓΕΣ 1.1: «ΜΕΤΑΦΟΡΤΩΣΗ» once hung 14–19 minutes, and `timeout` couldn't help because the call itself was stuck). A step stuck inside Windows can't be interrupted, only abandoned, so it is **never** retried, caught by `try` or continued with `on_error: continue`. The reset after the client closes the applications, which ends the stuck call. |

```yaml
- action: browser_click_uia
  text: "ΜΕΤΑΦΟΡΤΩΣΗ"
  timeout: 60          # look for the button for up to 60s
  max_seconds: 120     # but never let the whole step take more than 120s
```

Python flows: `with f.time_limit(120): …` gives every step inside `max_seconds: 120`.

## Window Management

### `open_app`
Launch an executable and wait for its window to appear.

```yaml
- action: open_app
  path: "C:\\Program Files\\Epsilon\\Epsilon.exe"
  window_class: TfAuthentication   # Win32 class of the expected window (optional)
  window_title: "Epsilon Smart"    # substring match in title (optional)
  wait_ready: true                 # wait for window before continuing (default true)
  timeout: 30                      # seconds (default 30)
```

### `connect_window`
Attach to an already-running window (does not launch anything).
Sets the current window context for all subsequent UI actions.

```yaml
- action: connect_window
  window_class: TfJL10Edit
  window_title: "Αποθηκευμένα στοιχεία"
  timeout: 10
  on_error: continue               # 'fail' (default) or 'continue'
```

### `focus_window`
Bring a window to the foreground, then confirm it really is in front (up to 4
tries). If another application stays in front, the step fails.

```yaml
- action: focus_window
  window_class: TfMain             # optional; uses connected window if omitted
  window_title: "Epsilon"          # optional substring match
```

### `maximize_window`
Maximize a window (`ShowWindow SW_MAXIMIZE`, no keys) — `focus_window` restores
Firefox to a small size (ΑΛΛΑΓΕΣ 1.8).

```yaml
- action: maximize_window
  window_class: MozillaWindowClass   # optional; uses connected window if omitted
```

### `minimize_all`
Minimize every window on the desktop (as Win+M). Use it at the start of a run:
with a cmd window, an editor or a browser in front, the foreground check stops the
flow's keys and clicks (ΑΛΛΑΓΕΣ 1.12).

```yaml
- action: minimize_all
```

### `close_app`
Close the currently connected window by posting `WM_CLOSE` to it (a polite
close request). If a modal popup is open the application may ignore it and stay
running — use `kill_process` when the process must be gone.

```yaml
- action: close_app
```

### `kill_process`
Force-kill processes (with their child processes) by executable name, then wait
until they have exited. A leftover popup cannot block it. If the connected window
belonged to a killed process, the connection is cleared (like `close_app`).

```yaml
- action: kill_process
  names: [HyperM.exe, firefox.exe, Acrobat.exe, AcroRd32.exe]   # or a single name
  ignore_missing: true             # default true; false = fail if a name is not running
  timeout: 10                      # seconds to wait for exit (default 10)
```

In batch mode the agent already kills its reset list after every client (see
[Process reset between clients](#process-reset-between-clients)), so this is
only needed mid-flow — e.g. to restart Epsilon, or in single-client runs.

### `dialog_buttons`
The buttons of a message, left to right (ΑΛΛΑΓΕΣ_ΑΠΔ Α.7). Epsilon's
`TMessageForm` text can't be read, only its title and buttons — enough to tell
«Σφάλμα» with one «OK» from «Επιβεβαίωση» with «Ναι» / «Όχι». Epsilon builds a
message in stages (an empty window, then the buttons), so this waits until the
buttons stay the same for `settle` seconds. `[]` when no such window opens in
`timeout`.

```yaml
- action: dialog_buttons
  class_name: TMessageForm
  title: "Σφάλμα"                  # fragment; accents and case ignored
  timeout: 5                       # default 5
  settle: 0.5                      # default 0.5
  store_as: buttons                # default dialog_buttons
```

Python: `if f.dialog_buttons("TMessageForm", "Επιβεβαίωση") == ["Ναι", "Όχι"]: …`

### `dialog_text`
The text of a message (Petros 20/11/2026, αίτημα Ε): to answer «Επιβεβαίωση» only
when it asks «Υποβλήθηκε;», never an «Έξοδος από την εφαρμογή;» with the same title.

First the text controls that have their own window (a Windows MessageBox). Epsilon's
`TMessageForm` paints its text, so then **Ctrl+C**, which every Delphi message (and
every Windows MessageBox) answers by copying «title / text / buttons». Ctrl+C is sent
only when the message is the window in front (it is brought forward first), presses
no button, and the copy counts only when its title is the window's own. The
clipboard's text is put back after (a picture on it is not). `""` when the text can't
be read for sure — the flow should then halt. `copy: false` never sends Ctrl+C.

```yaml
- action: dialog_text
  class_name: TMessageForm
  title: "Επιβεβαίωση"             # fragment; accents and case ignored
  timeout: 5                       # default 5, for the window to appear
  store_as: message                # default dialog_text
```

Python:

```python
text = f.dialog_text("TMessageForm", "Επιβεβαίωση")
if "Υποβλήθηκε" not in text:
    return f.halt(f"«Επιβεβαίωση»: «{text or 'κείμενο που δεν διαβάστηκε'}»")
```

### `control_exists`
Is there a child control with this class and/or text inside a window — **hidden
ones too** (ΑΛΛΑΓΕΣ_ΑΠΔ Α.9)? E.g. the «Ελλιπή στοιχεία» tab of
`TuAPDCaseReview`, whose table can't be read but whose tabs can.

```yaml
- action: control_exists
  class_name: TRzTabSheet
  text: "Ελλιπή στοιχεία"          # exact (accents, case, & ignored); exact: false = contains
  within: { class_name: TuAPDCaseReview }   # default: the connected window
  include_hidden: true             # default true
  timeout: 0                       # default 0
  store_as: missing                # default control_exists
```

### `close_font_capture`
Close Windows' «Font Capture: Windows - Σφάλμα εφαρμογής» (ΑΛΛΑΓΕΣ_ΑΠΔ Α.10). It
belongs to Windows (csrss), so window messages don't reach it; its «OK» is
pressed through UI Automation, and it looks again in case a second one follows.
The batch runner does this **by itself** between clients; the step is for the
middle of a flow.

```yaml
- action: close_font_capture
  title: "Font Capture"            # title prefix (default)
  rounds: 5                        # default 5
```

### `screen_size`
The primary screen's `[width, height]` in real pixels. Painted toolbars and
checkboxes (no control of their own) are clicked at fixed positions, which hold for
one screen size and a maximized window only: check it first. The agent is DPI aware
(the tray and the CLI alike), so on a scaled display these are the same pixels the
clicks use.

```python
f.check(f.screen_size() == [1920, 1080], "the positions are for a 1920×1080 screen")
f.maximize_window(window_class="TfMain")
f.click(method="coordinates", x=348, y=85)   # «Εισαγωγή»: fails if Epsilon isn't the window there
```

### `set_keyboard_layout`
Switch the active keyboard layout. Affects both the current thread and the
foreground window's input thread.

```yaml
- action: set_keyboard_layout
  layout: en        # αγγλικά

- action: set_keyboard_layout
  layout: el        # ελληνικά
```

Accepted values: `el`, `greek`, `en`, `english`, `en-us`, or a raw 8-digit KLID
(e.g. `"00000408"`).

---

## UI Interaction

All UI actions that use `method: element_name` (the default) operate on the window
most recently connected via `connect_window` or `open_app`.

### Keys and clicks only go to the flow's applications

Every process the flow attaches to is remembered: `open_app` and `connect_window`
add the window's process (e.g. `HyperM.exe`), and `browser_focus_uia` /
`browser_click_uia` add the browser's (e.g. `firefox.exe`). Before a **real**
keystroke or mouse click, the engine checks where it would land:

- keys (`press_key`, `hotkey`, `type_text`, and the typing in `fill`, `select`,
  `select_by_index`, `navigate_menu`): the window in front must belong to one of
  those processes;
- mouse clicks: the window under the click point must belong to the clicked
  control's process (or, for `coordinates` / `relative_coordinates`, to one of
  the flow's processes).

If not, it waits up to 2 s for the right window to come forward, then fails the
step **without sending anything**:
`before hotkey alt+F4: the window in front is «…» (claude.exe), not hyperm.exe — nothing was sent`.
Clicks by message (`click_method: message`, `set_checkbox`) and `WM_SETTEXT` need
no check. Before the first `open_app` / `connect_window` nothing is checked.

Optional keys on any of these actions:

```yaml
- action: hotkey
  keys: ctrl+p
  process: firefox.exe        # require this process (or a list) instead of any of the flow's
- action: press_key
  key: enter
  check_foreground: false     # skip the check for this step
```

### Target resolution methods

| method | required keys | description |
|---|---|---|
| `element_name` | `target` | find child HWND by window text (default) |
| `coordinates` | `x`, `y` | absolute screen pixel coordinates |
| `relative_coordinates` | `target`, `class_name`, `offset_x`, `offset_y` | pixel offset from a resolved control's own live top-left corner; with no `target` / `class_name`, from the connected window's (e.g. a popup list) |
| `image_match` | `image` | path to a reference screenshot; uses pyautogui |

### Choosing among several matching controls (`element_name`)

Besides `target` / `class_name` / `control_index`, any action that takes an
`element_name` target accepts:

| key | picks | example |
|---|---|---|
| `control_id` | the control with that dialog id (`GetDlgCtrlID`) | in Windows dialogs (`#32770`) «OK» / «Αποθήκευση» is always id **1**, whatever the caption or language |
| `within` | only controls inside another control's rectangle: `{target, class_name}` | the `TComboBox`es of the `TGroupBox` «Περίοδος - Χρήση» |
| `prefer` | by screen position: `topmost` / `bottommost` / `leftmost` / `rightmost` / `largest` | Από = topmost combo, Έως = bottommost |

```yaml
- action: click                    # «ΟΚ» of the Windows print dialog
  target: ""
  class_name: Button
  control_id: 1
  click_method: message
```

Hidden controls never match.

### What each `index` means

The numbering is not the same everywhere. Check this table before you write an
index:

| action | key | selects | numbering |
|---|---|---|---|
| `click`, `double_click`, `fill`, `select`, `set_checkbox` | `index` / `control_index` | which matching control | **0-based** (0 = first, default 0) |
| `get_control_text` | `index` | which matching control | **1-based** (1 = first, default 1) |
| `select_by_index` | `control_index` | which matching control | **0-based** (default 0) |
| `select_by_index` | `index` | which item in the dropdown | **1-based** position; the month number for a month list |
| `browser_click_uia` | `index` | which matching web element | **0-based** (default 0) |
| `navigate_menu` | `select` | which search result | **1-based** (default 1) |
| `repeat` | `{{ i }}` | the current iteration | **0-based** |

"Which matching control" means: among all controls that match `target` /
`class_name`, in the order Windows lists them.

### `click`
Click a UI element (default left button).

```yaml
- action: click
  target: "Υποβολή Αρχείου"
  method: element_name             # default
  class_name: TRzMenuButton        # optional class filter
  index: 0                         # 0-based; which matching element (default 0 = first)
  button: left                     # left (default) / right / middle
  on_error: continue               # optional
```

```yaml
- action: click
  method: coordinates
  x: 640
  y: 480
```

```yaml
# Click a specific owner-drawn row/button inside a larger control that has no
# accessible name and no separate HWND per item (so element_name can't find
# it, and an absolute (x, y) breaks if the window moves). The offset is
# computed from the control's own rectangle at runtime via GetWindowRect.
- action: click
  method: relative_coordinates
  target: "TRzSizePanel"
  class_name: "TRzSizePanel"
  prefer: rightmost                # optional; see note below
  offset_x: 55                     # pixels right of the control's left edge
  offset_y: 131                    # pixels below the control's top edge
```

> Get exact offsets with `scripts/calibrate_click.py <ClassName> [prefer]` —
> it prints the live mouse position as an offset from the target control's
> rectangle, so you can hover the exact row and read off the numbers instead
> of estimating from a screenshot.

> If more than one control shares the same class name (and often the same
> empty window text), the default lookup returns whichever one
> `EnumChildWindows` happens to enumerate first — which follows Z-order and is
> **not** guaranteed stable across runs. Add `prefer: rightmost` / `leftmost` /
> `topmost` / `bottommost` / `largest` to disambiguate by actual screen position instead
> (e.g. a left sidebar panel vs. a right action panel that are both class
> `TRzSizePanel`).

> `relative_coordinates` only survives the *window* moving/resizing — the
> offset itself is still a measured pixel value, so it breaks if the control's
> internal layout changes (different DPI/font scaling, rows added/removed
> above the target). It's the best available fallback for owner-drawn content
> that UI Automation and Win32 HWND enumeration both can't see at all — check
> both first (`inspect_windows.py`, and a UIA dump like
> `scripts/inspect_uia_panel.py`) before reaching for it.

**How the click is performed (`click_method`)**

| `click_method` | what happens | works on | needs |
|---|---|---|---|
| `mouse` (default) | real mouse move + click at the **centre** of the control (or at x/y) | anything | the spot must be visible on screen and not covered by another window; unlocked session |
| `message` | posts `BM_CLICK` straight to the control and returns at once; no mouse, no focus | button-type controls: `TButton`, `TBitBtn`, `TcxButton`, `TCheckBox`, `TRadioButton`, `TGroupButton`, Win32 `Button` | an `element_name` target |

With `method: element_name`, the log line shows what was actually clicked:
`click: target='OK' text='OK' class='TButton' rect=(812,540,887,565) hwnd=...`.
The text, class and rect stay the same across Epsilon restarts; the hwnd changes every time.

> A click on a checkbox **toggles** it: if it was already ticked, the click unticks it.
> To set a checkbox to a specific state, use `set_checkbox`.

### `set_checkbox`
Put a checkbox in a given state. It reads the current state and clicks (by `BM_CLICK`)
only if the state is different, then confirms that the state changed. Running it
twice is safe.

```yaml
- action: set_checkbox
  target: "Εμφάνιση εργάζ. με μηδενική μισθοδοσία"
  class_name: TCheckBox
  checked: true                    # true (default) / false
  control_index: 0                 # optional, 0-based, as in click
  timeout: 3                       # seconds to wait for the new state (default 3)
```

It fails if the state has not changed after `timeout`.

### `double_click`
Double-click a UI element. Same parameters as `click`.

```yaml
- action: double_click
  target: "MyFile.xml"
```

### `fill`
Clear a text field and type a value.

```yaml
- action: fill
  target: "Όνομα αρχείου:"
  method: element_name
  value: "{{ output.folder }}\\report_{{ output.filename }}"
  clear: true                      # Ctrl+A Delete before typing (default true)
  clear_method: keyboard           # 'keyboard' (default) or 'wm_settext'
  fill_method: keyboard            # 'keyboard' (default) or 'message'
```

With `fill_method: message` the text is set with `WM_SETTEXT` — no mouse,
no keys, no foreground needed — and read back; the step fails if the field
shows something else. `value: ""` empties the field. `clear` / `clear_method`
are ignored.

> **Plain edit fields only.** On a lookup combo (`TwwDBLookupCombo`) or a
> data-bound field it changes only what the field *shows*: the record behind it
> stays the same and the read-back still passes. «Μηνιαίες Βεβαιώσεις» then
> calculated with the month that was already selected. Type into those fields
> (`fill` with `clear_method: wm_settext`). The user at «Επιλογή Χρήστη» needs
> more: typed, it shows the name but Epsilon logs in with the user selected
> before — pick it from the list with [`select_lookup`](#select_lookup).

### `select`
Select an option in a combobox or listbox by value text.

```yaml
- action: select
  target: "Όνομα:"
  method: element_name
  value: "Microsoft Print to PDF"
```

With `select_method: message` the item is picked **by name** from the combo's own
list (`CB_SETCURSEL`), the application is told as if a user had chosen it
(`CBN_SELCHANGE` → Delphi `OnChange`), and the selection is read back. No mouse,
no keys. Names match ignoring accents, case, `&` and extra spaces. If the value
isn't in the list the step fails and lists the items, unless `fallback_index`
(1-based) is given.

```yaml
- action: select
  target: "TComboBox"
  class_name: TComboBox
  within: { class_name: TGroupBox, target: "Περίοδος - Χρήση" }
  prefer: topmost                  # Από
  value: "{{ month_name }}"
  fallback_index: "{{ period.month }}"
  select_method: message
```

### `select_by_index`
Select a combobox/listbox item by its 1-based position using keyboard navigation.

```yaml
- action: select_by_index
  target: "Από"
  index: "{{ apo_index }}"        # 1-based; {{ }} variables supported
  class_name: TcxDBImageComboBox  # optional
  control_index: 0                # optional, 0-based; picks which matching control when
                                   # several share target/class_name (default 0 = first)
```

> **Note:** If the dropdown has a blank first item (position 0), use index 2 for the
> first real option, index 3 for the second, etc.

### `select_lookup`
Pick a value from a lookup combo's drop-down list — the user at Epsilon's «Επιλογή
Χρήστη». Typed into the field (`fill`), the name **shows** but Epsilon logs in with
the user selected the time before (03/11/2026: «logimate» typed, logged in as
admin). So: focus the combo → Alt+Down → the list (`popup_class`) must open →
the value is typed into the list's search → Enter → the combo must show the value,
else the step fails.

```yaml
- action: select_lookup
  value: logimate
  class_name: TwwDBLookupCombo     # default
  popup_class: TwwPopupGrid        # the list's window (default)
  timeout: 5                       # seconds for the list to open (default 5)
```

```python
f.select_lookup("logimate")
```

### `get_control_text`
Read the text content of a UI control and store it in a variable.

```yaml
- action: get_control_text
  target: "TcxDBTextEdit"
  class_name: TcxDBTextEdit
  index: 8                         # 1-based; picks the Nth matching control (default 1)
  include_hidden: true             # count hidden controls too (default false)
  store_as: fmy_amount
```

With `include_hidden: true` hidden controls count when picking the Nth, and the
log lists every matching control with its position and value, to check that
`index` points at the right field. An empty control is read through its inner
edit control (DevExpress editors).

### Where keystrokes go

`press_key`, `hotkey` and `type_text` simulate real keyboard input (like a person
typing). The keys go to **whichever control has keyboard focus in the window at the
front**, not to the connected window or to a `target`. If focus is somewhere else
(another window, a different field, the menu), the keys go there.

- Before sending keys to a specific control, use `focus_control`.
- The log shows where each `press_key` / `hotkey` went:
  `hotkey: ctrl+a -> focus text='' class='TcxGridSite' rect=(...) hwnd=...`

### `focus_control`
Give keyboard focus to a specific control, then check that it stays there. It
brings the control's window to the front, waits until the window has finished
activating, and then focuses the control.

```yaml
- action: focus_control
  target: "TcxGridSite"
  class_name: TcxGridSite
  control_index: 0                 # optional, 0-based
  timeout: 2                       # seconds (default 2)
```

It fails if focus does not stay on the control or inside it (for example, a grid's
inner editor counts). Unlike `click`, it does not move the mouse or depend on
the screen position.

### `press_key`
Send a single keypress to the focused control (see "Where keystrokes go").

```yaml
- action: press_key
  key: enter                       # enter, escape, tab, down, up, f4, etc.
```

### `hotkey`
Send a key combination to the focused control (see "Where keystrokes go").

```yaml
- action: hotkey
  keys: ctrl+p                     # parts separated by +
```

### `type_text`
Type text into whatever control currently has keyboard focus.
Use after `hotkey` / `press_key` has already positioned the cursor.

```yaml
- action: type_text
  value: "some text to type"
  clear: false                     # Ctrl+A Delete before typing (default false)
```

### `navigate_menu`
Open an Epsilon menu entry through the «αναζήτηση μενού» search box. Each stage
is checked, and the step fails with a clear message instead of typing blind:

1. waits until the main window is free of dialogs for 1s (with the `fail_on` /
   `handlers` of the last `wait_ready`, if any — see below);
2. brings the main window to the front and clicks the search box until it has
   keyboard focus (3 tries) — *«does not take keyboard focus»*;
3. empties it with `WM_SETTEXT` (Ctrl+A does not select text in this box), types
   `item` and reads it back — *«search box shows «…» instead of «…»»*;
4. Enter, then waits up to 5s for the results list (`#32768` popup) — *«no results
   list appeared»*;
5. with `rows`: the list must have exactly that many results — otherwise a
   screenshot, Escape, and the step fails without picking anything (*«listed 4
   result(s) instead of 3 — nothing was picked»*). The results' text can't be
   read, their count can. Another Epsilon user or office may list more or fewer:
   «Μισθοδοτική Κατάσταση» is the 3rd result for «logimate» (no ΝΠΔΔ), the 4th for admin;
6. waits `wait`, then Down × `select`, Enter.

The box is found by its placeholder when empty, otherwise as the topmost `TRzEdit`
that isn't one of the other search boxes.

```yaml
- action: navigate_menu
  item: "Μισθοδοτική Κατάσταση"    # text to search for
  select: 3                        # which result to open (default 1)
  rows: 3                          # the list must have exactly 3 results (optional)
  wait: 1.0                        # seconds between the list appearing and Down (default 1.0)
  timeout: 60                      # max wait for the main window to be free (default 60)
```

### `click_text_in_control`
Click text that the application **paints** inside a control — it has no window of
its own and UI Automation can't see it, e.g. «Επιλογή Όλων» on the bottom bar of
«Μισθοδοτική Κατάσταση». The engine finds the control (the anchor, or its parent),
brings the window forward, checks nothing covers it, takes a picture of it, finds
the runs of dark text pixels from left to right, and clicks the centre of run
number `run`. The log shows what it found:
`control (494, 806, 1426, 846): text at [(25, 93), (127, 213), (289, 366)]`.
If the wanted run isn't there, the step fails and the picture is saved next to the
logs.

```yaml
- action: click_text_in_control
  anchor: { target: "Εξαγωγή σε EXCEL", class_name: TRzMenuButton }
  parent: true                     # the anchor's parent = the bar (default false)
  run: 0                           # 0-based, left to right (default 0)
  max_x: 400                       # only the first 400px of the control (default: all)
  min_width: 20                    # narrower runs are borders/specks (default 20)
  gap: 12                          # widest gap inside one label, px (default 12)
  capture: window                  # screen (default) | window — see below
```

`capture: window` takes the picture of the control's **window only**
(`PrintWindow`), not of the screen, so a window over it doesn't spoil the
picture (ΑΛΛΑΓΕΣ_ΑΠΔ Α.8). The click point is still checked before clicking.

### `text_runs_in_control`
The same picture and the same keys as `click_text_in_control` (`anchor`,
`parent`, `max_x`, `min_width`, `gap`, `dark`, `margin`), but **no click**: it
stores where the text runs are, `[[x1, x2], …]` in px from the control's left
edge. Always from a picture of the window only (`PrintWindow`). Use it to check
a bar before clicking it, e.g. TfAPD's «Υπολογισμός» / «Επόμενη» must be exactly 2
(ΑΛΛΑΓΕΣ_ΑΠΔ Α.8).

```yaml
- action: text_runs_in_control
  anchor: { class_name: TPanel, control_index: 3 }
  store_as: bar_runs               # default text_runs
```

Counting them is easiest in Python:
`if len(f.text_runs_in_control({"class_name": "TPanel", "control_index": 3})) != 2: f.halt(…)`.

### `click_menu_item`
Click an item in an open native popup/context menu (e.g. after a right-click) by
its visible text, using Windows UI Automation — no screen coordinates needed, so
it keeps working regardless of window position or screen resolution. Call it
again to pick a submenu/flyout entry after the first call opens it.

```yaml
- action: click
  target: "SomeGrid"
  class_name: "TRzSizePanel"
  button: right                    # opens the context menu

- action: click_menu_item
  text: "Εξαγωγή"                  # visible text of the menu item
  exact: true                      # exact match (default true); false = contains
  timeout: 5                       # seconds to wait for the menu to appear
```

**Caveat — owner-drawn menus:** some apps (e.g. Epsilon Smart's export dropdown:
Excel/Pdf/Text) paint their menu items themselves (`MF_OWNERDRAW`) instead of
using plain menu strings. UI Automation still reports these as a real
`Menu`/`MenuItem` tree, but every item's accessible name is empty, so
`click_menu_item` can never match them by text. If it raises "not found" even
though the menu clearly opened, dump it with `scripts/inspect_uia_menu.py` to
check for blank `name=''` entries — if so, fall back to keyboard navigation by
position instead (`press_key: down` N times, then `enter`).

**Caveat — make sure you're clicking the right sub-area first:** `click`
resolves a control by class/text and then clicks its *center* — if the real
target is one owner-drawn row inside a much bigger container control (e.g. one
entry in a sidebar list, not the whole panel), clicking the container's center
can land on a completely different row and open the wrong menu (or nothing).
Use `method: relative_coordinates` to click a specific offset inside the
container instead — see `click`'s docs above.

---

## Flow Control

### `if`
Execute a branch based on a condition or window check.

**Condition-based:**
```yaml
- action: if
  condition: "{{ fmy_match }} == true"
  then:
    - action: log
      message: "Matched!"
  else:
    - action: log
      message: "No match"
```

Conditions are evaluated after resolving `{{ }}` variables. Supported operators:
`==`, `!=`, `>`, `<`, `>=`, `<=`. Both sides are compared as strings unless both
parse as numbers.

Logical operators `and` / `or` are supported. `and` binds tighter than `or`:
```yaml
condition: "{{ success }} == false and {{ attempts }} < 3"
condition: "{{ status }} == ok or {{ retries }} == 0"
```

**Window-existence check:**

Use `window_exists` instead of `condition` to branch on whether a window is open.
Add `timeout` to poll for the window up to that many seconds before deciding.

```yaml
- action: if
  window_exists:
    class_name: "#32770"           # optional
    title: "Σφάλμα"               # optional substring match
  timeout: 10                      # poll up to 10 s (default 0 = instant check)
  then:
    - action: click
      target: "ΟΚ"
      class_name: Button
  else:
    - action: log
      message: "No error dialog"
```

> Without `timeout`, the check is a one-shot snapshot. If the dialog might appear
> a fraction of a second after the previous step, always add `timeout`.

> It looks where `wait_for` looks: top-level windows **and** windows inside the
> connected window. Epsilon opens some windows (e.g. «Εταιρίες», `TfCompanies`)
> inside `TfMain`; before this, `window_exists` only saw top-level windows and
> answered False for them while `wait_for` found them (ΑΛΛΑΓΕΣ 1.3).

### `while`
Repeat steps while a condition is true.

```yaml
- action: while
  condition: "{{ counter }} < 10"
  steps:
    - action: increment
      key: counter
  max_iter: 1000                   # safety limit (default 1000)
```

### `repeat`
**Mode A — list iteration:**
```yaml
- action: repeat
  for_each: excel_data             # context variable holding a list of dicts
  as: row                          # loop variable name
  steps:
    - action: log
      message: "Row: {{ row.ΑΦΜ }}"
```

The list is typically populated by `excel_read`.

**Mode B — fixed count:**
```yaml
- action: repeat
  times: 3                         # integer or {{ }} variable
  steps:
    - action: press_key
      key: down
```

Inside `steps`, `{{ i }}` holds the 0-based iteration index (0, 1, 2, …).

### `retry`
Execute a block of steps and automatically re-run it on failure, up to `attempts` times.

```yaml
- action: retry
  attempts: 3                      # max tries (default 3)
  delay: 5                         # seconds between attempts (default 0)
  steps:
    - action: click
      target: "Υποβολή"
    - action: wait_for
      target: "Πληροφορία"
      timeout: 30
  on_exhaust: fail                 # 'fail' (default) or 'continue'
```

Each failed attempt is logged as a warning. On success the action returns immediately.
If all attempts are exhausted and `on_exhaust` is `fail`, the flow aborts.

### `try`
Execute steps and handle any exception without aborting the flow.
The exception message is available as `{{ _error }}` inside the `catch` block.

```yaml
- action: try
  steps:
    - action: connect_window
      window_class: "TfJL10"
      timeout: 10
    - action: click
      target: "Υποβολή"
  catch:
    - action: log
      message: "Failed: {{ _error }}"
    - action: hotkey
      keys: alt+F4
```

`catch` is optional — omitting it silently swallows the error and continues.

### `set_var`
Set a context variable.

```yaml
- action: set_var
  key: apo_index
  value: "{{ period.month }}"
```

### `increment`
Add a number to a context variable.

```yaml
- action: increment
  key: apo_index
  by: 1                            # amount to add (default 1)
```

### `lookup`
Look up a value in an inline table.

```yaml
- action: lookup
  table:
    "1": "Ιανουάριος"
    "2": "Φεβρουάριος"
  key: "{{ period.month }}"
  store_as: month_name
  default: ""                      # value if key not found (default "")
```

### `match_amounts`
Compare two amounts, exactly to the cent. Both are rounded to `decimals` places
(half up, decimal arithmetic) and must then be equal: `834,49` vs `834.4900000001`
matches, `834,49` vs `834.50` does **not**.

Reads either format: `"1.234,56"`, `"1,234.56"`, `"1234,56"`, `"834.49"`, `"€ 834,49"`.
Empty / `None` counts as 0. Anything else — e.g. an unresolved `"{{ fmy_amount }}"` —
fails the step instead of silently counting as 0.

```yaml
- action: match_amounts
  a: "{{ fmy_amount }}"
  b: "{{ excel_fmy }}"
  store_as: fmy_match              # stores "true" or "false" (default "amounts_match")
  decimals: 2                      # round both to this many places first (default 2)
  tolerance: 0                     # allowed difference after rounding (default 0; was 0.01)
```

---

## Wait / Timing

### `wait`
Sleep for a fixed duration. Fractional seconds supported.

```yaml
- action: wait
  seconds: 2
```

### `wait_for`
Poll until a window matching `target` becomes visible.

```yaml
- action: wait_for
  target: "Επιβεβαίωση"
  class_name: TMessageForm         # optional class filter
  timeout: 20                      # seconds (default 30)
  on_timeout: continue             # 'fail' (default) or 'continue'
```

Without `on_timeout`, a timeout **stops the flow with an error** (or goes to the
enclosing `try` / `retry`). With `on_timeout: continue` it only logs a warning
and the flow moves on. The same applies to `wait_for_gone`.

### `wait_for_gone`
Poll until a window is no longer visible.

```yaml
- action: wait_for_gone
  target: "Επεξεργασία..."
  timeout: 60
  on_timeout: continue
```

### `wait_ready`
Wait until the application's main window is **free** — enabled, i.e. not locked
by a modal dialog — for `settle` seconds in a row. Use it after login: a dialog
such as «Επιλέξτε χρήση» can appear *after* `TfMain`, and while it is open every
following key or click is lost.

Every 0.4s, looking only at windows of the main window's own process:
- a window matching `fail_on` → the step fails, naming it and its buttons;
- a window matching a handler → its `steps` run (with that window connected),
  and the settle count restarts; a handler that runs `max_runs` times (default 3)
  without the window going away fails the step;
- main window locked → every 5s the log lists the open windows and their buttons;
  the first time, a screenshot `<client>_locked_<time>.png` goes to the run's
  logs folder.

By default (`remember: true`) `navigate_menu` repeats this check (settle 1s) with
the same `fail_on` / `handlers` before every search.

```yaml
- action: wait_ready
  window_class: TfMain             # default: the connected window
  settle: 3                        # seconds it must stay free (default 3)
  timeout: 60                      # default 60
  fail_on: { title: "Σφάλμα" }     # one spec or a list: {class_name, title}
  handlers:
    - window: { class_name: TfChangeXrisi }
      steps:
        - action: fill
          target: TNumEdit
          class_name: TNumEdit
          value: "{{ period.year }}"
          clear_method: wm_settext
        - action: click
          target: "Είσοδος"
          class_name: TBitBtn
          click_method: message
```

---

## Excel

All Excel paths support `{{ }}` variable substitution.
`sheet` defaults to 1 (first sheet) and accepts either a 1-based integer or a sheet name string.

### `excel_read`
Read all rows from a sheet into a list of dicts (first row = headers).

```yaml
- action: excel_read
  path: "{{ output.folder }}\\data.xlsx"
  sheet: 1
  store_as: excel_data             # list of dicts, one per row (default "excel_data")
```

Access in `repeat`: `{{ row.HeaderName }}`

`as_lists: true` returns every row as a list of cell values instead, the first row
too, and nothing is taken as headers. Use it when the header row isn't the first,
as in Epsilon's «Μισθοδοτική Κατάσταση» (title rows above it): the flow finds the
header row itself. Dicts keyed by a short title row would lose the other columns.

```python
rows = f.excel_read(xlsx, as_lists=True, store_as="payroll_rows")
header = next(i for i, r in enumerate(rows) if "Φ.Μ.Υ" in r)
```

### `excel_read_cell`
Read a single cell value.

```yaml
- action: excel_read_cell
  path: "{{ output.folder }}\\report.xlsx"
  sheet: 1
  cell: "B5"                       # takes priority over row/column
  store_as: cell_value             # default "cell_value"
```

```yaml
- action: excel_read_cell
  path: "C:\\data.xlsx"
  row: 5
  column: "B"                      # letter or 1-based number
  store_as: my_value
```

### `excel_find_value`
Search a column for a value, return one or more cells from that row.

```yaml
- action: excel_find_value
  path: "C:\\data.xlsx"
  sheet: 1
  search_column: "A"               # column letter or 1-based number
  search_value: "12345"
  return_column: "D"               # column to read from matched row
  store_as: found_value            # default "found_value"
  not_found: ""                    # value if no match (default "")
```

To return multiple columns as a dict:

```yaml
  return_column:
    - "B"
    - "D"
```

### `excel_read_by_header`
Find columns by header name and read the **totals row** for each.
Useful when totals are at the bottom and the row count is unknown. The header
row is the first row that holds all the headers, so title rows above it are
skipped.

```yaml
- action: excel_read_by_header
  path: "{{ output.folder }}\\{{ output.filename }}.xlsx"
  headers:
    - "Φ.Μ.Υ"
    - "Ψηφ.Τέλος"
  keys:
    - excel_fmy
    - excel_digital_sig
  required: true                      # fail if a header / the totals are missing (default false)
  rows_where_filled: "Εταιρεία"       # data rows have it, the totals row doesn't
  warn_if_multiple: ["Περίοδος", "Έτος"]   # warn if the data rows span several periods
```

- **`required: true`**: a missing header or totals row fails the step. Without
  it the value is `""`, and a flow that turns `""` into 0 would compare a 0
  that was never read.
- **`rows_where_filled`**: the totals row is the last row where this column
  is blank and a header value is present. With no such row, the data rows are
  summed (with a warning). The data-row count goes to `excel_rows` (with `keys`)
  or `<store_as>_rows`. Without it, the totals row is the last non-empty row.
- **`warn_if_multiple`**: logs a warning when the data rows hold more than one
  value in these columns, e.g. a wrong «Από» month exported two periods.

When `keys` is provided each value is stored as a separate context variable.
Without `keys`, a single dict is stored under `store_as` (default `"excel_totals"`).

### `excel_write_cell`
Write a value to a cell and save the file.

```yaml
- action: excel_write_cell
  path: "C:\\report.xlsx"
  sheet: 1
  row: 5
  column: "D"                      # letter or 1-based number
  value: "{{ my_var }}"
```

### `excel_row_count`
Count the non-empty rows of a sheet and store the number, for example to check that
an export contains every employee before trusting its totals.

```yaml
- action: excel_row_count
  path: "{{ output.folder }}\\export.xlsx"
  sheet: 1                         # index (1-based) or name (default 1)
  column: "A"                      # optional: only rows with a value in this column
  start_row: 2                     # optional, 1-based (default 1); 2 = skip the header
  store_as: employees              # default: excel_row_count
```

Choose a `column` that only employee rows fill in (e.g. ΑΜΚΑ), so header, blank and
totals rows are not counted.

**Example: redo a partial export.** `retry` runs the export again if the file has
fewer rows than expected:

```yaml
- action: retry
  attempts: 3
  delay: 2
  steps:
    - action: focus_control
      target: "TcxGridSite"
      class_name: TcxGridSite
    - action: hotkey
      keys: ctrl+a
    # ... Εξαγωγή σε EXCEL, save, find_file ...
    - action: excel_row_count
      path: "{{ export_file }}"
      column: "A"
      start_row: 2
      store_as: exported_rows
    - action: assert
      condition: "{{ exported_rows }} >= {{ expected_rows }}"
      message: "Partial export: {{ exported_rows }} rows, expected {{ expected_rows }}"
```

---

## File System

### `find_file`
Search a folder for files matching a glob pattern.

```yaml
- action: find_file
  folder: "C:\\exports"
  pattern: "APD_*.xml"
  newest: true                     # pick newest match (default true)
  store_as: found_file             # default "found_file"
```

To take only the file an export just wrote — never an older one already in the
folder — note the time before the export and wait for a newer file:

```yaml
- action: set_var
  key: export_started
  value: "{{ now_epoch }}"
# ... export ...
- action: find_file
  folder: "{{ output.folder }}"
  pattern: "*.xlsx"
  newer_than: "{{ export_started }}"  # modified at/after this (1s slack)
  timeout: 30                      # look again every second, up to 30s (default 0)
  required: true                   # none found = step fails (default: found=false)
  store_as: exported_file
```

Result stored as a dict: `{ found: bool, path: str, name: str }`

```yaml
- action: if
  condition: "{{ found_file.found }} == True"
  then:
    - action: log
      message: "Found: {{ found_file.name }}"
```

### `move_file`
Move or rename a file.

```yaml
- action: move_file
  source: "{{ found_file.path }}"
  destination: "C:\\archive"      # directory or full destination path
```

### `wait_for_file`
Wait until a file exists **and has finished being written** (size > 0, unchanged
for a second). Fails after `timeout` — e.g. a print-to-PDF that never started.

```yaml
- action: wait_for_file
  path: "{{ pdf_path }}.pdf"
  timeout: 60                      # default 60
  stable: true                     # wait for the size to stop changing (default true)
  store_as: pdf                    # optional: {found, path, name, size}
```

For a file another program writes in stages, such as Epsilon's CSL01 (one read at
50,688 of 62,450 bytes broke the ΑΠΔ check):

```yaml
- action: wait_for_file
  path: "{{ export_folder }}\\CSL01"
  newer_than: "{{ export_started }}" # written by this run ({{ now_epoch }} before the export)
  last_line: "EOF"                  # its last non-empty line must be exactly this
  encoding: cp1253                  # for last_line (default utf-8)
  settle: 2                         # size unchanged for 2s (default 1)
  timeout: 60
```

### Reading files: `read_text_file`, `pdf_text`

```yaml
- action: read_text_file           # → a list of lines, no line endings
  path: "{{ export_folder }}\\CSL01"
  encoding: cp1253                  # default utf-8
  store_as: csl                     # default file_lines

- action: pdf_text                 # → the text: every page's text joined with a newline (pdfplumber)
  path: "{{ output.folder }}\\Εισφορές 09-2026.pdf"
  store_as: contributions           # default pdf_text
```

**A flow may read only:** the client's output folder (`output.folder`) and the
exports root (`MYRMEX_OUTPUT_ROOT`), the user's `Downloads`, and a file named
`CSL01` (ΑΠΔ) or `JL10` (ΦΜΥ, the «Εξαγωγή» of «Μηνιαίες Βεβαιώσεις», in
`hmisthos\`) under Epsilon's folder (`…\epsilon net\`, either Program Files). Whatever a
flow reads can reach the cloud in a message, so the agent's `.env` (its token) and
the user's browser profile must stay out of reach. A `.env` is refused anywhere.
More folders can be added in the agent's `.env` (`MYRMEX_FLOW_READ_ROOTS`,
`;`-separated), never by a flow. `last_line` of `wait_for_file` follows the same rule.

The JL10 is overwritten by every «Εξαγωγή»: make sure it is this run's before reading it.

```python
started = f.now_epoch()
# … «Εξαγωγή» → «Αποθήκευση» …
jl10 = r"C:\Program Files (x86)\epsilon net\hmisthos\JL10"
f.wait_for_file(jl10, newer_than=started, timeout=30)   # not written by this export → fails
lines = f.read_text_file(jl10, encoding="cp1253")      # fixed width: lines[1][36:45] = employer ΑΦΜ
```

### `create_folder`
Create a directory (and any missing parents). No-op if it already exists.

```yaml
- action: create_folder
  path: "{{ output.folder }}"
```

### `rename_file`
Rename a file in-place (same directory). Preserves the original extension if
`new_name` has none.

```yaml
- action: rename_file
  path: "C:\\exports\\temp.xml"
  new_name: "APD_final"
```

### `delete_file`
Delete a file, e.g. the payroll Excel and the contributions PDF once the ΑΠΔ is
submitted, so only the receipt stays in the client's folder. Only a file inside the
client's output folder (`output.folder`) or the user's Downloads: never a folder,
never a `.env`, never another client's folder or the exports root. A file that isn't
there is an error unless `missing_ok: true`.

```yaml
- action: delete_file
  path: "{{ output.folder }}\\ΜΙΣΘΟΔΟΤΙΚΗ.xlsx"
  missing_ok: true
```

---

## Browser Automation (UIA)

Interact with Firefox (or Chrome/Edge) without CDP or remote debugging.
Uses Windows UI Automation via pywinauto's UIA backend.

### `browser_focus_uia`
Bring the browser window to the foreground. Use before `hotkey` actions
(e.g. `ctrl+p`) that require the browser to be focused.

```yaml
- action: browser_focus_uia
  window_class: MozillaWindowClass  # default; use Chrome_WidgetWin_1 for Chrome/Edge
  window_title: "ΠΡΟΣΩΡΙΝΗ ΔΗΛΩΣΗ"   # optional: substring of the window title (= page title)
  focus: document                   # 'window' (default) or 'document'
  timeout: 15
```

- `focus: window` only brings the browser to the front. Keyboard focus stays where it
  was, which may be the address bar.
- `focus: document` also puts keyboard focus **inside the web page**, so `ctrl+End`,
  `Page Down` etc. act on the page. Use this instead of clicking at fixed
  screen coordinates to "click into the page"; those coordinates change with screen
  resolution, scaling and window position.
- If no window title contains `window_title`, the first browser window is used and a
  warning is logged.

To click a specific element on the page, prefer `browser_click_uia` (by visible text).
Use `click` with `method: relative_coordinates` only as a last resort, for elements
UI Automation cannot see.

### `browser_click_uia`
Click a web element by its visible text using Windows UI Automation.
Works with `<input type="button" value="...">`, links, and other accessible elements.

```yaml
- action: browser_click_uia
  text: "ΜΕΤΑΦΟΡΤΩΣΗ"
  window_class: MozillaWindowClass  # default
  control_type: Button              # UIA control type filter — omit to search all types
  exact: true                       # true = exact match, false = contains match (default false)
  index: 0                          # 0-based; picks the Nth matching element (default 0)
  timeout: 60                       # seconds (default 30)
```

Instead of (or as well as) `text`, give the element's UIA `automation_id`: steadier
and faster than a text search (ΑΛΛΑΓΕΣ 1.10). «ΜΕΤΑΦΟΡΤΩΣΗ» is `doTheSubmission`:

```yaml
- action: browser_click_uia
  automation_id: doTheSubmission
  timeout: 60
```

**Browser window classes:**

| Browser | window_class |
|---|---|
| Firefox | `MozillaWindowClass` (default) |
| Chrome / Edge | `Chrome_WidgetWin_1` |

---

### Epsilon's Firefox (port 7055): `browser_print`, `browser_page`, `browser_screenshot_pdf`

Epsilon doesn't use the normal Firefox: `EWebDriver.exe` (Selenium) starts its
own Firefox 46, which accepts WebDriver commands on `127.0.0.1:7055` (or the
next free port; the agent finds the one `firefox.exe` listens on). These
actions talk to it directly: nothing is clicked or typed on the page
(ΑΛΛΑΓΕΣ 1.7 and 1.2, found by Petros on 19/10/2026).

```yaml
# Zero ΦΜΥ has no print button, and Ctrl+P never worked from automation:
- action: browser_print            # window.print() — opens «Εκτύπωση», no keys
- action: wait_for
  target: "Εκτύπωση"
  class_name: "#32770"
  timeout: 15
# … Microsoft Print to PDF → ΟΚ → file name → Αποθήκευση, as before.

# Fallback when the print dialog doesn't open: the whole page as a picture → A4 PDF
- action: browser_screenshot_pdf
  path: "{{ pdf_path }} (φωτογραφία).pdf"

# Read the page: {{ result.url }}, {{ result.title }}, {{ result.text }}, {{ result.found }}
- action: browser_page
  contains: "Επιτυχής Συναλλαγή"   # optional → found: true / false
  store_as: result                  # default browser_page
```

All three accept `port` (default: found automatically, else 7055) and
`timeout` (seconds per call to the port, default 20). One connection is opened
per client and reused; a stale one (Firefox restarted) is replaced once.

### Any window's page through UI Automation: `uia_page`, `uia_click`

For windows that are not the Firefox of port 7055, such as **EpsilonSubmit** (a
separate program with an embedded Chrome). It is sometimes its own window and
sometimes **inside** Epsilon's «Υποβολή ΑΠΔ»: `inside_class` finds it there.

```yaml
- action: uia_page                # → [{type, name, value}, …] in UI Automation order
  window_title: "EpsilonSubmit"
  inside_class: "TfESubmitDocs"   # also look inside this window
  types: [Text, Hyperlink, DataItem]   # optional filter
  store_as: page                  # default uia_page

- action: uia_click               # the «Επανέκδοση Αποδεικτικού» of September 2026 — not another month's
  window_title: "EpsilonSubmit"
  inside_class: "TfESubmitDocs"
  name: "Επανέκδοση Αποδεικτικού" # exact, accents and case ignored (exact: false = contains)
  control_type: Hyperlink
  row_start: "^\\d{6,}$"          # a row starts at the submission number (a link of digits)
  row_start_type: Hyperlink
  row_end: "^Περίοδος Υποβολής$"  # the table ends here
  row_contains: ["(?<!\\d)0?9/2026(?!\\d)"]   # regexes; the ONE row matching all of them
  wait_enabled: 10                # wait for it to be enabled (default 10s)
```

- A page is a flat list: a **row** runs from one element matching `row_start` to
  the next. **Always give `row_end`**: without it, whatever follows the table
  counts as part of the last row, and a footer mentioning the month could make the
  wrong row match.
- If **two** rows match `row_contains`, nothing is clicked (the step fails).
- It clicks with `invoke()`, and with a real click if that fails. For a button
  that must not be pressed twice («Υποβολή & Έκδοση Αποδεικτικού»), give
  `fallback_click: false`: `invoke()` may have worked and still raised.

## Status & Diagnostics

### `log`
Write a message to the run log.

```yaml
- action: log
  message: "ΦΜΥ={{ fmy_amount }}  ΨΗΦ={{ digital_signature }}"
  level: info                      # debug | info | warning | error (default info)
```

### `screenshot`
Save a screenshot to disk.

```yaml
- action: screenshot
  path: "C:\\debug\\step_01.png"   # default: screenshots/<timestamp>.png
  store_as: screenshot_path        # optional context variable for the saved path
```

### `assert`
Assert a condition; abort the flow if it fails.

```yaml
- action: assert
  condition: "{{ fmy_match }} == true"
  message: "Amount mismatch — expected {{ excel_fmy }}, got {{ fmy_amount }}"
```

### `fail`
Unconditionally abort the flow with a message.

```yaml
- action: fail
  message: "Unexpected state — cannot continue"
```

`retryable: false` — a known reason that another try won't fix until a person
does (an amount mismatch, wrong ΕΦΚΑ / TAXISnet codes, employees outside the ΑΠΔ).
The client is reported with `retryable: false` and the cloud doesn't hand it out
again; an enclosing `retry` doesn't try it again either. Default `true`: a technical
failure (a click, a window, a timeout), where another try may work.
Python: `f.fail("ΑΣΥΜΦΩΝΙΑ CSL01 — …", retryable=False)`.

```yaml
- action: fail
  message: "ΑΣΥΜΦΩΝΙΑ CSL01 — δεν στάλθηκε"
  retryable: false
```

### `send_status`
POST a status update to the MYRMEX API (requires `MYRMEX_API_URL` and
`MYRMEX_API_KEY` in `.env`). Logs a warning instead of raising if not configured.

```yaml
- action: send_status
  status: success                  # e.g. success | error | in_progress
  message: "APD submitted for {{ period.year }}/{{ period.month }}"
```

### `mark_skipped`
Signal that this client should be counted as **skipped** rather than failed.
In batch mode the batch runner catches this and records `status: skipped` with
the reason in the summary. In single-client mode it aborts the flow cleanly.

```yaml
- action: mark_skipped
  reason: "no employees this period"
```

Place it **after all cleanup steps** in the branch that detects the skip condition,
so windows are closed before the signal is raised.

`retry` and `try` don't catch it: skipping is a decision, not an error to retry.

### `mark_done`
Finish this client **successfully** here, with a message, e.g. «0 ΕΓΓΡΑΦΕΣ» or
«ΗΔΗ ΥΠΟΒΕΒΛΗΜΕΝΟΣ». The client counts as `success`; the reason goes into
`summary.json` (`message`), `summary.txt` and the cloud (ΑΛΛΑΓΕΣ 1.5). Like
`mark_skipped`, `retry` and `try` don't catch it.

```yaml
- action: mark_done
  reason: "0 ΕΓΓΡΑΦΕΣ — χωρίς υποβολή"
```

### `halt`
**Stop everything** for a person to look (ΑΛΛΑΓΕΣ_ΑΠΔ Α.6), e.g. «Ελλιπή στοιχεία»
in the ΑΠΔ or a screen the flow doesn't recognise after a submission. Unlike `fail`:

- the client is `failed` with `retryable: false` and the message
  «ΣΤΟΠ: {reason} — τα παράθυρα έμειναν ανοιχτά»;
- the applications are **not** closed (no reset after this client), so a person
  sees the screen as it was;
- the rest of the batch does not run (`summary.json` gets `"halted": reason`);
- under the cloud the agent **pauses** (tray: «halted») and takes no jobs until
  someone resumes it from the tray. `myrmex.exe` exits with code 1.

`retry` and `try` don't catch it. Python: `return f.halt("…")`.

```yaml
- action: halt
  reason: "«Ελλιπή στοιχεία» στην ΑΠΔ — δες την οθόνη"
```

### `set_result`
Record something this client's run produced, e.g. the ΑΠΔ's submission number
(ΑΛΛΑΓΕΣ_ΑΠΔ Α.11). It goes into `summary.json` (`"result": {key: value}`) and to
the cloud with the client's result, where the job page shows it — **also when
the client fails after it** (submitted, but the receipt didn't download: the
number is what tells you it was submitted). Set it right after it is known,
before `mark_committed`.

```yaml
- action: set_result
  key: "αριθμός_υποβολής"          # up to 64 characters; setting it again replaces it
  value: "{{submission_no}}"       # text, up to 500 characters
```

At most 20 keys per client. Python: `f.set_result("αριθμός_υποβολής", number)`.

### `mark_committed`
The point of no return, e.g. right after the final «Υποβολή» to ΑΑΔΕ. If the client
fails after this step, the error says so (the reason is appended) and the client is
reported with `retryable: false`, so it is not run, and submitted, again (ΑΛΛΑΓΕΣ 2.2:
a client was submitted, its PDF failed, and the cloud scheduled it again).

```yaml
- action: mark_committed
  reason: "υποβλήθηκε στην ΑΑΔΕ — να ΜΗΝ ξανατρέξει, PDF από την ΑΑΔΕ"
```

The cloud honours `retryable: false`: the client counts as settled and is not handed
out again in that job (as after `halt` and `fail` with `retryable: false`).

---

## Batch execution

Run the same flow once per client from a CSV list in a single invocation.
Epsilon Smart is single-instance so clients run sequentially.

### CSV format

Header row is required. Column names use dot-notation matching flow variables.

```
client.afm,client.name,output.subfolder
123456789,Παπαγεωργίου ΑΕ,papageorgiou
234567890,Δημητρίου & ΣΙΑ,dimitriou
```

| Column | Description |
|---|---|
| `client.afm` | Greek tax number — used as the loop key |
| `client.name` | Display name — appears in logs and summary |
| `output.subfolder` | Subfolder under `--output-root` for this client's files |

`{{ output.folder }}` resolves to `<output-root>/<output.subfolder>/` automatically.
All other `{{ }}` variables in the flow keep working unchanged.

### CLI flags

```
myrmex.exe flows\ΦΜΥ.py --clients clients.csv --output-root C:\exports \
             --var period.year=2026 --var period.month=3
```

| Flag | Description |
|---|---|
| `--clients CSV` | Path to client list — triggers batch mode |
| `--output-root PATH` | Root folder; each client writes to `<root>/<output.subfolder>/` |
| `--start-from AFM` | Skip all clients before this AFM in the list (resume after a crash) |
| `--only AFM[,AFM,...]` | Run only these AFMs, in CSV order |
| `--skip AFM[,AFM,...]` | Run all clients except these AFMs |

`--var` flags apply globally to every client run (e.g. `period.year`, `period.month`).

### Process reset between clients

The agent force-kills a list of processes **before and after every client**,
whatever the result (success, skipped, failed), so no window from one client
survives into the next. Flows don't need `close_app` chains on every exit path.

- Default list: `HyperM.exe`, `firefox.exe`.
- The `MYRMEX_RESET_PROCESSES` env var (comma-separated) replaces the default list.
- A flow can add more names with a top-level `reset_processes` key:

```yaml
reset_processes: [Acrobat.exe, AcroRd32.exe]

steps:
  ...
```

Each kill appears in the client's log (`state reset: killed HyperM.exe (pid=...)`).
To kill processes mid-flow, use [`kill_process`](#kill_process).

### Output

Each run creates `logs/<run-id>/`:

```
logs/
  2026-03-20-1145/
    123456789.log       ← full debug log for this client
    234567890.log
    ...
    summary.json        ← machine-readable results
    summary.txt         ← human-readable table, open in any text editor
```

**Console output while running:**
```
[1/12]  123456789    Παπαγεωργίου ΑΕ            SUCCESS  4m 12s
[2/12]  234567890    Δημητρίου & ΣΙΑ            FAILED   2m 41s  (ΦΜΥ mismatch)
[3/12]  345678901    Τεχνική Εταιρεία Α          SKIPPED  1m 05s  (no employees this period)
```

**summary.txt header block:**
```
================================================================
MYRMEX Agent — Batch Run Summary
================================================================
Run ID      : 2026-03-20-1145
Flow        : basic_fmy.yaml
Period      : 3/2026
Started     : 2026-03-20 11:45:23
Finished    : 2026-03-20 12:38:41
Duration    : 53m 18s
----------------------------------------------------------------
Total       : 12
Success     : 10
Failed      : 1
Skipped     : 1
================================================================
```

### Failure isolation

A failure on client N never stops client N+1. Each client's exception is caught,
logged to its own `.log` file, and recorded as `failed` in the summary.
The batch always runs to completion.

Use `--start-from AFM` as a manual resume if the batch is interrupted mid-run.

---

## Common patterns

### Normalize an empty UI field to 0

```yaml
- action: if
  condition: "{{ fmy_amount }} == "
  then:
    - action: set_var
      key: fmy_amount
      value: "0"
```

### Compute a dropdown index that accounts for a blank first entry

```yaml
- action: set_var
  key: apo_index
  value: "{{ period.month }}"
- action: increment
  key: apo_index
  by: 1
- action: select_by_index
  target: "Από"
  index: "{{ apo_index }}"
```

### Click a button inside a Win32 dialog (no keyboard focus needed)

```yaml
- action: connect_window
  window_class: TMessageForm
  window_title: "Επιβεβαίωση"
  timeout: 10
- action: click
  target: "TButton"
  class_name: TButton
  index: 1                         # 0-based: 0 = first button (Ναι), 1 = second (Όχι)
```

> Use `connect_window` + `click` (HWND-based) instead of `focus_window` + `press_key`
> when you cannot afford to steal keyboard focus from another window.

### Click one owner-drawn row in a larger sidebar/panel control

Some panels (e.g. Epsilon Smart's right-hand action sidebar) render their whole
row list as one owner-drawn control — no accessible name, no separate HWND per
row. `element_name` can't find an individual row, and `click`'s default
center-of-control click misses it if the panel is bigger than one row. Click at
a fixed offset from the panel's own live rectangle instead of an absolute
screen pixel — it survives the window moving/resizing:

```yaml
- action: click
  method: relative_coordinates
  target: "TRzSizePanel"
  class_name: "TRzSizePanel"
  prefer: rightmost                # disambiguate if another control shares this class
  offset_x: 55                     # measured once with scripts/calibrate_click.py
  offset_y: 131

- action: wait
  seconds: 1

# If that row opens its own small dropdown (e.g. Excel/Pdf/Text), and *that*
# menu's items are also owner-drawn/unnamed, pick one by keyboard position:
- action: press_key
  key: down                        # 1st entry

- action: press_key
  key: enter
```

> Before reaching for this, confirm the target really has no accessible name at
> all: dump it with `inspect_windows.py` (separate HWND?) and a UIA name search
> like `scripts/inspect_uia_panel.py` (accessible name via UI Automation?). Only
> fall back to a measured offset when both come up empty. It's more brittle
> than name-based matching — it breaks if the panel's internal layout changes
> (rows added/removed above the target, different DPI/font scaling) — but it
> survives the one thing a hard-coded absolute `(x, y)` can't: the window being
> in a different screen position.

### Dismiss a standard Windows `#32770` dialog

```yaml
- action: wait_for
  target: "Πληροφορία"
  class_name: "#32770"
  timeout: 30
- action: connect_window
  window_class: "#32770"
  window_title: "Πληροφορία"
  timeout: 5
- action: click
  target: "Button"
  class_name: Button               # standard Windows dialogs use 'Button', not 'TButton'
```

### Retry a submission block up to 3 times

```yaml
- action: retry
  attempts: 3
  delay: 5
  steps:
    - action: click
      target: "Υποβολή"
    - action: wait_for
      target: "Πληροφορία"
      class_name: "#32770"
      timeout: 30
```

### Handle an optional error dialog before the success path

```yaml
- action: if
  window_exists:
    class_name: "#32770"
    title: "Σφάλμα"
  timeout: 10
  then:
    - action: click
      target: "ΟΚ"
      class_name: Button
    - action: fail
      message: "Unexpected error dialog"
  else:
    - action: wait_for
      target: "Πληροφορία"
      class_name: "#32770"
      timeout: 30
```

### Recover from a crash without aborting the whole flow

```yaml
- action: try
  steps:
    - action: connect_window
      window_class: "TfJL10"
      timeout: 10
    - action: hotkey
      keys: alt+F4
  catch:
    - action: log
      message: "Window already closed: {{ _error }}"
```
