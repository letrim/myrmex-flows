# Flow SDK reference (SDK 1.5.0)

Generated from `myrmex_agent/sdk/flow.py` by `scripts/gen_sdk_reference.py`. Don't edit by hand.
Design and background: [FLOW_SDK_PLAN.md](FLOW_SDK_PLAN.md). What each key does: [../YAML_REFERENCE.md](../YAML_REFERENCE.md). The names are the same.

## A flow

```python
from myrmex_agent.sdk import Param

PARAMS = {"period.year": int, "period.month": Param(int, choices=range(1, 13))}

def run(f, ctx):
    with f.phase("launch_epsilon"):
        f.open_app(EPSILON_PATH, window_class="TfAuthentication", timeout=60)
    with f.phase("export_excel"):
        if f.window_exists("TMessageForm", "Πληροφορία", timeout=10):
            return f.skip("no employees this period")
```

## Rules

- **Arguments you leave out are not sent.** The action applies its own default, the one shown in each description (`None` in a signature means "use the action's default").
- **Every call accepts `on_error="continue"`.** A failed step is then logged as a warning and the flow goes on.
- **`**TargetOpts`** are the keys that locate a control: `method, index, control_index, control_id, within, prefer, include_hidden, x, y, offset_x, offset_y, image, confidence`.
- **Results are returned *and* stored in `ctx`**, under `store_as` or the action's default name.
- **`ctx`** is the run's variables: `ctx["client.afm"]` (a KeyError when unset), `ctx.get(key, default)`, `ctx.set(key, value)`.
- **Written as plain Python instead of an action:** if, increment, lookup, repeat, try, while. `assert` is `f.check(condition, message)`.
- **Steps are numbered per phase:** `[export_excel.12] click` in the log, `export_excel[12]` in summary.json.

## Structure

- `f.phase(name)`: Group steps under a name: [name.N] in the log, name[N] in summary.json, and a place --from can resume at. Phases can nest.
- `f.skip(reason) -> NoReturn`: Count this client as skipped (mark_skipped). Use: return f.skip("…").
- `f.done(reason) -> NoReturn`: Finish this client SUCCESSFULLY here, with a message for the summary and the cloud (mark_done) — «0 εγγραφές», «ήδη υποβεβλημένος», or «ΥΠΟΒΛΗΘΗΚΕ + PDF» at the end. Use: return f.done("…").
- `f.fail(reason, *, retryable=True) -> NoReturn`: Fail this client with a message (fail). retryable=False: a known reason that another try won't fix (a mismatch, wrong codes) — the cloud doesn't run it again.
- `f.halt(reason) -> NoReturn`: Stop for a person to look (halt): a message or window nobody expected. The client fails (not retried), the applications stay OPEN, the run stops, and an agent under the cloud pauses until resumed from the tray.
- `f.committed(reason)`: The point of no return (mark_committed), e.g. right after the submission: if the client fails after this, it is reported as not to be run again.
- `f.set_result(key, value)`: Record something this client produced, e.g. the submission number: it goes into summary.json ("result") and to the cloud — also if the client fails after it.
- `f.attach(path, kind='αποδεικτικό')`: Give the cloud a file this client produced — the receipt PDF — for the accountant's «Υποβολές» (agent 0.4.2). Only a PDF up to 10 MB in the output folder or Downloads; uploaded after the client's result, retried until the cloud has it.
- `f.check(condition, message)`: Fail with `message` unless `condition` is true (YAML assert).
- `f.window_exists(class_name=None, title=None, *, timeout=None) -> bool`: Is a window with this class / title fragment open — top-level, or inside the connected window (as wait_for looks)? Waits up to `timeout` seconds (default 0). --from skipping: False.
- `f.dialog_buttons(class_name=None, title=None, *, timeout=None, settle=None, store_as=None) -> list[str]`: The buttons of a message, left to right (e.g. ["OK"] or ["Ναι", "Όχι"]), once they stay the same for `settle` s (0.5) — Epsilon builds messages in stages. Waits up to `timeout` s (5); [] when no such window opens.
- `f.dialog_text(class_name=None, title=None, *, timeout=None, copy=None, store_as=None) -> str`: The text of a message, e.g. «Υποβλήθηκε η ΑΠΔ;» vs «Έξοδος από την εφαρμογή;». Its text controls, else Ctrl+C when the message is in front (a Delphi TMessageForm paints its text): the copy counts only when its title is the window's. "" when it can't be read for sure — then halt. Waits up to `timeout` s (5); copy=False: no Ctrl+C.
- `f.save_dialog(path, *, title=None, class_name=None, timeout=None, overwrite=None, file_timeout=None, store_as=None) -> str`: Save through a Windows «Αποθήκευση» dialog: the full `path` into its FILE NAME field (found by id — not the first Edit, which in Epsilon Extra is another box), «Αποθήκευση», «αντικατάσταση;» → Ναι (overwrite, default true), the dialog must close (timeout 20) and the file must be written (file_timeout 30; 0 = don't wait). Returns the path (SDK 1.5).
- `f.control_exists(class_name=None, text=None, *, within=None, exact=None, include_hidden=None, timeout=None, store_as=None) -> bool`: Is there a child control with this class / text inside `within` ({class_name, title}; default the connected window)? Hidden controls count (include_hidden true) — e.g. a TRzTabSheet «Ελλιπή στοιχεία» in TuAPDCaseReview.
- `f.close_font_capture(*, title=None, rounds=None)`: Close Windows' «Font Capture: … Σφάλμα εφαρμογής» box(es) through UI Automation (also done between clients).
- `f.screen_size(*, store_as=None) -> list[int]`: The primary screen's [width, height] in real pixels — check it before clicking at fixed positions: f.check(f.screen_size() == [1920, 1080], "…").
- `f.retry(fn, *, attempts=None, delay=None, on_exhaust=None)`: Run fn(f), retrying on any error (retry: attempts 3, delay 0).
- `f.time_limit(seconds)`: Every step inside gives up after `seconds` (YAML max_seconds): a step that hangs stops the client instead of blocking it for 15 minutes. It is never retried — it may still be running (ΑΛΛΑΓΕΣ 1.1).
- `f.handler(window, fn, *, max_runs=None) -> dict`: A wait_ready handler: when `window` ({class_name, title}) appears, fn(f) runs with that dialog connected.
- `f.now() -> str`: The time as YAML's {{ now }}: "20261017_123845" (for file names).
- `f.now_epoch() -> int`: Seconds since 1970, as YAML's {{ now_epoch }} (find_file newer_than).
- `f.downloads_folder() -> str`: The user's Downloads folder, where the browser saves files (e.g. ΕΦΚΑ's receipt PDF), even when it was moved off %USERPROFILE%. A flow may read files there (find_file, pdf_text).

## App

- `f.open_app(path, *, window_class=None, window_title=None, wait_ready=None, timeout=None)`: Launch an executable and connect to its first window (wait_ready True, timeout 30).
- `f.close_app()`: Close the connected application.
- `f.kill_process(names, *, timeout=None, ignore_missing=None)`: Kill processes by exe name (timeout 10, ignore_missing true).
- `f.focus_window(*, window_class=None, window_title=None)`: Bring a window to the front.
- `f.connect_window(*, window_class=None, window_title=None, timeout=None)`: Make a window the target of the following UI steps (timeout 30).
- `f.maximize_window(*, window_class=None, window_title=None)`: Maximize a window (SW_MAXIMIZE, no keys); the connected one if none is given.
- `f.minimize_all()`: Minimize every window on the desktop (Win+M), e.g. at the start of a run.

## UI

- `f.click(target=None, *, class_name=None, button=None, click_method=None, hover=None, hold=None, expect_hint=None, hint_timeout=None, **TargetOpts)`: Click a control (button left, click_method mouse | message). At a position (coordinates / relative_coordinates): hover = seconds the pointer rests there first, hold = seconds the button stays down — Epsilon's painted toolbar ignores an instant click (hover 1, hold 0.15). expect_hint: the pointer rests on the point first and the tooltip that appears (Epsilon's THintWindow) must contain this text — e.g. "Εξαγωγή σε PDF" for a toolbar button with no name — else nothing is clicked and the step fails (hint_timeout s for it, default 3).
- `f.double_click(target=None, *, class_name=None, **TargetOpts)`: Double-click a control.
- `f.set_checkbox(target, *, class_name=None, checked=None, timeout=None, state_from=None, box=None, **TargetOpts)`: Set a checkbox / radio button only if needed (checked true, timeout 3), then confirm it. A box that paints its state (InfoPower TwwCheckBox, BM_GETCHECK always 0) is read from its picture — automatic for TwwCheckBox, or state_from="picture"; box="right" when the box is at the control's right end (SDK 1.5).
- `f.checkbox_state(target, *, class_name=None, state_from=None, box=None, store_as=None, **TargetOpts) -> bool`: Is the checkbox ticked? Nothing is clicked; read as set_checkbox reads it (a TwwCheckBox from its picture — the log says how dark its box is) (SDK 1.5).
- `f.fill(value, *, target=None, class_name=None, fill_method=None, clear=None, clear_method=None, **TargetOpts)`: Type a value into a field (fill_method keyboard | message, clear true, clear_method keyboard | wm_settext).
- `f.select(value, *, target=None, class_name=None, fallback_index=None, select_method=None, **TargetOpts)`: Pick a combo item by name (select_method click | message).
- `f.select_by_index(index, *, target=None, class_name=None, control_index=None, method=None)`: Pick a combo/list item by its 1-based position, with keys.
- `f.select_lookup(value, *, target=None, class_name=None, popup_class=None, timeout=None)`: Pick a value from a lookup combo's drop-down list (class_name TwwDBLookupCombo, popup_class TwwPopupGrid, timeout 5): Alt+Down → type → Enter, then the field must show it. E.g. the user at «Επιλογή Χρήστη» — typed into the field, Epsilon keeps the user selected before.
- `f.get_control_text(*, target=None, class_name=None, index=None, include_hidden=None, store_as=None, method=None) -> str`: Read a control's text (index is 1-based here, default 1). Returns it.
- `f.press_key(key)`: Press one key (enter, down, tab, …).
- `f.hotkey(keys)`: Press a key combination (ctrl+a, alt+F4, …).
- `f.type_text(value, *, clear=None)`: Type text where the keyboard focus is (clear false).
- `f.navigate_menu(item, *, wait=None, select=None, rows=None, timeout=None)`: Open an Epsilon menu item through the menu search (wait 1.0, select 1, timeout 60). rows: the results list must have exactly that many rows, else nothing is picked and the step fails.
- `f.click_menu_item(text, *, exact=None, timeout=None)`: Click an item of an open popup menu (exact true, timeout 5).
- `f.click_text_in_control(anchor, *, parent=None, run=None, min_width=None, gap=None, max_x=None, dark=None, margin=None, capture=None, expect_hint=None, hint_timeout=None)`: Click text painted inside a control, found by its pixels (run 0, gap 12, …). capture="window": from a picture of its window only (PrintWindow), even if covered. expect_hint: as click — its tooltip must contain this text (e.g. "Διαγραφή"), else no click.
- `f.text_runs_in_control(anchor, *, parent=None, min_width=None, gap=None, max_x=None, dark=None, margin=None, store_as=None) -> list[list[int]]`: Where text is painted in a control, left to right, without clicking: [[x1, x2], …] in px from its left edge — e.g. check a bar has exactly 2 labels before clicking. From a picture of its window only (PrintWindow).
- `f.focus_control(target, *, class_name=None, timeout=None, **TargetOpts)`: Give a control the keyboard focus (timeout 2).
- `f.set_keyboard_layout(layout)`: Switch the keyboard layout (en, el).

## Wait

- `f.wait(seconds)`: Pause.
- `f.wait_for(target, *, class_name=None, timeout=None, on_timeout=None)`: Wait for a window to appear (timeout 30, on_timeout fail).
- `f.wait_for_gone(target, *, class_name=None, timeout=None, on_timeout=None)`: Wait for a window to close (timeout 30, on_timeout fail).
- `f.wait_ready(*, window_class=None, window_title=None, settle=None, timeout=None, fail_on=None, handlers=None, remember=None)`: Wait until the main window stays free of dialogs (settle 3, timeout 60). handlers: [f.handler({"class_name": …}, fn), …].

## Browser

- `f.browser_focus_uia(*, window_class=None, window_title=None, focus=None, timeout=None)`: Bring Firefox forward; focus window | document (timeout 15).
- `f.browser_click_uia(text=None, *, automation_id=None, window_class=None, control_type=None, exact=None, index=None, use_invoke=None, timeout=None)`: Click a page element by its text and/or UIA AutomationId (exact false, index 0, timeout 30).
- `f.browser_print(*, port=None, timeout=None)`: Open the page's «Εκτύπωση» with window.print() through Epsilon's Firefox (port 7055): no keys, focus or keyboard layout. Then wait_for «Εκτύπωση» as usual.
- `f.browser_page(*, contains=None, store_as=None, port=None, timeout=None) -> dict`: Read the page through Epsilon's Firefox: {url, title, html, text}, plus found ("true"/"false") when `contains` is given.
- `f.browser_screenshot_pdf(path, *, store_as=None, port=None, timeout=None) -> str`: The whole page as a picture in an A4 PDF (fallback when printing fails). Returns the path.
- `f.uia_page(*, window_title=None, window_class=None, inside_class=None, types=None, timeout=None, store_as=None) -> list[dict]`: Every named element of a window's page, in UI Automation order: [{type, name, value}]. For windows that aren't the Firefox of port 7055 (EpsilonSubmit). inside_class: also find the window inside a top-level window of that class (e.g. TfESubmitDocs).
- `f.uia_click(name, *, control_type=None, exact=None, row_start=None, row_start_type=None, row_end=None, row_contains=None, index=None, window_title=None, window_class=None, inside_class=None, timeout=None, wait_enabled=None, use_invoke=None, fallback_click=None)`: Click an element of a window's page once it is enabled (invoke, else a real click). row_start (regex) + row_contains (regexes): only inside the ONE row that matches — two rows match → nothing is clicked; row_end (regex) ends the table. fallback_click=False for buttons that must not be pressed twice.

## Files

- `f.find_file(folder, pattern, *, newest=None, newer_than=None, timeout=None, required=None, store_as=None) -> dict | None`: Find a file by glob. Returns {found, path, name}.
- `f.wait_for_file(path, *, timeout=None, stable=None, settle=None, newer_than=None, last_line=None, encoding=None, store_as=None) -> dict | None`: Wait for a file to exist with a stable size (timeout 60, settle 1s); newer_than (f.now_epoch()) = written by this run; last_line = e.g. "EOF" (encoding for it). Returns {found, path, name, size} when store_as is given.
- `f.read_text_file(path, *, encoding=None, store_as=None) -> list[str]`: A text file as a list of lines (encoding utf-8; the CSL01 is cp1253). Only from the output folder, Downloads or Epsilon's CSL01.
- `f.pdf_text(path, *, store_as=None) -> str`: The text of a PDF, its pages joined with a newline (pdfplumber). Only from the output folder or Downloads.
- `f.move_file(source, destination)`
- `f.create_folder(path)`
- `f.rename_file(path, new_name)`
- `f.delete_file(path, *, missing_ok=None)`: Delete a file — only inside the client's output folder or Downloads, never a folder. missing_ok=True: a file that isn't there is fine (default: an error).

## Excel

- `f.excel_read(path, *, sheet=None, as_lists=None, store_as=None) -> list`: All rows of a sheet: dicts keyed by the first row, or with as_lists=True every row as a list (for sheets whose header row isn't the first).
- `f.excel_read_cell(path, *, cell=None, row=None, column=None, sheet=None, store_as=None) -> str`: One cell, by cell="B4" or row + column.
- `f.excel_find_value(path, search_column, search_value, return_column, *, sheet=None, not_found=None, store_as=None) -> Any`: Look a row up by one column and return another.
- `f.excel_read_by_header(path, headers, *, keys=None, rows_where_filled=None, warn_if_multiple=None, required=None, sheet=None, store_as=None) -> dict`: Totals row by header names. Returns {header: value}, or {key: value} when `keys` is given (each also stored as its own variable).
- `f.excel_write_cell(path, row, column, value, *, sheet=None)`
- `f.excel_row_count(path, *, column=None, start_row=None, sheet=None, store_as=None) -> int`

## Data

- `f.set_var(key, value)`: Set a variable, as a logged step.
- `f.match_amounts(a, b, *, decimals=None, tolerance=None, store_as=None) -> bool`: Are two amounts equal after rounding (decimals 2, tolerance 0)? Accepts "1.234,56", "834.49", "" (= 0); anything else is an error.

## Status

- `f.log(message, *, level=None)`
- `f.screenshot(path=None, *, store_as=None) -> str | None`: Save a screenshot (default: the batch logs folder). Returns its path.
- `f.send_status(status=None, message=None)`
