"""
ΦΜΥ.py — ΦΜΥ (βεβαίωση + δήλωση) for one company and period, on every office PC.

    C:\\Myrmex\\myrmex.exe flows\\ΦΜΥ.py ^
        --var period.year=2026 --var period.month=9 ^
        --var client.afm=123456789 --var client.name="ΕΤΑΙΡΙΑ ΑΕ" ^
        --var output.folder=C:\\Logimate\\exports\\123456789 ^
        [--var test=yes | --var submit=no]

10/11/2026 (Lefteris's v6–v8, merged into v9 on 23/11/2026): the PDF opens in the PC's PDF app. A PDF
program (Acrobat, Reader, SumatraPDF, Foxit, PDF-XChange) is closed; in Chrome / Edge ONLY the PDF's tab
is closed (Ctrl+W on the window that shows its file name), never the browser. The PDF also goes to the
cloud («Υποβολές», f.attach, agent 0.4.2).
v9 (23/11/2026, agent 0.4.2 / SDK 1.4): every stop says «<reason> — ΔΕΙΤΕ ΤΟ SCREENSHOT»; ΑΣΥΜΦΩΝΙΑ /
ΛΑΘΟΣ ΚΩΔΙΚΟΙ fail with retryable=False (never run again by themselves; they used to be f.done = success);
the PDF is «<ΟΝΟΜΑ> ΦΜΥ MM YYYY.pdf»; after the submission the Μισθοδοτική Excel is deleted — only the PDF
stays (Petros 20/11/2026).

Version 5 = Petros's ΦΜΥ_final.py (08/11/2026, RESOURCES\\petros\\20261108) on top
of version 4, keeping the cloud's submit:
 a. finish(): every f.done closes Acrobat, Firefox and Epsilon first (a single
    client run with --var has no batch reset and left them open); Acrobat is
    waited for up to 20 s after the PDF, and killed again if it opens late.
 b. bring_main maximizes TfMain (focus_window restores it and Epsilon shrank).
 c. No 60 s wait for «Επιβεβαίωση … Υποβλήθηκε;» after closing Firefox: it is
    already answered right after EWebDriver's «Πληροφορία» (ALATI 05/11: 61 s
    lost). The 50 s pause before «Υποβολή Αρχείου» (δήλωση) stays.

Version 4 = Petros's ΦΜΥ_2.py (04/11/2026, RESOURCES\\petros\\20261104): flow 3
with the differences found on LENOVO-PC against the reference PC
(ΦΜΥ_ΒΗΜΑΤΑ_ELEMENTS.md §11, ΔΙΑΦΟΡΕΣ_2_PC_και_ΦΜΥ_2.md), plus the cloud's submit:

 1. «Μισθοδοτική Κατάσταση» in the menu search: NOT a fixed row. The list is in
    alphabetical order of the path and the right row is always the one BEFORE
    the last («… :: Σχεδιαζόμενη Μισθοδοτική Κατάσταση»):
        reference PC, logimate: 4 rows → 3rd · LENOVO-PC, logimate: 3 rows → 2nd
        reference PC, admin:    5 rows → 4th
    navigate_menu(rows=N, select=N-1): with any other number of rows nothing
    is picked, so a wrong screen is never opened by guessing. Then the window
    must be TfMisthodotikiKatastasi, else the client stops.
 2. «Εταιρίες» (TfCompanies) is a TAB inside TfMain on both PCs, and
    connect_window only finds top-level windows («Window not found», 04/11).
    The company list is reached through TfMain; TfCompanies is only waited
    for / waited to go.
 3. «Επιβεβαίωση … να αλλάξει σε Υποβλήθηκε;» after «Υποβολή Αρχείου»: on
    LENOVO-PC it comes AFTER EWebDriver's «Πληροφορία», behind Firefox. It is
    looked for right after that «Πληροφορία» too → «Όχι».
 3β. The login check for «Σφάλμα» looks ONLY at Epsilon's message windows
    (TMessageForm / #32770): on 04/11 a Chrome tab with «Σφάλμα» in its title
    stopped the run at select_year.
 4. After Epsilon's «Εξαγωγή» (JL10): on LENOVO-PC only «Πληροφορία – ΟΚ», no
    Ναι/Όχι. Both are handled: «Ναι» if there is one, else the first button.
    Button captions are compared without «&» («Ό&χι» = «Όχι»).
 5. Before «ΜΕΤΑΦΟΡΤΩΣΗ» the flow checks that the ΑΑΔΕ page is visible to UI
    Automation (browser_focus_uia, focus=document). If not, it stops there:
    nothing has reached ΑΑΔΕ yet, the rest is done by hand.

SUBMITS TO ΑΑΔΕ whenever the screen, the Excel and the JL10 agree. test=yes —
or the cloud's submit with any value but «yes» — stops right BEFORE
«ΜΕΤΑΦΟΡΤΩΣΗ» (after the UIA check): nothing reaches ΑΑΔΕ.

Outcomes (f.done = success with a message):
    ΥΠΟΒΛΗΘΗΚΕ + PDF · ΥΠΟΒΛΗΘΗΚΕ — PDF φωτογραφία       submitted
    0 ΕΓΓΡΑΦΕΣ · ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ                      nothing to submit
    ΑΣΥΜΦΩΝΙΑ · ΑΣΥΜΦΩΝΙΑ ΑΡΧΕΙΟΥ                         not submitted: check by hand
    ΗΔΗ ΥΠΟΒΕΒΛΗΜΕΝΟΣ · ΔΟΚΙΜΗ
Anything else fails the client; after the final «Υποβολή» (f.committed) it is
reported as not to be run again.

Every step is checked (the window opened / closed, the field shows the value,
the file exists) and retried up to 3 times, 10 s apart. Clicks that can't be
undone (ΜΕΤΑΦΟΡΤΩΣΗ, Οριστικοποίηση, Υποβολή) are retried only when the button
was not found: one click that waits for the button as long as 3 attempts would,
under f.time_limit, never pressed a second time.

The cloud sends submit=yes|no with every job: yes → submit; anything else
(no, "", true…) → test, so a job that wasn't meant to submit never does.
Without `submit` (a local run) it submits. test=yes always wins.

Needs the Epsilon user «logimate» (ΧΡΗΣΤΗΣ_epsilon.py, once per office).
Epsilon, Firefox and Acrobat are killed by the batch before and after every
client (RESET_PROCESSES), and whatever opens late after a submitted client too.
"""

import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from myrmex_agent.sdk import Param       # Myrmex 0.4.0 (20/11/2026; was logimate_agent.sdk)

FLOW_VERSION = "10"       # cloud «ΦΜΥ»: 10 = Petros's local line (JL10 0-ΦΜΥ rule 11/11, problem() messages,
                          # ΑΣΥΜΦΩΝΙΑ = problem, PDF name 20/11, retryable / delete_file, Επιλογές Δήλωσης 23/11) merged
                          # with Lefteris's v9 (Hyper or Extra, JL10 beside the exe, no local paths in fail messages) —
                          # 24/11/2026; 9 (Lefteris) = Hyper or Extra, fail_shot; 8 = the receipt also to the cloud
                          # (f.attach, agent 0.4.2); 7 = v6 for the
                          # MYRMEX agent 0.4.0; 6 = PDF apps / Chrome tab, TAXISnet login, no 60″ (Petros 10-11/11);
                          # 5 = his ΦΜΥ_final.py (08/11); 4 = his ΦΜΥ_2.py (LENOVO-PC, 04/11); 3 = Petros 03-04/11
MIN_SDK_VERSION = "1.4"   # retryable / delete_file (1.3), attach (1.4)

# Killed by the batch runner before and after every client, on top of the
# agent defaults (HyperM.exe, ExtraM.exe, firefox.exe).
# PDF programs: whichever the PC opens a .pdf with is closed like Acrobat always was (10/11/2026: the default
# differs per PC — Acrobat on the reference PC, Chrome on LENOVO-PC). In a browser (PDF_BROWSER = Chrome,
# Edge, any Chromium) ONLY the PDF's own tab closes, found by its file name — never the browser.
PDF_APPS = [
    ("AcrobatSDIWindow", ["Acrobat.exe", "AcroRd32.exe"]),                  # Adobe Acrobat / Reader
    ("SUMATRA_PDF_FRAME", ["SumatraPDF.exe"]),
    ("classFoxitReader", ["FoxitPDFReader.exe", "FoxitReader.exe", "FoxitPDFEditor.exe", "FoxitPhantomPDF.exe"]),
    ("DSUI:PDFXCViewer", ["PDFXEdit.exe", "PDFXCview.exe"]),                # PDF-XChange
]
PDF_PROCESSES = [p for _, procs in PDF_APPS for p in procs]
PDF_BROWSER = "Chrome_WidgetWin_1"

EPSILON_PROCESSES = ["HyperM.exe", "ExtraM.exe"]
RESET_PROCESSES = list(PDF_PROCESSES) + EPSILON_PROCESSES   # ExtraM.exe: an agent before 0.4.3 doesn't kill it by itself

PARAMS = {
    "period.year": int,
    "period.month": Param(int, choices=range(1, 13)),
    "client.afm": Param(str, pattern=r"\d{9}"),
    "client.name": str,
    "output.folder": str,
    "test": Param(str, required=False, choices=("yes", "no")),
    "submit": Param(str, required=False),      # the cloud's job option: only «yes» submits
    # «Μηνιαίες Βεβαιώσεις Αποδοχών» → «Επιλογές Δήλωσης» (Petros 23/11/2026: buttons of the job in the cloud,
    # they only tick the box): yes = one click on it (it starts unticked: fresh Epsilon), otherwise untouched
    "fmy.amended": Param(str, required=False, choices=("yes", "no")),        # Τροποποιητική
    "fmy.total_pause": Param(str, required=False, choices=("yes", "no")),    # Ολική παύση απασχόλησης προσωπικού
    "fmy.temp_absence": Param(str, required=False, choices=("yes", "no")),   # Προσωρινή απουσία προσωπικού
    "stop.before_calc": Param(str, required=False, choices=("yes", "no")),   # test: halt before «Υπολογισμός»
}

EPSILON = r"C:\Program Files (x86)\epsilon net\hmisthos\HyperM.exe"   # default; epsilon_app() finds the real one
EPSILON_ROOTS = (r"C:\Program Files (x86)\epsilon net", r"C:\Program Files\epsilon net")
# Hyper or Extra (the lighter edition: same windows, process ExtraM.exe, folder ExtraM) — known folders first.
EPSILON_PATTERNS = ("hmisthos/HyperM.exe", "ExtraM/ExtraM.exe", "*/HyperM.exe", "*/ExtraM.exe")
# «Εξαγωγή» of Μηνιαίες (JL10, overwritten every time) is read BESIDE the exe (hmisthos\JL10 on Hyper) — run()
EPSILON_USER = "logimate"     # empty password; made by ΧΡΗΣΤΗΣ_epsilon.py
# «Μισθοδοτική Κατάσταση»: rows of the search list, most likely first. The pick
# is always rows - 1 (the row before «Σχεδιαζόμενη Μισθοδοτική Κατάσταση»).
PAYROLL_ROWS = (4, 3, 5, 6)   # 4 = reference PC, 3 = LENOVO-PC, 5 = admin
PAYROLL_WAIT = 20             # s for the window after a pick (LENOVO-PC: 6.2 s)
FIREFOX = "MozillaWindowClass"
FMY_INDEX, PSIF_INDEX = 8, 9  # TcxDBTextEdit of «Αποθηκευμένα στοιχεία», 1-based, hidden ones counted
ATTEMPTS, DELAY = 3, 10

MONTHS = {
    1: "Ιανουάριος", 2: "Φεβρουάριος", 3: "Μάρτιος", 4: "Απρίλιος",
    5: "Μάιος", 6: "Ιούνιος", 7: "Ιούλιος", 8: "Αύγουστος",
    9: "Σεπτέμβριος", 10: "Οκτώβριος", 11: "Νοέμβριος", 12: "Δεκέμβριος",
}

CENT = Decimal("0.01")
ZERO = Decimal("0.00")


# ── values ─────────────────────────────────────────────────────────────────

def norm(text) -> str:
    """Lower case, no accents, no «&» (button accelerators), single spaces:
    «Ό&χι» → «οχι», «Εταιρεία» → «εταιρεια»."""
    text = unicodedata.normalize("NFD", " ".join(str(text or "").replace("&", "").split()).lower())
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


def amount(value) -> Decimal:
    """An Excel cell or screen text as an amount with 2 decimals: 834.49, "834,49",
    "1.234,56", "1,234.56", "€ 12", None / "" (= 0)."""
    if value is None:
        return ZERO
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP)
    s = str(value).strip().replace("€", "").replace(" ", "").replace("\u00a0", "")
    if not s:
        return ZERO
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return Decimal(s).quantize(CENT, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        raise ValueError(f"δεν είναι ποσό: «{value}»") from None


def afm9(value) -> str:
    """An ΑΦΜ cell as 9 digits (Excel may give 12345678.0 for 012345678)."""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value if value is not None else "").strip().zfill(9)


def pdf_name(prefix: str, month: int, year: int, name: str) -> str:
    """«ΕΤΑΙΡΙΑ ΑΕ ΦΜΥ 09 2026» (Petros 20/11/2026: «ΟΝΟΜΑ ΠΕΛΑΤΗ ΦΜΥ ΜΗΝΑΣ»; was «ΦΜΥ 09 2026 ΕΤΑΙΡΙΑ ΑΕ»),
    «ΕΤΑΙΡΙΑ ΑΕ ΜΗΔΕΝΙΚΟ ΦΜΥ 09 2026» — no trailing dots or spaces (Windows drops them)."""
    return f"{name.rstrip('. ')} {prefix}ΦΜΥ {month:02d} {year}"


def payroll_from_excel(f, rows) -> dict:
    """The Μισθοδοτική's Excel: the header row is the first with «Φ.Μ.Υ»; an
    employee row has an «Εταιρεία», the totals row doesn't (it says «Σύνολα» a
    few columns to the left — column V on LENOVO-PC).
    Returns {fmy, psif, rows, employees: {ΑΦΜ: (Συν.Αποδ., Φ.Μ.Υ)}}."""
    header = next((i for i, r in enumerate(rows) if any(norm(c) == "φ.μ.υ" for c in r)), None)
    if header is None:
        problem(f, "δεν βρέθηκε επικεφαλίδα «Φ.Μ.Υ» στο Excel")
    cols = {norm(h): i for i, h in enumerate(rows[header])}
    missing = [h for h in ("εταιρεια", "αφμ", "φ.μ.υ", "ψηφ.τελος") if h not in cols]
    if missing:
        problem(f, f"το Excel δεν έχει στήλη {', '.join(missing)}")

    def cell(row, name):
        i = cols.get(name)
        return row[i] if i is not None and i < len(row) else None

    def filled(row, name):
        return str(cell(row, name) if cell(row, name) is not None else "").strip() != ""

    data = rows[header + 1:]
    employees_rows = [r for r in data if filled(r, "εταιρεια")]
    totals = [r for r in data if not filled(r, "εταιρεια") and filled(r, "φ.μ.υ")]
    years = sorted({str(cell(r, "ετος")) for r in employees_rows}) if "ετος" in cols else []
    if len(years) > 1:
        f.log(f"⚠ πάνω από ένα έτος στο Excel: {years}", level="WARNING")
    if not totals:
        problem(f, "δεν βρέθηκε γραμμή συνόλων στο Excel (ένα μη διαβασμένο σύνολο δεν γίνεται 0)")

    employees = {}
    for r in employees_rows:
        a = afm9(cell(r, "αφμ"))
        gross, tax = employees.get(a, (ZERO, ZERO))
        employees[a] = (gross + amount(cell(r, "συν.αποδ.")), tax + amount(cell(r, "φ.μ.υ")))
    return {"fmy": amount(cell(totals[-1], "φ.μ.υ")), "psif": amount(cell(totals[-1], "ψηφ.τελος")),
            "rows": len(employees_rows), "employees": employees}


def jl10_problems(f, lines, client_afm, employees, excel_fmy) -> list:
    """What differs between the JL10 (fixed width, cp1253, amounts in cents) and
    the Excel. Line 0 header «0JL10…» · 1 employer (ΑΦΜ at 36–45) · 2 totals ·
    3 one per employee (ΑΦΜ 1–10, gross 56–67, ΦΜΥ 98–108) — the same on both
    PCs. An unknown layout fails. Gross differences are only logged."""
    employer = [ln for ln in lines if ln[:1] == "1"]
    if not lines or not lines[0].startswith("0JL10") or len(employer) != 1:
        problem(f, f"το JL10 δεν έχει τη γνωστή μορφή ({len(lines)} γραμμές)")

    def cents(digits):
        return (Decimal(int(digits)) / 100).quantize(CENT)

    in_file = {}
    for ln in lines:
        if ln[:1] != "3":
            continue
        a, gross, tax = ln[1:10], ln[56:67], ln[98:108]
        if not (a.isdigit() and gross.isdigit() and tax.isdigit()):
            problem(f, f"το JL10 δεν έχει τη γνωστή μορφή: «{ln[:60]}»")
        g0, t0 = in_file.get(a, (ZERO, ZERO))
        in_file[a] = (g0 + cents(gross), t0 + cents(tax))
    afm = employer[0][36:45]
    total = sum((t for _, t in in_file.values()), ZERO)
    f.log(f"JL10 — ΑΦΜ={afm}  εργαζόμενοι={len(in_file)}  ΦΜΥ={total}")

    problems = []
    if afm != afm9(client_afm):
        problems.append(f"ΑΦΜ εργοδότη {afm} αντί για {client_afm}")
    if total != excel_fmy:
        problems.append(f"ΦΜΥ αρχείου {total}, Excel {excel_fmy}")
    only_file, only_excel = sorted(set(in_file) - set(employees)), sorted(set(employees) - set(in_file))
    if only_file:
        problems.append(f"εργαζόμενοι μόνο στο αρχείο: {', '.join(only_file)}")
    # Epsilon leaves out of the JL10 whoever has 0 ΦΜΥ (interns «ΠΡΑΚΤΙΚΗ ΑΣΚΗΣΗ»/ΙΕΚ with 25 days
    # and 0 €, people with 0 days): 5 false ΑΣΥΜΦΩΝΙΑ on 10–11/11/2026 (ΚΑΡΑΤΖΑ, ΜΑΝΟΥΣΟΣ,
    # ΕΥΘΥΜΙΟΥ, ΝΤΕΓΙΑΝΝΗΣ, ΦΩΤΙΑΔΗ) with the same total ΦΜΥ. Petros 11/11: «αν είναι σωστό το
    # ΦΜΥ δεν πρέπει να το λάβουμε υπόψιν» → only an Excel-only employee WITH ΦΜΥ is a problem.
    zero_tax = [a for a in only_excel if employees[a][1] == ZERO]
    if zero_tax:
        f.log(f"εργαζόμενοι μόνο στο Excel με 0 ΦΜΥ (δεν μπαίνουν στο JL10): {', '.join(zero_tax)}")
    only_excel = [a for a in only_excel if a not in zero_tax]
    if only_excel:
        problems.append(f"εργαζόμενοι μόνο στο Excel: {', '.join(only_excel)}")
    for a in sorted(set(in_file) & set(employees)):
        (g1, t1), (g2, t2) = in_file[a], employees[a]
        if t1 != t2:
            problems.append(f"{a}: ΦΜΥ αρχείου {t1}, Excel {t2}")
        if g1 != g2:
            f.log(f"⚠ {a}: ακαθάριστα αρχείου {g1}, Excel {g2}", level="WARNING")
    return problems


# ── Epsilon helpers ────────────────────────────────────────────────────────

def retry(f, fn):
    """3 attempts, 10 s apart."""
    f.retry(fn, attempts=ATTEMPTS, delay=DELAY)


def bring(f, window_class, title=None, timeout=None):
    """Connect to a TOP-LEVEL window and bring it to the front. Not for tabs
    inside TfMain (TfCompanies): connect_window doesn't find those."""
    f.connect_window(window_class=window_class, window_title=title, timeout=timeout)
    f.focus_window(window_class=window_class, window_title=title)


def bring_main(f, timeout=30):
    """TfMain in front and connected: the steps on a tab of it (Εταιρίες) look
    for their controls inside it."""
    bring(f, "TfMain", "Μισθοδοσία", timeout=timeout)
    # focus_window κάνει restore → η Epsilon βγαίνει από το maximize και μικραίνει (04/11/2026)
    f.maximize_window(window_class="TfMain", window_title="Μισθοδοσία")
    f.wait(1)


def must_close(f, window_class, title, timeout):
    """The window must close (or hide, for a tab) within `timeout` s."""
    f.wait_for_gone(title, class_name=window_class, timeout=timeout)


def problem(f, reason, known=False, shot=None):
    """Every stop of a company (Petros 20/11/2026, as ΑΠΔ.py): the reason, then «ΔΕΙΤΕ ΤΟ SCREENSHOT: <photo>».
    known=True: a reason a person must fix (ασυμφωνία, κωδικοί, …) → f.fail(retryable=False): the cloud never
    runs it again by itself (agent 0.4.1, Lefteris 22/11/2026); the message also says «ΔΕΝ ΞΑΝΑΤΡΕΧΕΙ».
    known=False: technical (a click, a window that didn't open) → may be run again."""
    # Lefteris v9 (fail_shot, 23/11/2026): NO local path in the message — it goes to the cloud and to the
    # office's email; the cloud shows the agent's own failure screenshot. The photo's path goes to the log.
    if shot is None:
        shot = f.screenshot(store_as="problem_shot", on_error="continue")
    if shot:
        f.log(f"φωτογραφία: {shot}")
    f.fail(f"{reason}{' — ΔΕΝ ΞΑΝΑΤΡΕΧΕΙ' if known else ''} — ΔΕΙΤΕ ΤΟ SCREENSHOT", retryable=not known)


def answer_year(f, year):
    """«Επιλέξτε χρήση»: the year + «Είσοδος», up to 3 times, until it closes."""
    for _ in range(3):
        if not f.window_exists("TfChangeXrisi", "Επιλέξτε χρήση"):
            return
        f.connect_window(window_class="TfChangeXrisi", window_title="Επιλέξτε χρήση", timeout=5)
        f.fill(year, target="TNumEdit", class_name="TNumEdit", fill_method="message")
        f.click("Είσοδος", class_name="TBitBtn", click_method="message")
        f.wait_for_gone("Επιλέξτε χρήση", class_name="TfChangeXrisi", timeout=5, on_timeout="continue")
    if f.window_exists("TfChangeXrisi", "Επιλέξτε χρήση"):
        problem(f, "το «Επιλέξτε χρήση» δεν κλείνει μετά από 3 προσπάθειες")


def close_wizard(f):
    """«Οδηγός Μηχανογράφησης Εταιριών» → «Έξοδος», must close."""
    f.connect_window(window_class="TfCmpIT", timeout=5)
    f.click("Έξοδος", class_name="TcxButton", click_method="message")
    must_close(f, "TfCmpIT", "Οδηγός Μηχανογράφησης Εταιριών", 10)


def pick_menu_button(f, window_class, title, target, downs, pause=0.3):
    """A TRzMenuButton's menu → item downs + 1 (Down × downs, Enter). Its menus
    are the same on both PCs: «Εξαγωγή σε EXCEL» = XLS, XLSX · «Υποβολή
    Αρχείου» = Υποβολή Αρχείου, Υποβολή Δήλωσης."""
    bring(f, window_class, title, timeout=60)
    f.click(target, class_name="TRzMenuButton")
    f.wait(1)
    for _ in range(downs):
        f.press_key("down")
        f.wait(pause)
    f.press_key("enter")


def message_answer(f, title, button, index, timeout):
    """If a TMessageForm (with this title) appears within `timeout` s: press
    `button` (by its caption, «&» ignored; else the one at `index`), then it
    must close. False when none appeared."""
    if not f.window_exists("TMessageForm", title, timeout=timeout):
        return False
    buttons = f.dialog_buttons("TMessageForm", title, timeout=5)
    wanted = [norm(x) for x in button]
    at = next((i for i, b in enumerate(buttons) if norm(b) in wanted), index)
    f.log(f"μήνυμα «{title or ''}» με κουμπιά {buttons} → #{at}")

    def answer(f):
        if not f.window_exists("TMessageForm", title):
            return
        f.connect_window(window_class="TMessageForm", window_title=title, timeout=5)
        f.click("", class_name="TButton", index=at, click_method="message")
        must_close(f, "TMessageForm", title or "", 10)
    retry(f, answer)
    return True


def dialog_ok(f, title, timeout):
    """A #32770 dialog (e.g. EWebDriver's «Πληροφορία»): its button, then it must close."""
    f.wait_for(title, class_name="#32770", timeout=timeout)

    def ok(f):
        if not f.window_exists("#32770", title):
            return
        f.connect_window(window_class="#32770", window_title=title, timeout=5)
        f.click("", class_name="Button", click_method="message")
        must_close(f, "#32770", title, 10)
    retry(f, ok)


def save_as(f, title, path):
    """A Windows save dialog: the file name by message, «Αποθήκευση» (button
    id 1), then it must close."""
    def save(f):
        if not f.window_exists("#32770", title):
            return
        f.connect_window(window_class="#32770", window_title=title, timeout=5)
        f.fill(path, target="Edit", class_name="Edit", fill_method="message")
        f.click("", class_name="Button", control_id=1, click_method="message")
        must_close(f, "#32770", title, 15)
    retry(f, save)


def close_firefox(f):
    """Every Firefox window gets WM_CLOSE (like its ×); none may stay open."""
    for _ in range(5):
        if not f.window_exists(FIREFOX):
            return
        f.connect_window(window_class=FIREFOX, timeout=5)
        f.close_app()
        f.wait_for_gone("", class_name=FIREFOX, timeout=15, on_timeout="continue")
    if f.window_exists(FIREFOX):
        problem(f, "ο Firefox δεν έκλεισε")


def pdf_app_open(f) -> bool:
    """Is a PDF program (PDF_APPS) showing a window?"""
    return any(f.window_exists(cls) for cls, _ in PDF_APPS)


def kill_pdf_apps(f, also=()):
    """The PDF programs (and `also`) are killed; never a browser. Never fails the client."""
    f.kill_process(PDF_PROCESSES + list(also), timeout=10, ignore_missing=True, on_error="continue")


def close_pdf_tab(f, name) -> bool:
    """The PDF open as the ACTIVE tab of Chrome / Edge (window title «<name> - Google Chrome»): ONLY that tab
    closes — Ctrl+W on that window, never the browser (the user's Gmail etc. stay). The window is found by
    the PDF's own file name, so no other tab matches. True when it closed. Never fails the client."""
    if not f.window_exists(PDF_BROWSER, name):
        return False
    try:
        f.connect_window(window_class=PDF_BROWSER, window_title=name, timeout=5)
        f.focus_window(window_class=PDF_BROWSER, window_title=name)
    except RuntimeError as e:                   # not in front → no keys at all (they could go elsewhere)
        f.log(f"⚠ η καρτέλα «{name}» δεν ήρθε μπροστά ({e}) — δεν κλείνει", level="WARNING")
        return False
    f.hotkey("ctrl+w")
    f.wait_for_gone(name, class_name=PDF_BROWSER, timeout=5, on_timeout="continue")
    closed = not f.window_exists(PDF_BROWSER, name)
    f.log(f"PDF στον browser: η καρτέλα «{name}» " + ("έκλεισε" if closed else "ΔΕΝ έκλεισε"),
          level="INFO" if closed else "WARNING")
    return closed


def attach_receipt(f, path, kind):
    """Το αποδεικτικό και στο cloud («Υποβολές»), για να το βρει ο λογιστής από οποιονδήποτε υπολογιστή.
    Μόνο agent 0.4.2+ (SDK 1.4): σε παλιότερο agent το flow τρέχει όπως πριν και το PDF μένει μόνο εδώ.
    on_error="continue": ο πελάτης έχει ήδη υποβληθεί — το συνημμένο δεν τον ρίχνει ποτέ.
    (Lefteris 22/11/2026: keep this call wherever a receipt is saved.)"""
    if hasattr(f, "attach"):
        f.attach(path, kind=kind, on_error="continue")


def delete_work_files(f, paths, folder=None, patterns=()):
    """Petros 20/11/2026: after the submission only the PDF stays in the company's folder — the Μισθοδοτική
    Excel is deleted (f.delete_file, agent 0.4.1: only inside the client's folder / Downloads). Only after a
    submission: on a problem it stays, to see what happened. Petros 23/11/2026: also the SAME month's Excels
    of earlier runs (`patterns` in `folder`); never another month, never a PDF. Never fails the client."""
    if not hasattr(f, "delete_file"):
        return
    for path in dict.fromkeys(p for p in paths if p):
        f.delete_file(path, missing_ok=True, on_error="continue")
        f.log(f"σβήστηκε: {path}")
    for pattern in patterns:
        seen = set()
        for _ in range(30):                     # find_file gives one file at a time (the newest)
            hit = f.find_file(folder, pattern, newest=True, required=False, store_as="leftover_file",
                              on_error="continue")
            if not hit or hit.get("found") not in (True, "true") or hit.get("path") in seen:
                break
            seen.add(hit["path"])
            f.delete_file(hit["path"], missing_ok=True, on_error="continue")
            f.log(f"σβήστηκε (από παλιότερο run): {hit['path']}")


# «Μηνιαίες Βεβαιώσεις Αποδοχών» → «Επιλογές Δήλωσης»: param → the box's caption (Petros 23/11/2026)
DECLARATION_OPTIONS = (
    ("fmy.amended", "Τροποποιητική"),
    ("fmy.total_pause", "Ολική παύση απασχόλησης προσωπικού"),
    ("fmy.temp_absence", "Προσωρινή απουσία προσωπικού"),
)
# InfoPower check boxes inside the TRzGroupBox «Επιλογές Δήλωσης» (seen on Epsilon 23/11/2026, NOT TCheckBox;
# the group also has a 4th, «Αμοιβές Επ.ΙΚΑ», which the flow never touches)
DECLARATION_BOX = "TwwCheckBox"


def declaration_options(f, ctx):
    """The three «Επιλογές Δήλωσης» boxes of «Μηνιαίες Βεβαιώσεις Αποδοχών», as the job's buttons say: yes →
    ONE click on the box, anything else → not touched.
    Their state can NOT be read (23/11/2026, ΝΙΚΟΛΑΟΥ: TwwCheckBox — set_checkbox reads 0 after a click that did
    tick it, UI Automation shows a bare Pane; request to Lefteris). Petros 23/11/2026: «επειδή κλείνουμε το
    Hyper όταν τελειώνουμε, πάντα είναι unchecked» — every client starts with a fresh Epsilon (killed before
    and after), so the boxes start unticked: one click ticks; never inside a retry (a 2nd click unticks).
    A photo of the window after the clicks is kept. A box asked for but not found → stop (nothing calculated)."""
    chosen = []
    for key, caption in DECLARATION_OPTIONS:
        if norm(ctx.get(key) or "") != "yes":
            continue
        there = f.control_exists(DECLARATION_BOX, caption, within={"class_name": "TfJL10"}, exact=False,
                                 store_as=f"jl10_option_{key}")
        if not there:
            problem(f, f"δεν βρέθηκε το «{caption}» στις «Επιλογές Δήλωσης» της «Μηνιαίες Βεβαιώσεις Αποδοχών» "
                       f"({key}=yes) — δεν υπολογίστηκε τίποτα")
        f.click(caption, class_name=DECLARATION_BOX)          # ONCE (see the docstring)
        f.wait(0.5)
        chosen.append(caption)
    f.log("Επιλογές Δήλωσης: " + (", ".join(chosen) or "καμία"))
    if chosen:
        shot = f.screenshot(store_as="declaration_options_shot")
        f.set_result("Επιλογές Δήλωσης", f"{', '.join(chosen)} · φωτογραφία {shot}")
    return chosen


def finish(f, message, acrobat_wait=0, pdf_file=None):
    """Τέλος πελάτη όπως στα παλιά scripts (ΦΜΥ.py): κλείνουν το PDF, Firefox και Epsilon και μετά
    f.done. Χωρίς batch (μία εταιρία με --var) η μηχανή δεν κάνει state reset και έμεναν ανοιχτά.
    acrobat_wait: έως τόσα ″ να ανοίξει πρώτα το PDF (το ανοίγει η «Ταυτότητα Οφειλής» και αργεί).
    pdf_file: το όνομα του PDF (π.χ. «ΕΤΑΙΡΙΑ ΦΜΥ 09 2026.pdf»): αν άνοιξε σε Chrome / Edge, κλείνει ΜΟΝΟ
    η καρτέλα του. Σε αποτυχία (f.fail) ΔΕΝ κλείνει τίποτα: η μηχανή κρατά φωτογραφία και ανοιχτά παράθυρα."""
    # on_error="continue": το κλείσιμο δεν ρίχνει ποτέ τον πελάτη — 09/11/2026 AVANTIS, DI VA, SICAP βγήκαν
    # FAILED («could not kill Acrobat.exe … PID not found»: είχε κλείσει μόνο του) ενώ είχαν υποβληθεί με PDF
    for _ in range(int(acrobat_wait)):
        if pdf_app_open(f) or (pdf_file and f.window_exists(PDF_BROWSER, pdf_file)):
            break
        f.wait(1)
    kill_pdf_apps(f, also=["firefox.exe"] + EPSILON_PROCESSES)
    # 2ο πέρασμα όπως το ΦΜΥ.py: το PDF ανοίγει και αργά (ΓΕΩΡΓΑΝΤΑΣ 05/11: 1″ μετά) → έως 10″, ξανά κλείσιμο
    for _ in range(10):
        if pdf_app_open(f):
            kill_pdf_apps(f)
            break
        if pdf_file and f.window_exists(PDF_BROWSER, pdf_file):
            break
        f.wait(1)
    if pdf_file and not close_pdf_tab(f, pdf_file) and f.window_exists(PDF_BROWSER, pdf_file):
        f.set_result("PDF στον browser", f"η καρτέλα «{pdf_file}» έμεινε ανοιχτή")
    return f.done(message)


def aade_click(f, text, control_type=None, exact=True, index=None, use_invoke=None, timeout=60):
    """A button of the ΑΑΔΕ page that can be pressed again: 3 attempts."""
    def click(f):
        with f.time_limit(timeout + 60):
            f.browser_click_uia(text, window_class=FIREFOX, control_type=control_type, exact=exact,
                                index=index, use_invoke=use_invoke, timeout=timeout)
    retry(f, click)


def aade_click_once(f, text, control_type=None, exact=True, timeout=60):
    """A button that can't be undone (ΜΕΤΑΦΟΡΤΩΣΗ, Οριστικοποίηση, Υποβολή): ONE
    click, waiting for the button as long as 3 attempts would. Any error after
    it was found stops the client instead of pressing it a second time."""
    with f.time_limit(ATTEMPTS * timeout + 60):
        f.browser_click_uia(text, window_class=FIREFOX, control_type=control_type, exact=exact,
                            timeout=ATTEMPTS * timeout)


def open_payroll_window(f) -> bool:
    """«Μισθοδοτική Κατάσταση» from the menu search, on any PC: the row before
    the last, checked by the number of rows (see PAYROLL_ROWS). A list with
    another number of rows picks nothing (navigate_menu rows=). True when
    TfMisthodotikiKatastasi is open; False = none of PAYROLL_ROWS opened it."""
    title = "Μισθοδοτική Κατάσταση"
    if f.window_exists("TfMisthodotikiKatastasi", title):
        return True
    for rows in PAYROLL_ROWS:
        f.navigate_menu(title, select=rows - 1, rows=rows, on_error="continue")
        if f.window_exists("TfMisthodotikiKatastasi", title, timeout=PAYROLL_WAIT):
            f.log(f"«{title}»: γραμμή {rows - 1} από {rows}")
            f.set_result("Μισθοδοτική στο μενού", f"{rows - 1}/{rows}")
            return True
        # Nothing picked (other number of rows) — or a pick that opened something
        # else: then TfMain is not free and the next search waits for it.
        f.log(f"«{title}»: όχι με {rows} γραμμές", level="WARNING")
    return False


def epsilon_app(f) -> str:
    """Το exe της Μισθοδοσίας Epsilon σε αυτό το PC: Hyper (HyperM.exe) ή Extra (ExtraM.exe, Lefteris v9 23/11/2026:
    ίδιες οθόνες, λιγότερες λειτουργίες). Πρώτα οι γνωστοί φάκελοι, μετά οποιοσδήποτε κάτω από «epsilon net».
    Δεν βρέθηκε (ή --from: τα βήματα πριν τη φάση δεν τρέχουν) → η προεπιλογή του Hyper."""
    for pattern in EPSILON_PATTERNS:
        for root in EPSILON_ROOTS:
            hit = f.find_file(root, pattern, store_as="epsilon_exe")
            if hit and hit.get("found") and hit.get("path"):
                f.log(f"Epsilon: {hit['path']}")
                return hit["path"]
    f.log(f"Epsilon: δεν βρέθηκε HyperM.exe / ExtraM.exe κάτω από «epsilon net» — δοκιμάζω {EPSILON}")
    return EPSILON


# ── the flow ───────────────────────────────────────────────────────────────

def run(f, ctx):
    year, month = ctx["period.year"], ctx["period.month"]
    afm, name, folder = ctx["client.afm"], ctx["client.name"], ctx["output.folder"]
    month_name = MONTHS[month]
    # submit: the cloud's choice for the job — only «yes» submits, anything else
    # is a test. No submit at all (a local run): submit. test=yes wins.
    submit = ctx.get("submit")
    test = ctx.get("test") == "yes" or (submit is not None and str(submit).strip().lower() != "yes")
    why_test = "test=yes" if ctx.get("test") == "yes" else f"submit={submit}"
    f.log(f"υποβολή στην ΑΑΔΕ: {'ΟΧΙ (' + why_test + ')' if test else 'ΝΑΙ'}")

    # ── 0. Epsilon as «logimate» ──
    with f.phase("launch_epsilon"):
        f.minimize_all()
        f.set_keyboard_layout("en")
        app = epsilon_app(f)
        # «Εξαγωγή» of Μηνιαίες, overwritten every time: JL10 beside the exe (hmisthos\JL10 on Hyper)
        jl10 = app.rsplit("\\", 1)[0] + "\\JL10"

        def open_epsilon(f):
            f.open_app(app, window_class="TfAuthentication", window_title="Επιλογή Χρήστη",
                       wait_ready=True, timeout=60)
        retry(f, open_epsilon)

        # PICKED from the list (Alt+Down → name → Enter) and read back: on both PCs
        # the field may SHOW «logimate» while the list stands on «admin».
        retry(f, lambda f: f.select_lookup(EPSILON_USER))
        retry(f, lambda f: f.fill("", target="", class_name="TEdit", fill_method="message"))

        def choose(f):
            if f.window_exists("TfAuthentication", "Επιλογή Χρήστη"):
                f.connect_window(window_class="TfAuthentication", window_title="Επιλογή Χρήστη", timeout=5)
                f.click("Επιλογή", class_name="TBitBtn", click_method="message")
            must_close(f, "TfAuthentication", "Επιλογή Χρήστη", 30)
        retry(f, choose)

    with f.phase("select_year"):
        # Ό,τι έρθει πρώτο: «Επιλέξτε χρήση» ή ο Οδηγός (= μπήκε χωρίς ερώτηση χρήσης).
        # Πριν: 60″ σταθερά για το «Επιλέξτε χρήση», που σε αυτό το PC δεν βγαίνει (09/11/2026: 61″
        # χαμένα σε κάθε πελάτη· ο Οδηγός ήταν ήδη ανοιχτός). Αν βγει αργότερα, το απαντά το wait_ready.
        # Το «Σφάλμα» ελέγχεται αμέσως μετά, όπως πριν.
        asked = False
        for _ in range(60):
            if f.window_exists("TfChangeXrisi", "Επιλέξτε χρήση", timeout=1):
                asked = True
                break
            if f.window_exists("TfCmpIT", "Οδηγός Μηχανογράφησης Εταιριών"):
                break
        # ONLY Epsilon's message: window_exists(None, "Σφάλμα") matched ANY window
        # with «Σφάλμα» in its title — on 04/11 a Gmail tab in Chrome («… Σφάλμα με
        # ΑΠΔ …») stopped the run here.
        if f.window_exists("TMessageForm", "Σφάλμα") or f.window_exists("#32770", "Σφάλμα"):
            shot = f.screenshot(store_as="login_error_shot")
            problem(f, "η Epsilon έβγαλε «Σφάλμα» στη σύνδεση (πιθανόν άλλοι χρήστες μέσα)", shot=shot)
        if asked:
            retry(f, lambda f: answer_year(f, year))

    with f.phase("main_window"):
        retry(f, lambda f: f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=60))
        if f.window_exists("TfCmpIT", "Οδηγός Μηχανογράφησης Εταιριών", timeout=60):
            retry(f, close_wizard)
        else:
            f.log("οδηγός εταιριών: δεν εμφανίστηκε")
        f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=60)
        f.wait_ready(window_class="TfMain", settle=3, timeout=60,
                     fail_on={"title": "Σφάλμα"},          # other users logged in
                     handlers=[f.handler({"class_name": "TfChangeXrisi"}, lambda f: answer_year(f, year)),
                               f.handler({"class_name": "TfCmpIT"}, close_wizard)])

    # ── 1. The company ──
    with f.phase("open_company"):
        def company_list(f):
            if not f.window_exists("TfCompanies", "Εταιρίες"):
                f.navigate_menu("Δημιουργία - διαχείριση εταιριών")
            f.wait_for("Εταιρίες", class_name="TfCompanies", timeout=30)
        retry(f, company_list)
        f.wait(3)                                   # the list fills (slower PCs)

        # «Εταιρίες» is a tab of TfMain (both PCs): its list is found through TfMain.
        # The ΑΦΜ in the first column's filter: no retry (it would be typed twice).
        # Μόνο connect: το navigate_menu την έφερε ήδη μπροστά· focus_window + maximize εδώ έκαναν
        # restore και ξανά maximize χωρίς λόγο (09/11/2026).
        f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=30)
        f.click("TcxGridSite", class_name="TcxGridSite", prefer="largest")   # the companies list
        f.hotkey("ctrl+Home")
        f.press_key("up")
        f.press_key("home")
        f.type_text(afm)
        typed = f.get_control_text(target="TcxCustomInnerTextEdit", class_name="TcxCustomInnerTextEdit",
                                   include_hidden=True, store_as="company_search")
        f.check(norm(typed) == norm(afm), f"η αναζήτηση εταιριών γράφει «{typed}» αντί για «{afm}»")
        f.log(f"ΑΦΜ στην αναζήτηση: {afm}")
        f.wait(1.5)

        def open_it(f):
            if f.window_exists("TfCompanies", "Εταιρίες"):
                # Μόνο connect: focus/restore ανάμεσα στο ΑΦΜ και στο ↓ Enter μπορεί να πάρει
                # το focus από τη λίστα και τα πλήκτρα να πάνε αλλού (09/11/2026).
                f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=5)
                f.press_key("down")             # first filtered row
                f.press_key("enter")
            must_close(f, "TfCompanies", "Εταιρίες", 30)   # still shown: the ΑΦΜ wasn't found → stop
        retry(f, open_it)

        # «Επιβεβαίωση» about ΚΑΔ as the company opens → «Όχι», photo, and on.
        if f.window_exists("TMessageForm", "Επιβεβαίωση", timeout=5):
            shot = f.screenshot(store_as="kad_shot")

            def no(f):
                if not f.window_exists("TMessageForm", "Επιβεβαίωση"):
                    return
                f.connect_window(window_class="TMessageForm", window_title="Επιβεβαίωση", timeout=5)
                f.click("", class_name="TButton", index=1, click_method="message")   # 0 = Ναι, 1 = Όχι
                must_close(f, "TMessageForm", "Επιβεβαίωση", 10)
            retry(f, no)
            f.log(f"ΚΑΔ «Επιβεβαίωση» → Όχι (φωτογραφία {shot})")
            f.set_result("ΚΑΔ", f"«Επιβεβαίωση» → Όχι · φωτογραφία {shot}")
        retry(f, lambda f: f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=30))

    # ── 2. Μισθοδοτική Κατάσταση → Excel ──
    with f.phase("export_excel"):
        f.create_folder(folder)
        filename = f"{afm}_{year}_{month}_{f.now()}"
        f.wait(1.5)

        # 3 rounds over PAYROLL_ROWS, 10 s apart; none opened it → fail (outside
        # any retry, so the message is the step's, not «retry exhausted»).
        for attempt in range(ATTEMPTS):
            if open_payroll_window(f):
                break
            if attempt < ATTEMPTS - 1:
                f.wait(DELAY)
        else:
            shot = f.screenshot(store_as="payroll_menu_shot")
            problem(f, f"δεν άνοιξε η «Μισθοδοτική Κατάσταση» από την αναζήτηση μενού (δοκιμάστηκαν "
                       f"{list(PAYROLL_ROWS)} γραμμές, {ATTEMPTS} φορές) — πρόσθεσε τον αριθμό γραμμών στο "
                       f"PAYROLL_ROWS", shot=shot)
        retry(f, lambda f: f.wait_for("Μισθοδοτική Κατάσταση", class_name="TfMisthodotikiKatastasi", timeout=30))
        f.wait(2)
        PAY = ("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση")
        bring(f, *PAY, timeout=30)

        retry(f, lambda f: f.set_checkbox("Μοναδικά φίλτρα", class_name="TRadioButton"))
        f.wait(1)
        retry(f, lambda f: f.fill(year, target="TRzSpinEdit", class_name="TRzSpinEdit", index=0,
                                  fill_method="message"))
        f.wait(1)
        # It may open on «Μία Μισθ. Περίοδος» (LENOVO-PC): «Από-Εώς» shows the two combos.
        retry(f, lambda f: f.set_checkbox("Από-Εώς", class_name="TGroupButton"))
        f.wait(1)
        # Από AND Έως = the month: the group's top and bottom VISIBLE TComboBox (the
        # group has a third one, hidden, for «Μία Μισθ. Περίοδος»), by message.
        for position in ("topmost", "bottommost"):
            retry(f, lambda f, position=position: f.select(
                month_name, target="TComboBox", class_name="TComboBox",
                within={"class_name": "TGroupBox", "target": "Περίοδος - Χρήση"},
                prefer=position, fallback_index=month, select_method="message"))
        f.wait(1)
        retry(f, lambda f: f.set_checkbox("Εμφάνιση εργάζ. με μηδενική μισθοδοσία", class_name="TCheckBox"))
        f.wait(1)

        # «Επιλογή Όλων»: painted text on the bar of «Εξαγωγή σε EXCEL», found by its
        # pixels (no fixed point: x 516–580 on LENOVO-PC, 553 on the reference).
        bar = {"target": "Εξαγωγή σε EXCEL", "class_name": "TRzMenuButton"}
        runs = []
        for attempt in range(ATTEMPTS):
            bring(f, *PAY, timeout=30)
            f.wait(0.5)
            runs = f.text_runs_in_control(bar, parent=True, max_x=400, store_as="select_all_runs")
            if runs:
                break
            if attempt < ATTEMPTS - 1:
                f.wait(DELAY)
        if not runs:
            shot = f.screenshot(store_as="no_records_shot")
            return finish(f, f"0 ΕΓΓΡΑΦΕΣ — δεν έχει εγγραφές μισθοδοσίας για {month_name} {year} "
                          f"(δεν βρέθηκε το «Επιλογή Όλων») — δεν υποβλήθηκε. Φωτογραφία: {shot}")
        f.click_text_in_control(bar, parent=True, run=0, max_x=400)
        f.wait(1)

        # «Εξαγωγή σε EXCEL» → Down Down Enter (XLSX) → «Αποθήκευση ως», or
        # «Πληροφορία» (no employees).
        export_started = f.now_epoch()
        found = {}

        def export(f):
            found.clear()
            if not (f.window_exists("TMessageForm", "Πληροφορία") or f.window_exists("#32770", "Αποθήκευση")):
                pick_menu_button(f, *PAY, "Εξαγωγή σε EXCEL", downs=2)
            for _ in range(15):
                if f.window_exists("TMessageForm", "Πληροφορία"):
                    found["kind"] = "info"
                    return
                if f.window_exists("#32770", "Αποθήκευση", timeout=1):
                    found["kind"] = "save"
                    return
            problem(f, "δεν άνοιξε το παράθυρο αποθήκευσης μετά το «Εξαγωγή σε EXCEL»")
        retry(f, export)
        if found.get("kind") == "info":
            f.connect_window(window_class="TMessageForm", window_title="Πληροφορία", timeout=5)
            f.click("", class_name="TButton", click_method="message")
            return finish(f, f"ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ — δεν έχει εργαζόμενους για {month_name} {year} — δεν υποβλήθηκε")

    with f.phase("save_export"):
        save_as(f, "Αποθήκευση", f"{folder}\\{filename}")
        f.wait(2)
        # «Πληροφορία – Η Εξαγωγή ολοκληρώθηκε με επιτυχία» → ΟΚ
        if f.window_exists("TMessageForm", timeout=10):
            f.connect_window(window_class="TMessageForm", timeout=5)
            f.click("", class_name="TButton", index=0, click_method="message")
            f.wait_for_gone("", class_name="TMessageForm", timeout=10, on_timeout="continue")

        def close_payroll(f):
            if f.window_exists(*PAY):
                f.connect_window(window_class=PAY[0], window_title=PAY[1], timeout=5)
                f.close_app()                   # WM_CLOSE, like its ×, not Alt+F4
            must_close(f, PAY[0], PAY[1], 30)
        retry(f, close_payroll)

        def find_export(f):
            exported = f.find_file(folder, "*.xlsx", newest=True, newer_than=export_started,
                                   timeout=5, required=True, store_as="exported_file")
            if exported["name"] != f"{filename}.xlsx":
                f.rename_file(exported["path"], f"{filename}.xlsx")
        f.retry(find_export, attempts=6, delay=5)
        xlsx = f"{folder}\\{filename}.xlsx"
        f.log(f"Excel: {xlsx}")

    # ── 3. Μηνιαίες Βεβαιώσεις Αποδοχών ──
    with f.phase("calculate_fmy"):
        retry(f, lambda f: f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=60))
        f.wait(1)

        def open_choose(f):
            if f.window_exists("TfJL10", "Μηνιαίες Βεβαιώσεις"):
                return
            if not f.window_exists("TfChooseCompanies", "Επιλογή Εταιριών"):
                f.navigate_menu("Μηνιαίες Βεβαιώσεις Αποδοχών", rows=1)   # 1 row on both PCs
            f.wait_for("Επιλογή Εταιριών", class_name="TfChooseCompanies", timeout=30)
        retry(f, open_choose)
        f.wait(1)

        def choose_company(f):
            if f.window_exists("TfChooseCompanies", "Επιλογή Εταιριών"):
                f.connect_window(window_class="TfChooseCompanies", window_title="Επιλογή Εταιριών", timeout=5)
                f.click("Επιλογή", class_name="TBitBtn", click_method="message")
            f.wait_for("Μηνιαίες Βεβαιώσεις", class_name="TfJL10", timeout=60)
        retry(f, choose_company)
        f.wait(1)
        JL = ("TfJL10", "Μηνιαίες Βεβαιώσεις")
        bring(f, *JL, timeout=60)

        retry(f, lambda f: f.fill(f"{month} {year}", target="TEdit", class_name="TEdit", fill_method="message"))
        f.wait(1)
        retry(f, lambda f: f.fill(year, target="TRzSpinEdit", class_name="TRzSpinEdit", fill_method="message"))
        f.wait(1)
        # The month is TYPED: in a lookup combo WM_SETTEXT changes only what it shows.
        retry(f, lambda f: f.fill(month_name, target="TwwDBLookupCombo", class_name="TwwDBLookupCombo",
                                  clear_method="wm_settext"))
        f.wait(1)
        chosen = declaration_options(f, ctx)
        # stop.before_calc=yes: stop HERE, before «Υπολογισμός», everything open, to look at the boxes on a
        # real Epsilon (Petros 23/11/2026: test of the «Επιλογές Δήλωσης» buttons) — nothing calculated or sent
        if norm(ctx.get("stop.before_calc") or "") == "yes":
            shot = f.screenshot(store_as="before_calc_shot")
            return f.halt(f"ΣΤΟΠ ΠΡΙΝ ΤΟΝ «Υπολογισμό» (stop.before_calc=yes) — Επιλογές Δήλωσης: "
                          f"{', '.join(chosen) or 'καμία'} — τίποτα δεν υπολογίστηκε ούτε στάλθηκε, όλα ανοιχτά. "
                          f"Φωτογραφία {shot}")

        # «Υπολογισμός» → «Αποθηκευμένα στοιχεία» (LENOVO-PC 2.5 s, reference 9.7 s).
        f.click("Υπολογισμός", class_name="TcxButton", click_method="message")
        if not f.window_exists("TfJL10Edit", "Αποθηκευμένα στοιχεία", timeout=120):
            f.log("«Υπολογισμός»: δεν άνοιξαν τα «Αποθηκευμένα στοιχεία» σε 120s — ξανά", level="WARNING")
            with f.time_limit(60):
                f.uia_click("Υπολογισμός", window_class="TfJL10", fallback_click=False, timeout=10)
            f.wait_for("Αποθηκευμένα στοιχεία", class_name="TfJL10Edit", timeout=120)
        f.log("«Υπολογισμός»: άνοιξαν τα «Αποθηκευμένα στοιχεία»")
        f.wait(1)

    with f.phase("read_results"):
        SAVED = ("TfJL10Edit", "Αποθηκευμένα στοιχεία")
        shown = {}

        def read_screen(f):
            f.connect_window(window_class=SAVED[0], window_title=SAVED[1], timeout=30)
            shown["fmy"] = amount(f.get_control_text(target="TcxDBTextEdit", class_name="TcxDBTextEdit",
                                                     index=FMY_INDEX, include_hidden=True, store_as="fmy_amount"))
            shown["psif"] = amount(f.get_control_text(target="TcxDBTextEdit", class_name="TcxDBTextEdit",
                                                      index=PSIF_INDEX, include_hidden=True,
                                                      store_as="digital_signature"))
        retry(f, read_screen)
        fmy_ui, psif_ui = shown["fmy"], shown["psif"]
        f.log(f"Τιμές οθόνης — AFM={afm}  ΦΜΥ={fmy_ui}  ΨΗΦ={psif_ui}")
        f.set_result("ΦΜΥ", str(fmy_ui))
        f.set_result("Ψηφ.Τέλος", str(psif_ui))

    # ── 4. Excel + comparison ──
    with f.phase("compare"):
        payroll = payroll_from_excel(f, f.excel_read(xlsx, as_lists=True, store_as="payroll_rows"))
        fmy_xlsx, psif_xlsx = payroll["fmy"], payroll["psif"]
        f.log(f"Τιμές Excel — ΦΜΥ={fmy_xlsx}  ΨΗΦ={psif_xlsx}  ({payroll['rows']} γραμμές)")
        f.log(f"Match — ΦΜΥ={str(fmy_ui == fmy_xlsx).lower()}  ΨΗΦ={str(psif_ui == psif_xlsx).lower()}")
        # Petros 20/11/2026: an ΑΣΥΜΦΩΝΙΑ is a PROBLEM (it was finish → f.done = success in the cloud),
        # with a photo, never run again by itself
        if fmy_ui != fmy_xlsx:
            problem(f, f"ΑΣΥΜΦΩΝΙΑ Φ.Μ.Υ — οθόνη {fmy_ui}, Excel {fmy_xlsx} "
                       f"(Ψηφ. οθόνη {psif_ui}, Excel {psif_xlsx}) — δεν υποβλήθηκε", known=True)
        if psif_ui != psif_xlsx:
            problem(f, f"ΑΣΥΜΦΩΝΙΑ Ψηφ.Τέλους — οθόνη {psif_ui}, Excel {psif_xlsx} — δεν υποβλήθηκε", known=True)

    # ── 5. Submission ──
    with f.phase("submit"):
        with f.phase("export_file"):
            # Epsilon: «Εξαγωγή» → «Αποθήκευση» → [Ναι] → ΟΚ
            export_started = f.now_epoch()

            def export_file(f):
                if not f.window_exists("#32770", "Αποθήκευση"):
                    f.connect_window(window_class=SAVED[0], window_title=SAVED[1], timeout=30)
                    f.click("Εξαγωγή", class_name="TcxButton", click_method="message")
                f.wait_for("Αποθήκευση", class_name="#32770", timeout=20)
            retry(f, export_file)

            def save_file(f):
                if f.window_exists("#32770", "Αποθήκευση"):
                    f.connect_window(window_class="#32770", window_title="Αποθήκευση", timeout=5)
                    f.click("", class_name="Button", control_id=1, click_method="message")   # id 1
                must_close(f, "#32770", "Αποθήκευση", 10)
            retry(f, save_file)
            # Reference PC: a question about the file (→ Ναι), then ΟΚ. LENOVO-PC:
            # only «Πληροφορία – Η εξαγωγή ολοκληρώθηκε επιτυχώς» (→ its ΟΚ).
            message_answer(f, None, ["Ναι"], 0, 10)
            message_answer(f, None, ["ΟΚ", "OK"], 0, 10)

        with f.phase("check_jl10"):
            f.wait_for_file(jl10, newer_than=export_started, timeout=30)
            lines = f.read_text_file(jl10, encoding="cp1253", store_as="jl10_lines")
            problems = jl10_problems(f, lines, afm, payroll["employees"], fmy_xlsx)
            if problems:
                problem(f, "ΑΣΥΜΦΩΝΙΑ ΑΡΧΕΙΟΥ — ΑΣΥΜΦΩΝΙΑ JL10 — " + "; ".join(problems) + " — δεν υποβλήθηκε",
                        known=True)
            f.log("JL10 = Excel (ΑΦΜ εργοδότη, εργαζόμενοι, ΦΜΥ ανά εργαζόμενο)")

        with f.phase("send_certificate"):
            # «Υποβολή Αρχείου» (βεβαίωση) → [Επιβεβαίωση → Όχι] → EWebDriver's «Πληροφορία»
            # Λάθος κωδικοί TAXISnet της εταιρίας στην Epsilon: η «Πληροφορία» δεν έρχεται ποτέ και ο Firefox
            # γράφει «Καθορίστηκε λανθασμένο όνομα χρήστη ή κωδικός πρόσβασης» (ΓΕΩΡΓΟΥΛΗΣ 09/11/2026: 2×70″
            # αναμονή και ξανά «Υποβολή Αρχείου»). Πέτρος 09/11: σφάλμα αμέσως, χωρίς 2ο «Υποβολή Αρχείου».
            bad_login = {}

            def send_cert(f):
                if bad_login:
                    return
                if not (f.window_exists("#32770", "Πληροφορία") or f.window_exists("TMessageForm", "Επιβεβαίωση")):
                    pick_menu_button(f, *SAVED, "Υποβολή Αρχείου", downs=1)
                message_answer(f, "Επιβεβαίωση", ["Όχι"], 1, 10)
                for _ in range(10):                     # ~70″, as before
                    if f.window_exists("#32770", "Πληροφορία", timeout=5):
                        break
                    page = f.browser_page(contains="λανθασμένο όνομα χρήστη", timeout=3, on_error="continue",
                                          store_as="login_page") or {}
                    if str(page.get("found", "")).lower() == "true":
                        bad_login["shot"] = f.screenshot(store_as="bad_login_shot")
                        return                          # out of f.retry: no 2nd «Υποβολή Αρχείου»
                f.wait_for("Πληροφορία", class_name="#32770", timeout=1)
            f.retry(send_cert, attempts=2, delay=DELAY)
            if bad_login:
                problem(f, "ΛΑΘΟΣ ΚΩΔΙΚΟΙ TAXISnet — ο Firefox γράφει «Καθορίστηκε λανθασμένο όνομα χρήστη ή κωδικός "
                           "πρόσβασης»: διόρθωση των κωδικών TAXISnet της εταιρίας στην Epsilon. Δεν υποβλήθηκε τίποτα",
                        known=True, shot=bad_login["shot"])
            dialog_ok(f, "Πληροφορία", 10)
            # LENOVO-PC: the «Επιβεβαίωση … Υποβλήθηκε;» comes HERE, behind Firefox.
            message_answer(f, "Επιβεβαίωση", ["Όχι"], 1, 5)
            f.wait(5)

            page = f.browser_page(on_error="continue") or {}
            if page.get("url"):
                f.log(f"σελίδα ΑΑΔΕ: «{page.get('title', '')}» {page['url'][:90]}")

        with f.phase("check_uia"):
            # The ΑΑΔΕ buttons are pressed through UI Automation. LENOVO-PC (04/11):
            # a first probe saw no UIA elements (a fault of the probe, not of the PC).
            # Nothing has reached ΑΑΔΕ yet: without a UIA page, stop here rather than
            # fail in the middle.
            uia_ok = True
            try:
                f.browser_focus_uia(window_class=FIREFOX, focus="document", timeout=30)
            except RuntimeError as e:
                uia_ok = False
                f.log(f"UIA: {e}", level="WARNING")
            if not uia_ok:
                shot = f.screenshot(store_as="no_uia_shot")
                problem(f, "ΣΤΟΠ ΠΡΙΝ ΤΗ ΜΕΤΑΦΟΡΤΩΣΗ — ο Firefox της Epsilon δεν δείχνει τη σελίδα στο UI "
                           "Automation σε αυτό το PC, οπότε τα κουμπιά της ΑΑΔΕ δεν βρίσκονται. Τίποτα δεν "
                           "στάλθηκε στην ΑΑΔΕ: συνέχισε με το χέρι από τη σελίδα που είναι ανοιχτή "
                           "(ΜΕΤΑΦΟΡΤΩΣΗ → ΛΙΣΤΑ ΥΠΟΒΟΛΩΝ → Οριστικοποίηση …)", shot=shot)
            if test:
                shot = f.screenshot(store_as="test_shot")
                return finish(f, f"ΔΟΚΙΜΗ — {why_test}: σταμάτησε ΠΡΙΝ τη ΜΕΤΑΦΟΡΤΩΣΗ (UIA εντάξει), τίποτα δεν "
                              f"στάλθηκε στην ΑΑΔΕ. Φωτογραφία: {shot}")

        with f.phase("aade_certificate"):
            aade_click_once(f, "ΜΕΤΑΦΟΡΤΩΣΗ", "Button", timeout=60)
            f.wait(5)
            aade_click(f, "ΛΙΣΤΑ ΥΠΟΒΟΛΩΝ", "Button", exact=False, timeout=60)
            f.wait(5)
            aade_click_once(f, "Οριστικοποίηση", timeout=60)
            f.wait(5)
            aade_click_once(f, "Οριστικοποίηση Υποβολής", timeout=60)
            f.set_result("βεβαίωση", "οριστικοποιήθηκε")
            f.wait(1)
            with f.time_limit(120):
                f.browser_click_uia("OK", window_class=FIREFOX, control_type="Button", timeout=60,
                                    on_error="continue")
            f.wait(3)
            close_firefox(f)
            # ΑΝΕΝΕΡΓΟ προσωρινά (05/11/2026): το «Επιβεβαίωση … Υποβλήθηκε;» απαντιέται ήδη «Όχι» αμέσως μετά
            # το «Πληροφορία» του EWebDriver (send_certificate)· εδώ περίμενε άδικα 60″ (ALATI 05/11: 61″).
            # if f.window_exists("TMessageForm", "Επιβεβαίωση", timeout=60):
            #     message_answer(f, "Επιβεβαίωση", ["Όχι"], 1, 1)
            f.wait(50)

        with f.phase("send_return"):
            # «Υποβολή Αρχείου» (δήλωση): no retry. «Σφάλμα» = already submitted.
            pick_menu_button(f, *SAVED, "Υποβολή Αρχείου", downs=2)
            f.wait(30)
            if f.window_exists("#32770", "Σφάλμα", timeout=10):
                dialog_ok(f, "Σφάλμα", 5)
                return finish(f, "ΗΔΗ ΥΠΟΒΕΒΛΗΜΕΝΟΣ — έχει καταχωρηθεί ξανά (η βεβαίωση οριστικοποιήθηκε ξανά)")
            dialog_ok(f, "Πληροφορία", 60)   # LENOVO-PC: after 21 s
            f.wait(2)

        with f.phase("aade_submit"):
            # ΠΡΟΣΩΡΙΝΗ ΔΗΛΩΣΗ: focus in the page + Ctrl+End + the bottom «Υποβολή»
            # (an <input type="submit">: UIA Edit; invoke only focuses it → a real click).
            def temporary(f):
                f.browser_focus_uia(window_class=FIREFOX, window_title="ΠΡΟΣΩΡΙΝΗ ΔΗΛΩΣΗ", focus="document",
                                    timeout=60)
                f.wait(1)
                f.hotkey("ctrl+End")
                f.wait(5)                       # the page loads this row asynchronously
                with f.time_limit(120):
                    f.browser_click_uia("Υποβολή", window_class=FIREFOX, control_type="Edit", exact=True,
                                        index=1, use_invoke=False, timeout=60)
            retry(f, temporary)
            f.wait(2)
            aade_click_once(f, "Υποβολή", "Button", timeout=30)       # the final one: can't be undone
            f.committed("ΥΠΟΒΛΗΘΗΚΕ στην ΑΑΔΕ — αν λείπει το PDF, κατέβασμα από ΑΑΔΕ· να ΜΗΝ ξανατρέξει")
            f.set_result("Υποβολή", f.now())
            f.wait(5)

            page = f.browser_page(contains="Επιτυχής", on_error="continue") or {}
            if "Επιτυχής" in page.get("text", ""):
                m = re.search(r"Αριθμό Καταχώρησης\s*myAADE\s*(\d+)", page["text"])
                reg_no = m.group(1) if m else ""
                if reg_no:
                    f.set_result("myAADE", reg_no)
                f.log("ΑΑΔΕ: «Επιτυχής Συναλλαγή»" + (f" — αρ. καταχώρησης myAADE {reg_no}" if reg_no else ""))
            else:
                f.log("⚠ δεν διαβάστηκε «Επιτυχής Συναλλαγή» στη σελίδα — έλεγξε στην ΑΑΔΕ", level="WARNING")

        with f.phase("save_pdf"):
            # Zero (ΦΜΥ 0 AND Ψηφ.Τέλος 0) → window.print() through port 7055;
            # otherwise «Ταυτότητα Οφειλής» → «Εκτύπωση».
            zero = fmy_ui == 0 and psif_ui == 0
            path = f"{folder}\\{pdf_name('ΜΗΔΕΝΙΚΟ ' if zero else '', month, year, name)}"
            if zero:
                opened = False
                for attempt in range(ATTEMPTS):
                    if not f.window_exists("#32770", "Εκτύπωση"):
                        f.browser_print(on_error="continue")
                    if f.window_exists("#32770", "Εκτύπωση", timeout=15):
                        opened = True
                        break
                    if attempt < ATTEMPTS - 1:
                        f.wait(DELAY)
                if not opened:
                    f.log("η «Εκτύπωση» δεν άνοιξε — PDF από φωτογραφία σελίδας", level="WARNING")
                    out = f.browser_screenshot_pdf(f"{path} (φωτογραφία).pdf")
                    f.set_result("PDF", out)
                    attach_receipt(f, out, "αποδεικτικό (φωτογραφία σελίδας)")
                    delete_work_files(f, [xlsx], folder, [f"{afm}_{year}_{month}_*.xlsx"])
                    return finish(f, f"ΥΠΟΒΛΗΘΗΚΕ — PDF φωτογραφία: {out}",
                                  pdf_file=f"{path} (φωτογραφία).pdf".split("\\")[-1])
            else:
                aade_click(f, "Ταυτότητα Οφειλής", "Button", timeout=30)
                f.wait(3)

                def open_print(f):
                    if not f.window_exists("#32770", "Εκτύπωση"):
                        with f.time_limit(90):
                            f.browser_click_uia("Εκτύπωση", window_class=FIREFOX, control_type="Button",
                                                exact=True, index=0, timeout=30)
                    f.wait_for("Εκτύπωση", class_name="#32770", timeout=15)
                retry(f, open_print)

            # «Εκτύπωση»: Microsoft Print to PDF (the default printer differs per PC:
            # LENOVO-PC SEC8425192BF882) → ΟΚ (id 1) → «Αποθήκευση εξόδου εκτύπωσης ως».
            SAVE_PDF = "Αποθήκευση εξόδου εκτύπωσης ως"

            def printer(f):
                if f.window_exists("#32770", SAVE_PDF):
                    return
                f.connect_window(window_class="#32770", window_title="Εκτύπωση", timeout=15)
                f.select("Microsoft Print to PDF", target="ComboBox", class_name="ComboBox", index=0,
                         select_method="message")
                f.click("", class_name="Button", control_id=1, click_method="message")
                f.wait_for(SAVE_PDF, class_name="#32770", timeout=15)
            retry(f, printer)

            def pdf_file_name(f):
                f.connect_window(window_class="#32770", window_title=SAVE_PDF, timeout=15)
                f.fill(path, target="Edit", class_name="Edit", fill_method="message")
            retry(f, pdf_file_name)
            f.click("", class_name="Button", control_id=1, click_method="message")   # «&Αποθήκευση» = id 1
            f.wait_for_file(f"{path}.pdf", timeout=60)
            f.set_result("PDF", f"{path}.pdf")
            attach_receipt(f, f"{path}.pdf", "αποδεικτικό")   # Lefteris v8: the PDF to the cloud («Υποβολές»)
            delete_work_files(f, [xlsx], folder, [f"{afm}_{year}_{month}_*.xlsx"])   # Petros 20/11: only the PDF stays
            # Acrobat ανοίγει και στη μηδενική (ΓΕΩΡΓΑΝΤΑΣ 05/11: 1″ μετά το PDF) → πάντα έως 20″
            return finish(f, f"ΥΠΟΒΛΗΘΗΚΕ + PDF: {path}.pdf", acrobat_wait=20,
                          pdf_file=f"{path}.pdf".split("\\")[-1])
