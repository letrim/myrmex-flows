"""
ΑΠΔ.py — ΑΠΔ (ΕΦΚΑ / ΤΕΚΑ) for one company and period, on every office PC.

    C:\\Myrmex\\myrmex.exe flows\\ΑΠΔ.py ^
        --var period.year=2026 --var period.month=10 ^
        --var client.afm=123456789 --var client.name="ΕΤΑΙΡΙΑ ΑΕ" ^
        --var output.folder=C:\\Logimate\\exports_apd\\ΕΤΑΙΡΙΑ ΑΕ

period.month = the month OF THE ΑΠΔ: one month back (November → ΑΠΔ 10).

PHASE 1 (09/11/2026): up to the Μισθοδοτική Κατάσταση Excels.
PHASE 2 (09/11/2026): Εισφορές ανά Ταμείο → PDF and έλεγχος ταμείων (ΚΥΡ of the
Μισθοδοτικές = ΕΦΚΑ of the PDF, ΕΠΙΚ = ΤΕΚΑ + other auxiliary funds).
PHASE 3 (09/11/2026, from fmy_python\\ΑΠΔ.py step_apd): only when the funds
agree — ΑΠΔ ΕΦΚΑ («ΑΠΔ» → 1st of 9) and ΤΕΚΑ («ΑΠΔ ΤΕΚΑ» → 1st of 2, only when
the PDF has ΤΕΚΑ): Νέος υπολογισμός, 01 Κανονική, the month → «Υπολογισμός» →
[«Έχετε ήδη δημιουργήσει Α.Π.Δ.» → Αποθηκευμένα → Διαγραφή → again] →
[«Ελλιπή στοιχεία» → photo and STOP, Petros 09/11/2026 «σταματάμε τη λειτουργία
και το αναφέρουμε»; «Μη επανυπολογισμένοι περίοδοι» → OK] → «Επόμενη» →
«Εξαγωγή» (CSL01) → CSL01 check (period, header = lines, καταβλητέες = PDF
εισφορών, every employee of the Μισθοδοτική with contributions inside) →
«Αποστολή ΑΠΔ» → EpsilonSubmit: Έλεγχος Υποβολής → Υποβολή → Συνέχεια →
«Υποβολή & Έκδοση Αποδεικτικού» (once, f.committed).
test=yes (or the cloud's submit ≠ yes) stops right BEFORE «Αποστολή ΑΠΔ».
10/11/2026 (Lefteris's v1–v3, merged into v4 on 23/11/2026): the receipt opens in the PC's PDF app.
A PDF program (Acrobat, Reader, SumatraPDF, Foxit, PDF-XChange) is closed; in Chrome / Edge ONLY the
receipt's tab is closed (Ctrl+W on the window that shows its file name), never the browser. The
FastReport «Εξαγωγή σε PDF» and the saved ΑΠΔ's «Διαγραφή» are clicked only when their tooltip says so
(expect_hint). The receipt also goes to the cloud («Υποβολές», f.attach, agent 0.4.2).
v4 (23/11/2026, agent 0.4.2 / SDK 1.4): known reasons fail with retryable=False (never run again by
themselves); after a full submission the Μισθοδοτική Excel(s) and the εισφορές PDF are deleted — only
the receipt(s) stay in the company's folder (Petros 20/11/2026).

Up to «Μισθοδοτική Κατάσταση» the flow is THE SAME as flows\\ΦΜΥ.py (version 5
+ 09/11/2026): Epsilon as «logimate» picked from the list, «Επιλέξτε χρήση» /
Οδηγός without a fixed 60 s wait, «Σφάλμα» only from Epsilon's own messages,
the ΑΦΜ in the first column's filter with only connect_window (no focus /
maximize), ΚΑΔ → «Όχι», the Μισθοδοτική as the row before the last
(PAYROLL_ROWS), finish() closes Acrobat / Firefox / Epsilon.

ΑΠΔ specifics (fmy_python\\ΑΠΔ.py, confirmed with ΕΥΘΥΜΙΟΥ 20/10/2026 and
ΝΙΚΟΛΑΟΥ 27/10/2026):
 - The seasonal gift goes to the ΑΠΔ of the NEXT month: Δώρο Πάσχα (April) →
   ΑΠΔ May, Δώρο Χριστουγέννων (December) → ΑΠΔ January. Everything else
   (Επίδομα Αδείας, gift shares on leaving) stays in its month.
 - ΑΠΔ May / January: TWO Μισθοδοτικές — the ΑΠΔ month and the previous one
   (April / December of the year before); from the previous one only its gift.
 - ΑΠΔ April / December: from its own Μισθοδοτική the gift is left out.
 - The 2nd month in a NEW Μισθοδοτική window: with only the filters changed
   Epsilon brought fewer employees (ΝΙΚΟΛΑΟΥ 12/2025 after 1/2026: 8 instead
   of 11 — Petros 27/10/2026).

Outcomes (f.done = success with a message):
    ΥΠΟΒΛΗΘΗΚΕ — ΕΦΚΑ αρ. … · ΤΕΚΑ αρ. …
    ΔΟΚΙΜΗ — test: the CSL01 agree, stopped before «Αποστολή ΑΠΔ»
    ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ — Epsilon / ΕΦΚΑ already has it: nothing sent again
    0 ΕΓΓΡΑΦΕΣ · ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ — for the ΑΠΔ month
Anything else fails the client with «<reason> — ΔΕΙΤΕ ΤΟ SCREENSHOT: <photo>» (problem(), Petros
20/11/2026). A known reason (ΑΣΥΜΦΩΝΙΑ ΤΑΜΕΙΩΝ / CSL01, ΕΡΓΑΖΟΜΕΝΟΙ ΕΚΤΟΣ ΑΠΔ, ΛΑΘΗ ΣΤΟΝ ΕΛΕΓΧΟ, ΛΑΘΟΣ
ΚΩΔΙΚΟΙ, …) with retryable=False: never run again by itself. A message nobody expected → f.halt
(Epsilon stays open, the office gets an email). After the final Υποβολή (f.committed) it is reported
as not to be run again.

The cloud sends submit=yes|no with every job: yes → submit; anything else → test.
Without `submit` (a local run) it submits. test=yes always wins (as flows\\ΦΜΥ.py).
The receipt PDF lands in Downloads (f.downloads_folder(), SDK 1.2 — or --var
downloads.folder=… to look elsewhere): it is checked and moved to the company's folder
as «<ΟΝΟΜΑ> ΑΠΔ MM YYYY.pdf» / «<ΟΝΟΜΑ> ΑΠΔ ΤΕΚΑ MM YYYY.pdf».
Needs the Epsilon user «logimate» (once per office).
"""

import re
import unicodedata
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from myrmex_agent.sdk import Param

FLOW_VERSION = "6"        # cloud «ΑΠΔ»: 6 = «ΑΠΔ» / «ΑΠΔ ΤΕΚΑ»: the 1st row of however many the menu search lists
                          # (Hyper 9 / 2, Epsilon Extra 1 — Φραντζής 25/11/2026: rows=9 picked nothing there);
                          # 5 = Petros's runs 13–23/11 (Open period, pending ΑΠΔ, problem() messages,
                          # ΑΣΥΜΦΩΝΙΑ = problem, halt, retryable, delete_file, review.accept) merged with Lefteris's
                          # v4 (Hyper or Extra, no local paths in fail messages) — 24/11/2026;
                          # 4 (Lefteris) = Hyper or Extra (ExtraM.exe), fail_shot; 3 = the receipt also to the cloud
                          # (f.attach, agent 0.4.2); 2 = v1 for the MYRMEX agent 0.4.0; 1: phases 1–3 (09/11/2026)
MIN_SDK_VERSION = "1.4"   # retryable / delete_file (1.3), attach (1.4); expect_hint, downloads_folder (1.2)

# Killed by the batch runner before and after every client, on top of the
# agent defaults (HyperM.exe, ExtraM.exe, firefox.exe).
# PDF programs: whichever the PC opens a .pdf with is closed like Acrobat always was (10/11/2026: the default
# differs per PC). Browsers (PDF_BROWSER = Chrome, Edge, any Chromium) are NEVER killed: only reported.
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
    "downloads.folder": Param(str, required=False),   # default f.downloads_folder() (see the docstring)
    # yes = at «Ελλιπή Στοιχεία Εργαζομένων» press OK and go on — ONLY for a company the accountant approved
    # (Νίνα 23/11/2026: ΠΛΟΥΚΙΔΗΣ ΙΚΕ λοχεία, ΑΔΑΜΙΔΗΣ, ΜΑΝΟΥΣΟΣ «αγνοείς το μήνυμα και προχωράς»). Without
    # it the flow stops there, always (Petros 20/11/2026). A button of the job in the cloud (Petros 23/11/2026).
    "review.accept": Param(str, required=False, choices=("yes", "no")),
    # yes = go up to the last EpsilonSubmit page and STOP there (f.halt, everything open) WITHOUT
    # «Υποβολή & Έκδοση Αποδεικτικού»: to see what that page shows (Petros 13/11/2026)
    # list = stop at the ΕΦΚΑ list of submissions, before «Έλεγχος Υποβολής» (to look at the pending one)
    "stop.before_final": Param(str, required=False, choices=("yes", "no", "list")),
}

EPSILON = r"C:\Program Files (x86)\epsilon net\hmisthos\HyperM.exe"   # default; epsilon_app() finds the real one
EPSILON_ROOTS = (r"C:\Program Files (x86)\epsilon net", r"C:\Program Files\epsilon net")
# Hyper or Extra (the lighter edition: same windows, process ExtraM.exe, folder ExtraM) — known folders first.
EPSILON_PATTERNS = ("hmisthos/HyperM.exe", "ExtraM/ExtraM.exe", "*/HyperM.exe", "*/ExtraM.exe")
EPSILON_USER = "logimate"     # empty password; made by ΧΡΗΣΤΗΣ_epsilon.py
# «Μισθοδοτική Κατάσταση»: rows of the search list, most likely first. The pick
# is always rows - 1 (the row before «Σχεδιαζόμενη Μισθοδοτική Κατάσταση»).
PAYROLL_ROWS = (4, 3, 5, 6)   # 4 = reference PC, 3 = LENOVO-PC, 5 = admin
PAYROLL_WAIT = 20             # s for the window after a pick (LENOVO-PC: 6.2 s)
ATTEMPTS, DELAY = 3, 10

MONTHS = {
    1: "Ιανουάριος", 2: "Φεβρουάριος", 3: "Μάρτιος", 4: "Απρίλιος",
    5: "Μάιος", 6: "Ιούνιος", 7: "Ιούλιος", 8: "Αύγουστος",
    9: "Σεπτέμβριος", 10: "Οκτώβριος", 11: "Νοέμβριος", 12: "Δεκέμβριος",
}

# Gifts (ΕΥΘΥΜΙΟΥ 20/10/2026): EXTRA_PERIODS = from the 2nd month's Μισθοδοτική
# ONLY this «Περίοδος»; OWN_EXCLUDED = from the ΑΠΔ month's own Μισθοδοτική NOT
# this one (it goes to the next month's ΑΠΔ).
EXTRA_PERIODS = {5: ("Δώρο Πάσχα",), 1: ("Δώρο Χριστουγέννων",)}
OWN_EXCLUDED = {4: ("Δώρο Πάσχα",), 12: ("Δώρο Χριστουγέννων",)}

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


def apd_months(year, month) -> list:
    """The Μισθοδοτικές to export: the ΑΠΔ month, and the previous one ONLY for
    May (Δώρο Πάσχα) / January (Δώρο Χριστουγέννων, December of the year before)."""
    if month == 5:
        return [(year, 5), (year, 4)]
    if month == 1:
        return [(year, 1), (year - 1, 12)]
    return [(year, month)]


def periods_from_excel(f, rows) -> dict:
    """The Μισθοδοτική's Excel: ΚΥΡ / ΕΠΙΚ («Σύν. Εισφ. Κύριου / Επίκ. Ταμείου»)
    per «Περίοδος» (month, gifts, allowances), employee rows only (they have an
    «Εταιρεία»; the totals row doesn't). {period: (κυρ, επικ, rows)}."""
    header = next((i for i, r in enumerate(rows) if any(norm(c) == "περιοδος" for c in r)), None)
    if header is None:
        problem(f, "δεν βρέθηκε επικεφαλίδα «Περίοδος» στο Excel της Μισθοδοτικής")
    cols = {norm(h): i for i, h in enumerate(rows[header])}
    kyr, epik = "συν. εισφ. κυριου ταμειου", "συν. εισφ. επικ. ταμειου"
    missing = [h for h in ("εταιρεια", "περιοδος", kyr, epik) if h not in cols]
    if missing:
        problem(f, f"το Excel δεν έχει στήλη {', '.join(missing)}")

    def cell(row, name):
        i = cols.get(name)
        return row[i] if i < len(row) else None

    out = {}
    for r in rows[header + 1:]:
        if str(cell(r, "εταιρεια") if cell(r, "εταιρεια") is not None else "").strip() == "":
            continue
        p = str(cell(r, "περιοδος") or "").strip() or "?"
        k, e, n = out.get(p, (ZERO, ZERO, 0))
        out[p] = (k + amount(cell(r, kyr)), e + amount(cell(r, epik)), n + 1)
    return out


def apd_totals(apd_month, per_month) -> tuple:
    """ΚΥΡ / ΕΠΙΚ for the ΑΠΔ: the ΑΠΔ month's periods without OWN_EXCLUDED, plus
    from the 2nd month only EXTRA_PERIODS. per_month = [periods of the ΑΠΔ month,
    periods of the 2nd month (May / January only)]. Returns (κυρ, επικ, parts)."""
    kyr, epik, parts = ZERO, ZERO, []
    skip = [norm(x) for x in OWN_EXCLUDED.get(apd_month, ())]
    for p, (k, e, _) in (per_month[0] if per_month else {}).items():
        if norm(p) in skip:
            parts.append(f"χωρίς {p} (πάει στην επόμενη ΑΠΔ)")
            continue
        kyr, epik = kyr + k, epik + e
        parts.append(f"{p} {k}/{e}")
    if len(per_month) > 1:
        take = [norm(x) for x in EXTRA_PERIODS.get(apd_month, ())]
        for p, (k, e, _) in per_month[1].items():
            if norm(p) in take:
                kyr, epik = kyr + k, epik + e
                parts.append(f"+ {p} {k}/{e}")
    return kyr, epik, parts


def last_day(year, month) -> int:
    return (date(year + (month == 12), month % 12 + 1, 1) - timedelta(days=1)).day


def contrib_totals(text) -> dict:
    """The PDF «Ασφαλιστικές Εισφορές ανά Ταμείο» → {ταμείο: (Εργαζ.+Εργοδ., στήλη «Σύνολο», επιδότηση)}.
    The line after the header «Εργαζομένου Εργοδότη …» or after every «Σύνολο :» is the fund's name
    (ΕΦΚΑ, TEKA, ΤΕΑΥΦΕ, …). For the check Εργαζ. + Εργοδ. counts: the «Σύνολο» column differs when
    there is a ΔΥΠΑ subsidy (as fmy_python\\ΑΠΔ.py)."""
    tot, section, expect = {}, None, False
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("Εργαζομένου Εργοδότη"):
            expect = True
            continue
        if s.startswith("Σύνολο Εταιρείας") or s.startswith("Γενικό Σύνολο"):
            break
        t = re.match(r"Σύνολο :\s*([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", s)
        if t:
            if section and section not in tot:
                tot[section] = (amount(t.group(1)) + amount(t.group(2)), t.group(4), amount(t.group(3)))
            section, expect = None, True
            continue
        if expect and s:
            section = {"TEKA": "ΤΕΚΑ"}.get(s, s)        # Latin «TEKA» in the PDF
            expect = False
    return tot


# ── Epsilon helpers (as flows\ΦΜΥ.py) ──────────────────────────────────────

def retry(f, fn):
    """3 attempts, 10 s apart."""
    f.retry(fn, attempts=ATTEMPTS, delay=DELAY)


def bring(f, window_class, title=None, timeout=None):
    """Connect to a TOP-LEVEL window and bring it to the front. Not for tabs
    inside TfMain (TfCompanies): connect_window doesn't find those."""
    f.connect_window(window_class=window_class, window_title=title, timeout=timeout)
    f.focus_window(window_class=window_class, window_title=title)


def must_close(f, window_class, title, timeout):
    """The window must close (or hide, for a tab) within `timeout` s."""
    f.wait_for_gone(title, class_name=window_class, timeout=timeout)


def problem(f, reason, known=False, shot=None):
    """Every stop of a company (Petros 20/11/2026): the reason, then «ΔΕΙΤΕ ΤΟ SCREENSHOT: <photo>».
    known=True: a reason a person must fix (ασυμφωνία, εργαζόμενοι εκτός ΑΠΔ, λάθη στον έλεγχο, κωδικοί, …)
    → f.fail(retryable=False): the cloud never runs it again by itself (agent 0.4.1, Lefteris 22/11/2026);
    the message also says «ΔΕΝ ΞΑΝΑΤΡΕΧΕΙ» for the person reading it. known=False: technical (a click, a
    window that didn't open, a timeout) → it may be run again."""
    # Lefteris v4 (fail_shot, 23/11/2026): NO local path in the message — it goes to the cloud and to the
    # office's email, where «C:\\Myrmex\\…png» means nothing; the cloud shows the agent's own failure
    # screenshot on the job page. The photo's path goes to the log.
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
    """A TRzMenuButton's menu → item downs + 1 (Down × downs, Enter).
    «Εξαγωγή σε EXCEL» = XLS, XLSX on both PCs."""
    bring(f, window_class, title, timeout=60)
    f.click(target, class_name="TRzMenuButton")
    f.wait(1)
    for _ in range(downs):
        f.press_key("down")
        f.wait(pause)
    f.press_key("enter")


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


def pdf_app_open(f) -> bool:
    """Is a PDF program (PDF_APPS) showing a window?"""
    return any(f.window_exists(cls) for cls, _ in PDF_APPS)


def kill_pdf_apps(f, also=()):
    """The PDF programs (and `also`) are killed; never a browser. Never fails the client."""
    f.kill_process(PDF_PROCESSES + list(also), timeout=10, ignore_missing=True, on_error="continue")


def close_pdf_tab(f, name) -> bool:
    """The receipt open as the ACTIVE tab of Chrome / Edge (window title «<name> - Google Chrome»): ONLY that
    tab closes — Ctrl+W on that window, never the browser (the user's Gmail etc. stay). The window is found
    by the receipt's own file name, so no other tab matches. True when it closed. Never fails the client."""
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
    f.log(f"αποδεικτικό στον browser: η καρτέλα «{name}» " + ("έκλεισε" if closed else "ΔΕΝ έκλεισε"),
          level="INFO" if closed else "WARNING")
    return closed


def close_pdf_viewers(f, wait=0, also=()):
    """PDF programs (PDF_APPS) close: up to `wait` s for the PDF to open first, then a 2nd pass up to 10 s
    (it can open late, ΓΕΩΡΓΑΝΤΑΣ 05/11). A receipt in a browser tab is closed in receipt(), by its name."""
    for _ in range(int(wait)):
        if pdf_app_open(f):
            break
        f.wait(1)
    kill_pdf_apps(f, also)
    for _ in range(10):
        if pdf_app_open(f):
            kill_pdf_apps(f)
            break
        f.wait(1)


def attach_receipt(f, path, kind):
    """Το αποδεικτικό και στο cloud («Υποβολές»), για να το βρει ο λογιστής από οποιονδήποτε υπολογιστή.
    Μόνο agent 0.4.2+ (SDK 1.4): σε παλιότερο agent το flow τρέχει όπως πριν και το PDF μένει μόνο εδώ.
    on_error="continue": ο πελάτης έχει ήδη υποβληθεί — το συνημμένο δεν τον ρίχνει ποτέ.
    (Lefteris 22/11/2026: keep this call wherever a receipt is saved.)"""
    if hasattr(f, "attach"):
        f.attach(path, kind=kind, on_error="continue")


def delete_work_files(f, paths, folder=None, patterns=()):
    """Petros 20/11/2026: after the submission only the receipt stays in the company's folder — the
    Μισθοδοτική Excel(s) and the εισφορές PDF are deleted (f.delete_file, agent 0.4.1: only inside the
    client's folder / Downloads). Only on a full submission: on a problem they stay, to see what happened.
    Petros 23/11/2026: also the SAME month's leftovers of earlier runs (`patterns` in `folder`: older
    Excels, «_παλιό_» εισφορές — ΠΕΤΣΑΣ ΕΕ had 6 + 6 from 13–14/11); never another month, never a receipt.
    Never fails the client (it is already submitted)."""
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


def finish(f, message, acrobat_wait=0):
    """Τέλος πελάτη όπως στα παλιά scripts: κλείνουν τα προγράμματα PDF, Firefox και Epsilon και μετά f.done.
    Χωρίς batch (μία εταιρία με --var) η μηχανή δεν κάνει state reset και έμεναν ανοιχτά.
    Σε αποτυχία (f.fail) ΔΕΝ κλείνει τίποτα: η μηχανή κρατά φωτογραφία και ανοιχτά παράθυρα.
    acrobat_wait: έως τόσα ″ να ανοίξει πρώτα το PDF σε πρόγραμμα PDF."""
    # on_error="continue": το κλείσιμο δεν ρίχνει ποτέ τον πελάτη (ΦΜΥ 09/11/2026: «could not kill Acrobat.exe
    # … PID not found» έβγαλε FAILED τρεις υποβληθείσες εταιρίες)
    close_pdf_viewers(f, wait=acrobat_wait, also=["firefox.exe"] + EPSILON_PROCESSES)
    return f.done(message)


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


PAY = ("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση")


def close_payroll(f):
    """WM_CLOSE (like its ×, not Alt+F4); it must close."""
    if f.window_exists(*PAY):
        f.connect_window(window_class=PAY[0], window_title=PAY[1], timeout=5)
        f.close_app()
    must_close(f, PAY[0], PAY[1], 30)


def export_payroll(f, afm, folder, year, month, n) -> dict:
    """One Μισθοδοτική Κατάσταση (year / month) → Excel, read per «Περίοδος».
    As flows\\ΦΜΥ.py's export_excel + save_export. n = 1 for the ΑΠΔ month, 2 for
    the previous one (own store_as for each read, so a scenario answers each).
    {status: ok | 0 ΕΓΓΡΑΦΕΣ | ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ, xlsx, periods, shot}."""
    month_name = MONTHS[month]
    filename = f"{afm}_ΜΙΣΘΟΔΟΤΙΚΗ_{year}_{month:02d}_{f.now()}"
    f.log(f"── {month_name} {year}: Μισθοδοτική → Excel")
    f.wait(1.5)

    # 3 rounds over PAYROLL_ROWS, 10 s apart; none opened it → fail.
    for attempt in range(ATTEMPTS):
        if open_payroll_window(f):
            break
        if attempt < ATTEMPTS - 1:
            f.wait(DELAY)
    else:
        shot = f.screenshot(store_as=f"payroll_menu_shot_{n}")
        problem(f, f"δεν άνοιξε η «Μισθοδοτική Κατάσταση» από την αναζήτηση μενού (δοκιμάστηκαν "
                   f"{list(PAYROLL_ROWS)} γραμμές, {ATTEMPTS} φορές) — πρόσθεσε τον αριθμό γραμμών στο PAYROLL_ROWS",
                shot=shot)
    retry(f, lambda f: f.wait_for(PAY[1], class_name=PAY[0], timeout=30))
    f.wait(2)
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

    # «Επιλογή Όλων»: painted text on the bar of «Εξαγωγή σε EXCEL», found by its pixels.
    bar = {"target": "Εξαγωγή σε EXCEL", "class_name": "TRzMenuButton"}
    runs = []
    for attempt in range(ATTEMPTS):
        bring(f, *PAY, timeout=30)
        f.wait(0.5)
        runs = f.text_runs_in_control(bar, parent=True, max_x=400, store_as=f"select_all_runs_{n}")
        if runs:
            break
        if attempt < ATTEMPTS - 1:
            f.wait(DELAY)
    if not runs:
        shot = f.screenshot(store_as=f"no_records_shot_{n}")
        retry(f, close_payroll)
        return {"status": "0 ΕΓΓΡΑΦΕΣ", "shot": shot, "periods": {}}
    f.click_text_in_control(bar, parent=True, run=0, max_x=400)
    f.wait(1)

    # «Εξαγωγή σε EXCEL» → Down Down Enter (XLSX) → «Αποθήκευση ως», or «Πληροφορία» (no employees).
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
        retry(f, close_payroll)
        return {"status": "ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ", "periods": {}}

    save_as(f, "Αποθήκευση", f"{folder}\\{filename}")
    f.wait(2)
    # «Πληροφορία – Η Εξαγωγή ολοκληρώθηκε με επιτυχία» → ΟΚ
    if f.window_exists("TMessageForm", timeout=10):
        f.connect_window(window_class="TMessageForm", timeout=5)
        f.click("", class_name="TButton", index=0, click_method="message")
        f.wait_for_gone("", class_name="TMessageForm", timeout=10, on_timeout="continue")
    # Closed after EVERY month: the next one in a new window (see the docstring),
    # and TfMain free for what comes next.
    retry(f, close_payroll)

    def find_export(f):
        exported = f.find_file(folder, "*.xlsx", newest=True, newer_than=export_started,
                               timeout=5, required=True, store_as=f"exported_file_{n}")
        if exported["name"] != f"{filename}.xlsx":
            f.rename_file(exported["path"], f"{filename}.xlsx")
    f.retry(find_export, attempts=6, delay=5)
    xlsx = f"{folder}\\{filename}.xlsx"
    f.log(f"Excel: {xlsx}")

    rows = f.excel_read(xlsx, as_lists=True, store_as=f"payroll_rows_{n}")
    periods = periods_from_excel(f, rows)
    text = " | ".join(f"{p}: ΚΥΡ {k} / ΕΠΙΚ {e} ({c} γρ.)" for p, (k, e, c) in periods.items()) or "-"
    f.log(f"Excel {month:02d}/{year} ανά περίοδο: {text}")
    f.set_result(f"Μισθοδοτική {month:02d}/{year}", text)
    return {"status": "ok", "xlsx": xlsx, "periods": periods, "rows": rows}   # rows: CSL01 check


EISF = ("Tfeisfores", "Επιλογή Εργαζομένων για Εκτύπωση")
PREVIEW = "TfrxPreviewForm"
PDF_DIALOG = "TfrxPDFExportDialog"
# FastReport preview toolbar (TToolBar, a real control): «Εξαγωγή σε PDF» is its 7th button, x 73–96,
# y 4–27 inside the toolbar (reference PC 09/11/2026, 1920×1080). Its buttons have NO UIA name — only
# a tooltip says «Εξαγωγή σε PDF» (fmy_python\ΑΠΔ.py checks it by hovering). So: a click at this
# offset ONLY when its tooltip says so (expect_hint, SDK 1.2), then the PDF dialog MUST open.
PDF_BUTTON = (84, 15)


def contributions_pdf(f, afm, folder, year, month) -> dict:
    """Ασφαλιστικές Εισφορές ανά Ταμείο, 1/μ – τέλος μήνα, «Βάσει Ημ.Υποβ.ΑΠΔ.» → PDF → read back.
    As fmy_python\\ΑΠΔ.py step_contributions. {pdf, ΕΦΚΑ, ΤΕΚΑ, άλλα}."""
    end = last_day(year, month)
    want_from, want_to = f"1/{month}/{year}", f"{end}/{month}/{year}"
    path = f"{folder}\\{afm}_ΕΙΣΦΟΡΕΣ_{year}_{month:02d}"
    old = f.find_file(folder, f"{afm}_ΕΙΣΦΟΡΕΣ_{year}_{month:02d}.pdf", required=False, store_as="old_contrib_pdf")
    if old and old.get("found") in (True, "true"):
        f.rename_file(old["path"], f"{afm}_ΕΙΣΦΟΡΕΣ_{year}_{month:02d}_παλιό_{f.now()}.pdf")

    def open_dialog(f):
        if f.window_exists(*EISF):
            return
        if not f.window_exists("TfChooseCompanies", "Επιλογή Εταιριών"):
            f.navigate_menu("Ασφαλιστικές Εισφορές ανά Ταμείο", rows=1)   # 1 row (logimate, 09/11/2026)
        if f.window_exists("TfChooseCompanies", "Επιλογή Εταιριών", timeout=30):
            f.connect_window(window_class="TfChooseCompanies", window_title="Επιλογή Εταιριών", timeout=5)
            f.click("Επιλογή", class_name="TBitBtn", click_method="message")
        f.wait_for(EISF[1], class_name=EISF[0], timeout=30)
    retry(f, open_dialog)
    bring(f, *EISF, timeout=30)

    retry(f, lambda f: f.set_checkbox("Ημερομηνίες", class_name="TGroupButton"))
    retry(f, lambda f: f.set_checkbox("Βάσει Ημ.Υποβ.ΑΠΔ.", class_name="TCheckBox"))

    def dates(f):
        # From the MIDDLE: month/year first, then the day — else «31» falls on the field's current
        # month (03/11/2026: a November field rejected «31» and the digits went to wrong places).
        for position, day in (("leftmost", 1), ("rightmost", end)):
            f.focus_control("", class_name="TDateTimePicker", prefer=position)
            for _ in range(3):
                f.press_key("left")
            f.press_key("right")
            f.type_text(f"{month:02d}/{year}")
            f.press_key("left")
            f.press_key("left")
            f.type_text(f"{day:02d}")
            f.wait(0.5)
        shown = {f.get_control_text(class_name="TDateTimePicker", index=1, store_as="date_1"),
                 f.get_control_text(class_name="TDateTimePicker", index=2, store_as="date_2")}
        f.check(shown == {want_from, want_to}, f"οι ημερομηνίες γράφουν {sorted(shown)} αντί για {want_from} – {want_to}")
    retry(f, dates)
    f.log(f"Ημ/νία από {want_from} έως {want_to}")

    # «Εκτύπωση»: painted on the dialog's bottom bar — Αποστολή E-Mail | Εκτύπωση | Έξοδος
    # [| Αποστολή SMS / Viber — admin only, not logimate (03/11/2026)] → the 2nd run.
    # The bar is the dialog's bottom TPanel (569,678)–(1351,717) on the reference PC; it holds a smaller
    # TPanel (574–807, no text), which is the «bottommost» one → its PARENT is the bar (09/11/2026).
    bar = {"target": "", "class_name": "TPanel", "prefer": "bottommost"}

    # 3 rounds, 10 s apart, outside f.retry: a failure says what it saw, not «retry exhausted».
    runs = []
    for attempt in range(ATTEMPTS):
        if f.window_exists(PREVIEW):
            break
        bring(f, *EISF, timeout=30)
        runs = f.text_runs_in_control(bar, parent=True, store_as="print_bar_runs")
        if len(runs) in (3, 4):
            f.click_text_in_control(bar, parent=True, run=1, capture="window")
            if f.window_exists(PREVIEW, timeout=60):
                break
        if attempt < ATTEMPTS - 1:
            f.wait(DELAY)
    else:
        shot = f.screenshot(store_as="print_bar_shot")
        if len(runs) not in (3, 4):
            problem(f, f"η κάτω μπάρα των Εισφορών έχει {len(runs)} κείμενα αντί για 3 ή 4 — δεν πατάω", shot=shot)
        problem(f, "η «Εκτύπωση» των Εισφορών δεν άνοιξε την «Προεπισκόπηση»", shot=shot)

    # «Εξαγωγή σε PDF»: a position on the toolbar (see PDF_BUTTON), ONCE — a wrong button would open
    # something else; the result (the PDF dialog) is the check.
    f.check(f.screen_size(store_as="screen") == [1920, 1080], "η θέση του «Εξαγωγή σε PDF» είναι για 1920×1080")
    bring(f, PREVIEW, timeout=30)
    # focus_window RESTORES it (agent 0.3.2; fixed in 0.3.3): 09/11/2026 the preview shrank to
    # (200,127)–(320,155) and the click landed at (284,171) on Epsilon's menu → maximize first.
    f.maximize_window(window_class=PREVIEW)
    f.wait(1)
    if not f.window_exists(PDF_DIALOG):
        f.click("", method="relative_coordinates", class_name="TToolBar",
                offset_x=PDF_BUTTON[0], offset_y=PDF_BUTTON[1], expect_hint="Εξαγωγή σε PDF")
    if not f.window_exists(PDF_DIALOG, timeout=20):
        shot = f.screenshot(store_as="pdf_button_shot")
        problem(f, "δεν άνοιξε το «Εξαγωγή σε PDF» μετά το κλικ στη γραμμή εργαλείων", shot=shot)

    # As the PAD (Tab Tab Space …): ✔ PDF/A, ✔ fonts, ✔ «Βέλτιστη Εκτύπωση» — by message, read back.
    def pdf_options(f):
        f.connect_window(window_class=PDF_DIALOG, timeout=10)
        for name in ("PDF/A", "Συμπεριλαμβάνει γραμματοσειρές", "Βέλτιστη Εκτύπωση"):
            f.set_checkbox(name, class_name="TCheckBox")
    retry(f, pdf_options)

    def ok(f):
        if f.window_exists(PDF_DIALOG):
            f.connect_window(window_class=PDF_DIALOG, timeout=5)
            f.click("OK", class_name="TButton", click_method="message")
        f.wait_for("Αποθήκευση", class_name="#32770", timeout=20)
    retry(f, ok)
    started = f.now_epoch()
    save_as(f, "Αποθήκευση", f"{path}.pdf")
    # Epsilon EXTRA (ΦΡΑΝΤΖΗΣ 24/11/2026): the name typed into the save dialog didn't take — the PDF was saved
    # as a bare «.pdf» in the company's folder, and waiting for «…_ΕΙΣΦΟΡΕΣ_….pdf» failed after 60 s. As for the
    # Μισθοδοτική's Excel: the newest PDF written in the folder since the save is ours → renamed if needed.
    want = f"{afm}_ΕΙΣΦΟΡΕΣ_{year}_{month:02d}.pdf"
    saved = f.find_file(folder, "*.pdf", newest=True, newer_than=started - 2, timeout=60, required=False,
                        store_as="contrib_pdf_saved")
    if saved and saved.get("found") in (True, "true") and saved.get("name") != want:
        f.log(f"το PDF εισφορών αποθηκεύτηκε ως «{saved['name']}» — μετονομασία σε «{want}»", level="WARNING")
        f.rename_file(saved["path"], want)
    f.wait_for_file(f"{path}.pdf", timeout=30, newer_than=started - 2)

    text = f.pdf_text(f"{path}.pdf", store_as="contrib_pdf_text")
    per = re.search(r"Περίοδος:\s*(\S+)\s*-\s*(\S+)", text)
    pdf_afm = re.search(r"Α\.Φ\.Μ\.\s*:\s*(\d+)", text)
    f.check(per is not None and per.groups() == (want_from, want_to),
            f"το PDF εισφορών γράφει περίοδο {per.groups() if per else '-'} αντί για {want_from} – {want_to}")
    f.check(pdf_afm is not None and pdf_afm.group(1) == afm,
            f"το PDF εισφορών είναι για ΑΦΜ {pdf_afm.group(1) if pdf_afm else '-'} αντί για {afm}")
    tot = contrib_totals(text)
    efka, teka = (tot["ΕΦΚΑ"][0] if "ΕΦΚΑ" in tot else ZERO), (tot["ΤΕΚΑ"][0] if "ΤΕΚΑ" in tot else ZERO)
    others = {k: v[0] for k, v in tot.items() if k not in ("ΕΦΚΑ", "ΤΕΚΑ")}   # e.g. ΤΕΑΥΦΕ (pharmacies)
    f.log(f"Εισφορές {month:02d}/{year} (Εργαζ.+Εργοδ.): ΕΦΚΑ {efka} · ΤΕΚΑ {teka}"
          + "".join(f" · {k} {v}" for k, v in others.items()))

    # The PDF is saved → Προεπισκόπηση and «Επιλογή Εργαζομένων για Εκτύπωση» close (like their ×),
    # as the PAD's 2 Escapes.
    for cls, title in ((PREVIEW, ""), EISF):
        def close(f, cls=cls, title=title):
            if f.window_exists(cls, title or None):
                f.connect_window(window_class=cls, window_title=title or None, timeout=5)
                f.close_app()
            f.wait_for_gone(title, class_name=cls, timeout=15)
        retry(f, close)
    return {"pdf": f"{path}.pdf", "ΕΦΚΑ": efka, "ΤΕΚΑ": teka, "άλλα": others,
            "επιδοτήσεις": {k: v[2] for k, v in tot.items()}}            # for the receipt check


# ── phase 3: ΑΠΔ ΕΦΚΑ / ΤΕΚΑ (as fmy_python\ΑΠΔ.py step_apd) ───────────────

APD = "TfAPD"
# Menu search with the «logimate» user: the right one is ALWAYS the 1st row, whatever the count —
# «ΑΠΔ»: Hyper 9 rows (1st = Εκτυπώσεις :: Ασφαλιστικές :: ΕΦΚΑ :: ΑΠΔ, 96 runs), Epsilon Extra 1 row
# (Φραντζής 25/11/2026); «ΑΠΔ ΤΕΚΑ»: Hyper 2 rows, 1st = ΤΕΚΑ :: ΑΠΔ ΤΕΚΑ. So no rows= (Petros 25/11/2026:
# «την 1η γραμμή απ' όσες βγουν»). The window is checked after: «Επιλογή Εταιριών», then the title —
# «Α.Π.Δ.» is also a part of the titles of the other ΑΠΔ windows, so those words must NOT be in it.
APD_KINDS = {
    "ΕΦΚΑ": {"menu": "ΑΠΔ", "title": "Α.Π.Δ.", "prefix": "ΑΠΔ",
             "not": ("ΤΕΚΑ", "Δημοσ", "Οικοδομ", "Ναυτ")},
    "ΤΕΚΑ": {"menu": "ΑΠΔ ΤΕΚΑ", "title": "Α.Π.Δ. ΤΕΚΑ", "prefix": "ΑΠΔ ΤΕΚΑ",
             "not": ("Οικοδομ",)},
}
# Bottom bar of «Α.Π.Δ.»: painted labels on the bottom TPanel (327,820)–(1592,853) on the reference PC:
# «Υπολογισμός | Ακύρωση» (page 1), «Προηγούμενη | Επόμενη | Ακύρωση» (lists), and the same runs from
# its left edge on every page (x 1071–1130 Υπολογισμός/Επόμενη, 1195–1238 Ακύρωση) — 09/11/2026.
APD_BAR = {"target": "", "class_name": "TPanel", "prefer": "bottommost"}
# «Αποθηκευμένα Στοιχεία» → «Υπολογισμένα Δεδομένα»: painted toolbar (TActionToolBar «acttlbGeneral»).
# With the month's row the ENGINE finds 7 labels (icon + word together; ΚΟΤΖΙΝΟΣ 09/11/2026):
# [6,89] Προηγούμενη | [100,159] Επόμενη | [171,236] Μεταβολή | [248,312] Διαγραφή | [336,486]
# Συμπληρωματικός Υπολ. Α.Π.Δ. | [511,617] Οριστικοποίηση Α.Π.Δ. | [629,746] Αποριστικοποίηση Α.Π.Δ.
# (fmy_python\ΑΠΔ.py counted 8 — it split «Α.Π.Δ.» off; the tooltip of the 4th = «Διαγραφή Α.Π.Δ.»). Greyed (disabled) labels have no dark pixels and
# are not counted: an empty list gives fewer → no click. The buttons have no UIA names → the check is
# the count, the 4th label's tooltip «Διαγραφή…» (expect_hint, SDK 1.2) AND the question that must
# follow: «Διαγραφή» [Ναι / Όχι], else nothing is answered.
APD_TOOLBAR = {"target": "acttlbGeneral", "class_name": "TActionToolBar"}
TOOLBAR_LABELS, DELETE_LABEL = 7, 3
# Labels of that toolbar with an EMPTY list (the month's ΑΠΔ «Υποβλήθηκε», hidden): not seen with the
# engine yet → until then a second «ήδη δημιουργήσει» without a deletion stops (never «ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ»
# on a guess: 09/11/2026 the engine counted 7, nothing was deleted, and that was reported as submitted).
EMPTY_TOOLBAR_LABELS = ()
REVIEW = "TuAPDCaseReview"     # «Ελλιπή Στοιχεία Εργαζομένων / Εταιριών / Αιτία μη εμφάνισης εργαζομένων»
CALC_SECONDS = 240
EXPORT_PAGE = {"class_name": "TBitBtn", "text": "Εξαγωγή"}
# EpsilonSubmit: a window of its own, or INSIDE Epsilon's «Υποβολή ΑΠΔ» (TfESubmitDocs) — 20 and 22/10/2026.
SUBMIT_WIN = {"window_title": "EpsilonSubmit", "inside_class": "TfESubmitDocs"}
ROW = {"row_start": r"^\d{6,}$", "row_start_type": "Hyperlink", "row_end": r"^Περίοδος Υποβολής"}
# Petros 14/11/2026: up to 10 «Αποστολή ΑΠΔ» (each «Open period» try ≈ 30″, nothing uploaded), 5″ apart
# (60″ apart did not help: ΔΑΦΝΗΣ ΟΕ); no EpsilonSubmit at all for 60″ → next try, not 180″
UPLOAD_SECONDS, EMPTY_PAGE_SECONDS, SEND_TRIES = 180, 60, 10
UNREAD_TRIES = 2                # EpsilonSubmit never readable in 2 attempts → ΛΑΘΟΣ ΚΩΔΙΚΟΙ ΕΦΚΑ; (17/11/2026)
SEND_PAUSE, RETRY_PAUSE, NO_SUBMIT_SECONDS = 10, 5, 60
CHECK_SECONDS = 300
FINAL = "Υποβολή & Έκδοση Αποδεικτικού"


def epsilon_message(f, after):
    """Any Epsilon message (TMessageForm) still open here was not expected. Epsilon builds messages in
    stages: dialog_buttons waits for them. Petros 20/11/2026:
     - an INFORMATIONAL message = «Πληροφορία» with only [ΟΚ] (and no other titled message) → photo, ΟΚ, go on,
       noted in the result («ενημέρωση Epsilon»);
     - anything else → f.halt WITHOUT answering it (CLAUDE.md rule 6): Epsilon stays open with the message
       for debugging and the run stops — «να ενημερωνόμαστε άμεσα» (request to Lefteris: notify on halt)."""
    if not f.window_exists("TMessageForm"):
        return
    buttons = f.dialog_buttons("TMessageForm", store_as="unexpected_buttons")
    titles = [t for t in ("Σφάλμα", "Πληροφορία", "Επιβεβαίωση", "Διαγραφή", "Προσοχή")
              if f.window_exists("TMessageForm", t)]
    shot = f.screenshot(store_as="unexpected_shot")
    if titles == ["Πληροφορία"] and len(buttons) == 1 \
            and len(f.dialog_buttons("TMessageForm", "Πληροφορία", store_as="unexpected_info_buttons")) == 1:
        answer(f, "Πληροφορία", 0)
        f.log(f"ενημέρωση της Epsilon μετά {after}: «Πληροφορία» {buttons} → ΟΚ (φωτογραφία {shot})", level="WARNING")
        f.set_result("ενημέρωση Epsilon", f"μετά {after}: «Πληροφορία» → ΟΚ")
        return
    f.log(f"φωτογραφία: {shot}")                # no local path in the message (Lefteris v4, fail_shot)
    return f.halt(f"μήνυμα της Epsilon μετά {after}: «{titles[0] if titles else '?'}» {buttons} — ΔΕΝ απαντήθηκε, "
                  f"το Epsilon έμεινε ανοιχτό — ΔΕΙΤΕ ΤΟ SCREENSHOT")


def answer(f, title, button_index, timeout=10):
    """Press button `button_index` (0 = Ναι / ΟΚ, 1 = Όχι — as the ΚΑΔ question) of the TMessageForm
    `title`, by message; it must close."""
    f.connect_window(window_class="TMessageForm", window_title=title, timeout=5)
    f.click("", class_name="TButton", index=button_index, click_method="message")
    must_close(f, "TMessageForm", title, timeout)


def known_send_messages(f, where) -> bool:
    """The two messages «Αποστολή ΑΠΔ» brings, in any order (20–22/10/2026): «Πληροφορία» [ΟΚ] → ΟΚ and
    «Επιβεβαίωση» [Ναι, Όχι] (the ΑΠΔ's status → «Υποβλήθηκε»?) → Όχι (Petros 20/10/2026). When both
    are open «Πληροφορία» is on top: it first. True when one was answered."""
    if f.window_exists("TMessageForm", "Πληροφορία"):
        buttons = f.dialog_buttons("TMessageForm", "Πληροφορία", store_as="info_buttons")
        if len(buttons) == 1:
            answer(f, "Πληροφορία", 0)
            f.log(f"«Πληροφορία» {where} → ΟΚ")
            return True
    if f.window_exists("TMessageForm", "Επιβεβαίωση"):
        buttons = f.dialog_buttons("TMessageForm", "Επιβεβαίωση", store_as="status_buttons")
        if [norm(b) for b in buttons] == ["ναι", "οχι"]:
            shot = f.screenshot(store_as="status_question_shot")
            answer(f, "Επιβεβαίωση", 1)
            f.log(f"«Επιβεβαίωση» (κατάσταση ΑΠΔ → Υποβλήθηκε;) {where} → Όχι (φωτογραφία {shot})")
            return "status"                     # truthy; send_apd: this «Όχι» may cancel the new sending
    return False


def bar_click(f, title, counts, pick, what, key, parents=(False,)):
    """A label of the bottom bar of «Α.Π.Δ.»: only when the bar has one of `counts` labels; then label
    pick(n), ONCE (a second «Υπολογισμός» must never happen). 3 looks, 10 s apart.
    parents: which bars to look at, in turn — False = the bottommost TPanel itself; True = its PARENT.
    Page 2 (after «Υπολογισμός») has a small TPanel (330,823)–(663,850) INSIDE the bar for «Διαφορές ΑΠΔ -
    Περιόδων» | «Έλεγχος Υποκαταστημάτων»: it is the bottommost one, and the bar is its parent (ΚΟΤΖΙΝΟΣ
    09/11/2026: only those 2 labels were found, no «Επόμενη»)."""
    runs = []
    for attempt in range(ATTEMPTS):
        bring(f, APD, title, timeout=30)
        for parent in parents:
            runs = f.text_runs_in_control(APD_BAR, parent=parent, store_as=f"{key}_parent" if parent else key)
            if len(runs) in counts:
                f.click_text_in_control(APD_BAR, parent=parent, run=pick(len(runs)), capture="window")
                return runs
        if attempt < ATTEMPTS - 1:
            f.wait(DELAY)
    shot = f.screenshot(store_as=f"{key}_shot")
    problem(f, f"η κάτω μπάρα της «{title}» έχει {len(runs)} κείμενα {runs} αντί για {' ή '.join(map(str, counts))} — "
               f"δεν πατάω «{what}»", shot=shot)


def apd_period(f, year, month):
    """«Μισθολογική Περίοδος»: Από / Έως = the month (top / bottom combo, by message) and the year (the
    TNumEdit next to each). The same on «Νέος υπολογισμός» and «Αποθηκευμένα Στοιχεία»."""
    group = {"class_name": "TGroupBox", "target": "Μισθολογική Περίοδος"}
    for position in ("topmost", "bottommost"):
        retry(f, lambda f, position=position: f.select(
            MONTHS[month], target="TComboBox", class_name="TComboBox", within=group, prefer=position,
            fallback_index=month, select_method="message"))
        retry(f, lambda f, position=position: f.fill(
            year, target="TNumEdit", class_name="TNumEdit", within=group, prefer=position, fill_method="message"))


def open_apd(f, kind):
    """«Α.Π.Δ.» / «Α.Π.Δ. ΤΕΚΑ» from the menu search → «Επιλογή Εταιριών» → «Επιλογή» → TfAPD, and it
    must be the right one (title)."""
    k = APD_KINDS[kind]
    for attempt in range(ATTEMPTS):
        if f.window_exists(APD):
            break
        if not f.window_exists("TfChooseCompanies", "Επιλογή Εταιριών"):
            f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=60)
            f.navigate_menu(k["menu"], select=1, on_error="continue")     # the 1st row, any count
            if f.window_exists("TfChooseCompanies", "Επιλογή Εταιριών", timeout=30):
                f.log(f"«{k['menu']}»: 1η γραμμή")
        if f.window_exists("TfChooseCompanies", "Επιλογή Εταιριών"):
            f.connect_window(window_class="TfChooseCompanies", window_title="Επιλογή Εταιριών", timeout=5)
            f.click("Επιλογή", class_name="TBitBtn", click_method="message")
        if f.window_exists(APD, timeout=30):
            break
        if attempt < ATTEMPTS - 1:
            f.wait(DELAY)
    else:
        shot = f.screenshot(store_as=f"apd_menu_shot_{kind}")
        problem(f, f"δεν άνοιξε η «{k['title']}» από την αναζήτηση μενού «{k['menu']}» (1η γραμμή, {ATTEMPTS} "
                   f"φορές)", shot=shot)
    wrong = [x for x in k["not"] if f.window_exists(APD, x)]
    if not f.window_exists(APD, k["title"]) or wrong:
        shot = f.screenshot(store_as=f"apd_wrong_shot_{kind}")
        problem(f, f"άνοιξε άλλο παράθυρο ΑΠΔ αντί για «{k['title']}» ({', '.join(wrong) or 'άλλος τίτλος'})",
                shot=shot)
    bring(f, APD, k["title"], timeout=30)


def apd_settings(f, kind, title, year, month):
    """Νέος υπολογισμός · Τύπος Δήλωσης «01 Κανονική» (the leftmost combo, outside the group) · the
    month — then read back: the 3 combos = {01 Κανονική, month, month}, the 2 years = year."""
    bring(f, APD, title, timeout=30)
    retry(f, lambda f: f.set_checkbox("Νέος υπολογισμός Α.Π.Δ.", class_name="TRadioButton"))
    f.wait(0.5)
    retry(f, lambda f: f.select("01 Κανονική", target="TComboBox", class_name="TComboBox", prefer="leftmost",
                                select_method="message"))
    apd_period(f, year, month)
    combos = sorted(norm(f.get_control_text(class_name="TComboBox", index=i, store_as=f"apd_combo_{kind}_{i}"))
                    for i in (1, 2, 3))
    years = [f.get_control_text(class_name="TNumEdit", index=i, store_as=f"apd_year_{kind}_{i}") for i in (1, 2)]
    want = sorted([norm("01 Κανονική"), norm(MONTHS[month]), norm(MONTHS[month])])
    f.check(combos == want and all(str(y).strip() == str(year) for y in years),
            f"ρυθμίσεις «{title}»: μήνες/τύπος {combos}, έτη {years} αντί για 01 Κανονική, {MONTHS[month]} {year}")
    f.log(f"ρυθμίσεις: Νέος υπολογισμός · 01 Κανονική · {MONTHS[month]} {year} – {MONTHS[month]} {year}")


def case_review(f, kind, title, folder, year, month):
    """TuAPDCaseReview after «Υπολογισμός».
    «Ελλιπή στοιχεία» (tab) or the yellow «Προσοχή!» balloon = Epsilon leaves employees OUT of the ΑΠΔ
    (AVANTIS 27/10/2026; ΝΙΚΟΛΑΟΥ 10/2026 on 09/11: «Δεν έγινε υπολογισμός», «Μηδενική μισθοδοσία λόγω
    … λοχείας»). Petros 09/11/2026: «σταματάμε τη λειτουργία και το αναφέρουμε» → photo (also in the
    company's folder), the reason of the selected row, and stop WITHOUT OK.
    Otherwise «Μη επανυπολογισμένοι περίοδοι» → OK (Petros 22/10/2026: the funds already agree)."""
    f.wait(2)                                   # built in stages: tabs / balloon come a little later
    f.connect_window(window_class=REVIEW, timeout=10)
    # include_hidden=False: the «Ελλιπή στοιχεία» sheet EXISTS (hidden, no tab) also when only «Μη
    # επανυπολογισμένοι περίοδοι» is shown (13/11/2026 ΚΕΜΑΝΕ: stopped as «ΕΛΛΙΠΗ» with the company's
    # memo as «reason»). Real «Ελλιπή» (ΑΓΓΕΛΑΚΗΣ, ΠΛΟΥΚΙΔΗΣ, ΑΔΑΜΙΔΗΣ 13/11): the sheet is shown AND the
    # «Προσοχή!» balloon is up — either one stops.
    missing = f.control_exists("TRzTabSheet", "Ελλιπή στοιχεία", within={"class_name": REVIEW},
                               include_hidden=False, store_as=f"review_missing_{kind}")
    balloon = f.window_exists("TJvBalloonWindowEx")
    if missing or balloon:
        reason = f.get_control_text(class_name="TDBMemo", store_as=f"review_reason_{kind}", on_error="continue")
        photo = f"{folder}\\ΕΛΛΙΠΗ ΣΤΟΙΧΕΙΑ {APD_KINDS[kind]['prefix']} {month:02d} {year}.png"
        f.screenshot(photo, on_error="continue")
        shot = f.screenshot(store_as=f"review_shot_{kind}")
        why = " ".join(str(reason or "-").split())
        # Petros 23/11/2026: review.accept=yes (a button of the job, given ONLY for a company the accountant
        # approved) → OK and on, said in the result with the photo. The CSL01 check that follows still stops
        # if someone WITH contributions was left out (ΑΣΥΜΦΩΝΙΑ CSL01, nothing sent).
        if norm(f.ctx.get("review.accept") or "") == "yes":
            f.log(f"«Ελλιπή στοιχεία» ({why}) → OK με έγκριση χρήστη (review.accept=yes) — φωτογραφία {photo}",
                  level="WARNING")
            f.set_result(f"ελλιπή {kind}", f"συνέχισε παρά τα ελλιπή (αιτιολόγηση: «{why}») — με έγκριση χρήστη "
                                           f"(review.accept=yes) · φωτογραφία στον φάκελο της εταιρίας: "
                                           f"«{photo.split(chr(92))[-1]}»")
            f.click("OK", class_name="TBitBtn", click_method="message")
            must_close(f, REVIEW, "", 10)
            return
        # Petros 20/11/2026: STOP ALWAYS (no OK), but not called «ΕΛΛΙΠΗ» — ΝΙΚΟΛΑΟΥ 20/11: nothing was
        # missing, one employee had zero pay (λοχεία) while the company had records; ΜΑΚΡΟΠΟΥΛΟΣ: a real
        # missing ΑΜΙΚΑ. A person looks at the photo; never run again by itself.
        problem(f, f"ΕΡΓΑΖΟΜΕΝΟΙ ΕΚΤΟΣ ΑΠΔ — {title} {month:02d}/{year}: η Epsilon αφήνει εργαζόμενους εκτός ΑΠΔ "
                   f"(αιτιολόγηση: «{why}») — ΔΕΝ πατήθηκε OK, δεν στάλθηκε τίποτα. Φωτογραφία και στον φάκελο "
                   f"της εταιρίας: «{photo.split(chr(92))[-1]}»",
                known=True, shot=f"{photo} / {shot}")
    f.click("OK", class_name="TBitBtn", click_method="message")
    must_close(f, REVIEW, "", 10)
    f.log("«Μη επανυπολογισμένοι περίοδοι» → OK (τα ταμεία συμφωνούν)")


def calculate(f, kind, title, folder, year, month) -> str:
    """«Υπολογισμός» (once) and wait up to 240 s: "ok" (page 2: the «Διαφορές» button) or "saved"
    («Σφάλμα» [ΟΚ] = «Έχετε ήδη δημιουργήσει Α.Π.Δ.», left OPEN for the caller)."""
    bar_click(f, title, (2,), lambda n: 0, "Υπολογισμός", f"calc_bar_runs_{kind}")
    for _ in range(CALC_SECONDS):
        if f.window_exists(REVIEW, timeout=1):
            case_review(f, kind, title, folder, year, month)
            continue
        if f.window_exists("TMessageForm", "Σφάλμα"):
            if len(f.dialog_buttons("TMessageForm", "Σφάλμα", store_as=f"calc_error_buttons_{kind}")) == 1:
                return "saved"
            epsilon_message(f, "τον «Υπολογισμό»")
        epsilon_message(f, "τον «Υπολογισμό»")
        # page 2: «Διαφορές ΑΠΔ - Περιόδων» (and «Έλεγχος Υποκαταστημάτων») — part of the text, not exact
        # (09/11/2026: with exact it waited 240″ on a finished «Υπολογισμός»)
        if f.control_exists("TcxButton", "Διαφορές", within={"class_name": APD}, include_hidden=False, exact=False,
                            store_as=f"calc_done_{kind}"):
            f.close_font_capture(on_error="continue")
            return "ok"
    shot = f.screenshot(store_as=f"calc_timeout_shot_{kind}")
    problem(f, f"ο «Υπολογισμός» της «{title}» δεν τελείωσε σε {CALC_SECONDS}″", shot=shot)


def delete_saved(f, kind, title, year, month) -> tuple:
    """«Σφάλμα: Έχετε ήδη δημιουργήσει Α.Π.Δ.» → the PAD's way (21/10/2026, Petros: «ήδη δημιουργήσει →
    διαγραφή της αποθηκευμένης»): ΟΚ → «Αποθηκευμένα Στοιχεία Α.Π.Δ.» (without «Εμφάνιση όλων»: the
    «Υποβλήθηκε» ones stay hidden and are never deleted) → Από / Έως = the month, so the list has ONLY that
    month's ΑΠΔ (Petros 09/11/2026) → «Επόμενη» → its row → «Διαγραφή» → «Διαγραφή» Ναι → «Προηγούμενη».
    (deleted, labels): deleted False = «Διαγραφή» not pressed (labels = what the toolbar showed)."""
    shot = f.screenshot(store_as=f"saved_error_shot_{kind}")
    answer(f, "Σφάλμα", 0)
    f.log(f"«Σφάλμα» μετά τον «Υπολογισμό» (ήδη δημιουργήσει, φωτογραφία {shot}) → ΟΚ → αποθηκευμένη ΑΠΔ "
          f"{month:02d}/{year}")
    bring(f, APD, title, timeout=30)
    retry(f, lambda f: f.set_checkbox("Αποθηκευμένα Στοιχεία Α.Π.Δ.", class_name="TRadioButton"))
    f.wait(1.5)
    retry(f, lambda f: f.set_checkbox("Εμφάνιση όλων των Α.Π.Δ.", class_name="TCheckBox", checked=False))
    apd_period(f, year, month)
    bar_click(f, title, (2,), lambda n: 0, "Επόμενη (Αποθηκευμένα)", f"saved_bar_runs_{kind}")
    if not f.control_exists("TEpsilonDbGrid", within={"class_name": APD}, include_hidden=False, timeout=10,
                            store_as=f"saved_grid_{kind}"):
        epsilon_message(f, "την «Επόμενη» (Αποθηκευμένα)")
        shot = f.screenshot(store_as=f"saved_list_shot_{kind}")
        problem(f, "δεν άνοιξε η λίστα «Υπολογισμένα Δεδομένα»", shot=shot)
    epsilon_message(f, "την «Επόμενη» (Αποθηκευμένα)")
    bring(f, APD, title, timeout=30)
    f.click("", method="relative_coordinates", class_name="TEpsilonDbGrid", offset_x=60, offset_y=28)   # 1st row
    f.wait(1)
    labels = f.text_runs_in_control(APD_TOOLBAR, gap=5, min_width=30, store_as=f"saved_toolbar_runs_{kind}")
    f.log(f"μπάρα «Υπολογισμένα Δεδομένα»: {len(labels)} κουμπιά {labels}")
    deleted = len(labels) == TOOLBAR_LABELS
    if deleted:
        f.click_text_in_control(APD_TOOLBAR, gap=5, min_width=30, run=DELETE_LABEL, capture="window",
                                expect_hint="Διαγραφή")
        buttons = f.dialog_buttons("TMessageForm", "Διαγραφή", timeout=10, store_as=f"delete_buttons_{kind}")
        if [norm(b) for b in buttons] != ["ναι", "οχι"]:
            epsilon_message(f, "το κλικ στη «Διαγραφή»")
            shot = f.screenshot(store_as=f"delete_shot_{kind}")
            problem(f, f"μετά το κλικ στη «Διαγραφή» δεν βγήκε η ερώτηση «Διαγραφή» [Ναι / Όχι] (κουμπιά {buttons})",
                    shot=shot)
        shot = f.screenshot(store_as=f"delete_question_shot_{kind}")
        answer(f, "Διαγραφή", 0)
        f.wait(2)
        epsilon_message(f, "τη «Διαγραφή»")
        f.log(f"✔ διαγράφηκε η αποθηκευμένη {title} {month:02d}/{year} (φωτογραφία {shot})")
        f.set_result(f"διαγραφή αποθηκευμένης {kind}", f"{month:02d}/{year} · {shot}")
    else:
        # ΝΙΚΟΛΑΟΥ 21/10/2026: empty list, and the new «Υπολογισμός» went through
        f.log(f"«Υπολογισμένα Δεδομένα»: {len(labels)} κουμπιά αντί για {TOOLBAR_LABELS} — «Διαγραφή» ΔΕΝ "
              f"πατήθηκε ({month:02d}/{year})", level="WARNING")
    bar_click(f, title, (3,), lambda n: 0, "Προηγούμενη", f"back_bar_runs_{kind}")   # Προηγ. | Επόμενη | Ακύρωση
    if not f.control_exists("TRadioButton", "Νέος υπολογισμός Α.Π.Δ.", within={"class_name": APD},
                            include_hidden=False, timeout=10, store_as=f"back_page1_{kind}"):
        shot = f.screenshot(store_as=f"back_shot_{kind}")
        problem(f, f"δεν γύρισε στην 1η σελίδα της «{title}» μετά την «Προηγούμενη»", shot=shot)
    return deleted, labels


def to_export_page(f, kind, title):
    """«Επόμενη» (usually once, as the PAD up to twice) until «Εξαγωγή» shows. The bar there: «Προηγούμενη
    | Επόμενη | Ακύρωση» on the right, and on page 2 the buttons «Διαφορές ΑΠΔ - Περιόδων» | «Έλεγχος
    Υποκαταστημάτων» on its left (real controls, their captions count too: 5 labels — 09/11/2026)
    → «Επόμενη» = the one before the last; the result (the «Εξαγωγή» page) is checked."""
    def export_page(timeout=0):
        return f.control_exists(within={"class_name": APD}, include_hidden=False, exact=False, timeout=timeout,
                                store_as=f"export_page_{kind}", **EXPORT_PAGE)
    for i in (1, 2):
        if export_page():
            break
        bar_click(f, title, (3, 4, 5), lambda n: n - 2, f"Επόμενη #{i}", f"next_bar_runs_{kind}_{i}",
                  parents=(True, False))
        for _ in range(8):
            if export_page(timeout=1):
                break
            epsilon_message(f, f"την «Επόμενη» #{i}")
    if not export_page():
        epsilon_message(f, "την «Επόμενη»")
        shot = f.screenshot(store_as=f"export_page_shot_{kind}")
        problem(f, f"δεν φάνηκε η σελίδα με την «Εξαγωγή» της «{title}»", shot=shot)
    f.log("✔ σελίδα «Εξαγωγή Στοιχείων Α.Π.Δ.»")


def export_csl(f, kind, title) -> tuple:
    """«Εξαγωγή» → CSL01 in the folder its «Φάκελος Εξαγωγής» shows (usually …\\epsilon net\\hmisthos\\,
    AVANTIS …\\epsilon net — 22/10/2026). Epsilon writes it in stages (AVANTIS 27/10: read at 50688 of
    62450 bytes) → wait for «EOF» and a steady size. (path, lines)."""
    bring(f, APD, title, timeout=30)
    folder = None
    for i in range(1, 7):
        text = f.get_control_text(class_name="TEdit", index=i, store_as=f"apd_edit_{kind}_{i}", on_error="continue")
        if text and re.match(r"^[A-Za-z]:\\", str(text).strip()):
            folder = str(text).strip()
            break
    if not folder:
        shot = f.screenshot(store_as=f"export_folder_shot_{kind}")
        problem(f, f"δεν βρέθηκε ο «Φάκελος Εξαγωγής» της «{title}»", shot=shot)
    csl = folder.rstrip("\\") + "\\CSL01"
    f.log(f"φάκελος εξαγωγής: {folder}")
    started = f.now_epoch()
    f.click("Εξαγωγή", class_name="TBitBtn", click_method="message")
    f.wait_for_file(csl, timeout=60, newer_than=started - 1, last_line="EOF", encoding="cp1253", settle=2)
    epsilon_message(f, "την «Εξαγωγή»")
    lines = f.read_text_file(csl, encoding="cp1253", store_as=f"csl_{kind}")
    f.log(f"✔ Εξαγωγή: {csl} ({len(lines)} γραμμές)")
    return csl, lines


def read_csl(lines) -> dict:
    """CSL01 (cp1253, fixed width). Header (record 1): …MMYYYY + days(8) + pay(12) + due(12) + date(8).
    Record 2 = employee (ΑΦΜ = last 9). Record 3 = insurance line: days [49:52], pay [62:72], due [123:134]
    (26/10/2026, ΕΦΚΑ + ΤΕΚΑ). [52:62] is the day wage (≠0 for day workers: ΝΙΚΟΛΑΟΥ 03/11/2026 «0000007706»
    = 77,06); with [60:72] its last 2 digits stuck in front of the pay. [112:123] = subsidy. None = no header."""
    head = re.search(r"(\d{6})(\d{8})(\d{12})(\d{12})(\d{8})\s*$", lines[0]) if lines else None
    if not head or not lines[0].startswith("1"):
        return None

    def cents(s):
        return (Decimal(int(s)) / 100).quantize(CENT)
    r3 = [x for x in lines if x.startswith("3")]
    return {"period": head.group(1),
            "header": (int(head.group(2)), cents(head.group(3)), cents(head.group(4))),
            "lines": (sum(int(x[49:52]) for x in r3), sum((cents(x[62:72]) for x in r3), ZERO),
                      sum((cents(x[123:134]) for x in r3), ZERO)),
            "afm": {x.rstrip()[-9:]: x[21:71].strip() for x in lines if x.startswith("2")}}


def payroll_people(rows, column, skip=(), only=()) -> dict:
    """Employees of a Μισθοδοτική's Excel with contributions in `column` → {ΑΦΜ: (name, amount)}. Without
    the `skip` periods; with `only`, ONLY those (the 2nd month's gift)."""
    header = next((i for i, r in enumerate(rows) if any(norm(c) == "περιοδος" for c in r)), None)
    if header is None:
        return {}
    cols = {norm(h): i for i, h in enumerate(rows[header])}
    afm_col = next((i for i, h in enumerate(rows[header]) if norm(h).replace(".", "") == "αφμ"), None)

    def cell(r, name, i=None):
        i = cols.get(name) if i is None else i
        return r[i] if i is not None and i < len(r) and r[i] is not None else ""
    skip, only = [norm(x) for x in skip], [norm(x) for x in only]
    out = {}
    for r in rows[header + 1:]:
        p = norm(cell(r, "περιοδος"))
        if not str(cell(r, "εταιρεια")).strip() or p in skip or (only and p not in only):
            continue
        value = amount(cell(r, norm(column)))
        if value > 0:
            afm = str(cell(r, "", afm_col)).strip().split(".")[0].zfill(9)
            name, total = out.get(afm, (f"{cell(r, 'επωνυμο')} {cell(r, 'ονομα')}".strip(), ZERO))
            out[afm] = (name, total + value)
    return out


def csl_problems(c, kind, year, month, due_pdf, people, others) -> list:
    """Before «Αποστολή ΑΠΔ» (Petros 26/10/2026, after TRAVEL SENSE: an ΑΠΔ without one employee): what
    will be uploaded = what we have. ΤΕΚΑ may miss ONLY people of another auxiliary fund (e.g. ΤΕΑΥΦΕ):
    accepted only if their contributions add up exactly to the other funds of the PDF."""
    if c is None:
        return ["το CSL01 δεν έχει την αναμενόμενη κεφαλίδα"]
    problems = []
    if c["period"] != f"{month:02d}{year}":
        problems.append(f"περίοδος {c['period']} αντί για {month:02d}{year}")
    if c["header"] != c["lines"]:
        problems.append(f"κεφαλίδα (ημέρες {c['header'][0]} · αποδοχές {c['header'][1]} · καταβλητέες "
                        f"{c['header'][2]}) ≠ γραμμές (ημέρες {c['lines'][0]} · αποδοχές {c['lines'][1]} · "
                        f"καταβλητέες {c['lines'][2]})")
    if c["header"][2] != due_pdf:
        problems.append(f"καταβλητέες {c['header'][2]} αντί για {due_pdf} (PDF εισφορών {kind})")
    missing = {a: v for a, v in people.items() if a not in c["afm"]}
    if missing:
        names = ", ".join(f"{n} ({a})" for a, (n, _) in missing.items())
        gone = sum((v[1] for v in missing.values()), ZERO)
        other_sum = sum(others.values(), ZERO)
        if not (kind == "ΤΕΚΑ" and others and gone == other_sum):
            problems.append(f"λείπουν από το CSL01: {names} — η Epsilon τους άφησε έξω από την ΑΠΔ")
    return problems


def submit_page(f, key, timeout=2) -> list:
    return f.uia_page(timeout=timeout, store_as=key, on_error="continue", **SUBMIT_WIN) or []


def submit_rows(page) -> list:
    """The ΕΦΚΑ list of submissions: each row starts with its number (a link of digits only).
    → [{"number", "texts", "actions"}] in page order (as fmy_python\\ΑΠΔ.py submit_rows)."""
    rows, cur = [], None
    for el in page:
        kind, name = el.get("type"), str(el.get("name") or "").strip()
        if kind == "Hyperlink" and name.isdigit() and len(name) >= 6:
            cur = {"number": name, "texts": [], "actions": []}
            rows.append(cur)
        elif cur is not None:
            if norm(name) == norm("Περίοδος Υποβολής") or kind == "Button":
                break                                   # end of the table
            if kind == "Hyperlink":
                cur["actions"].append(name)
            elif name not in cur["texts"]:
                cur["texts"].append(name)
    return rows


def has_period(texts, year, month) -> bool:
    joined = " ".join(str(t) for t in texts)
    return any(re.search(rf"(?<!\d){re.escape(p)}(?!\d)", joined) for p in (f"{month}/{year}", f"{month:02d}/{year}"))


def page_has(page, text, kinds=None) -> bool:
    return any(norm(el.get("name")) == norm(text) and (kinds is None or el.get("type") in kinds) for el in page)


def wait_page(f, cond, seconds, what, key):
    """Poll the EpsilonSubmit page (every ~2 s) until cond(page) is true; then the page, else fail.
    The sending's late «Πληροφορία» [ΟΚ] / «Επιβεβαίωση» are answered meanwhile: they block the page
    (13/11/2026 ΠΕΤΣΑΣ ΕΕ: the pending's number was clicked under an open «Πληροφορία» → no popup)."""
    for _ in range(max(1, seconds // 2)):
        if known_send_messages(f, f"στην αναμονή για {what}") or late_known(f, f"στην αναμονή για {what}"):
            continue
        page = submit_page(f, key)
        if cond(page):
            return page
        f.wait(1)
    shot = f.screenshot(store_as=f"{key}_shot")
    problem(f, f"EpsilonSubmit: δεν φάνηκε {what} σε {seconds}″", shot=shot)


def row_click(f, action, number, timeout=30):
    """A link of ONE row of the list (the row whose number is `number`)."""
    f.uia_click(action, control_type="Hyperlink", exact=True, row_contains=[rf"^{number}$"], timeout=timeout,
                **ROW, **SUBMIT_WIN)


def send_apd(f, kind, title, year, month) -> dict:
    """«Αποστολή ΑΠΔ» → EpsilonSubmit uploads the CSL01. Success = «ανέβηκε επιτυχώς» on its page and 3 s
    without an Epsilon message. Also seen (fmy_python\\ΑΠΔ.py, 20–23/10/2026):
     - the page «Επιλογή Αρχείου προς Υποβολή» WITHOUT the file (1618, ACTION, EONIAN, ΑΓΓΕΛΑΚΗΣ) → after
       60 s: close EpsilonSubmit / «Υποβολή ΑΠΔ» and «Αποστολή ΑΠΔ» again, up to SEND_TRIES;
     - the list instead: an ΑΠΔ of the month already PENDING in ΕΦΚΑ (801646253) → that one is submitted;
     - «Σφάλμα» [ΟΚ] with «Έχετε υποβάλλει ήδη … MM/YYYY» (ΨΑΛΛΙΔΑ) → already submitted.
    {state: uploaded | pending | already, row?}."""
    unread = 0                                  # attempts in a row with EpsilonSubmit never readable
    for attempt in range(1, SEND_TRIES + 1):
        # Petros 14/11/2026: give ΕΦΚΑ time — 10″ before the 1st «Αποστολή ΑΠΔ», 60″ before each retry
        # («Open period button not found» ~1 in 5 sends on 13–14/11; retries 1–2″ apart failed in a row)
        # the pause ANSWERS the late «κατάσταση → Υποβλήθηκε;» (→ Όχι) of the killed EpsilonSubmit: it
        # came during the 60″ and «Αποστολή ΑΠΔ» was pressed under it (Petros 14/11/2026 «εδώ πάτα ΌΧΙ»)
        for _ in range(SEND_PAUSE if attempt == 1 else RETRY_PAUSE):
            if not known_send_messages(f, "στην αναμονή πριν την «Αποστολή ΑΠΔ»"):
                f.wait(1)
        bring(f, APD, title, timeout=30)
        f.click("Αποστολή ΑΠΔ", class_name="TBitBtn", click_method="message")
        f.log(f"«Αποστολή ΑΠΔ» — προσπάθεια {attempt}/{SEND_TRIES}")
        empty_since, ok_since, reclick_at, reclicked, seen_submit = None, None, None, None, False
        t0 = f.now_epoch()                     # real seconds: one loop round takes ~1–2 s (14/11/2026)

        def on_known(answered, now):
            # 14/11/2026 (ΔΑΦΝΗΣ ΟΕ, ΑΦΕΝΤΑΚΗ, ΒΑΡΔΑΚΩΣΤΑΣ, ΓΕΩΡΓΑΝΤΑΣ): on a RETRY Epsilon first asks «κατάσταση →
            # Υποβλήθηκε;» about the killed sending; the «Όχι» sometimes also cancels the new one (no EpsilonSubmit,
            # 3 min lost) → if EpsilonSubmit has not shown up 60″ after it, «Αποστολή ΑΠΔ» once more (once per
            # attempt). 60″: after the same «Όχι» EpsilonSubmit did open ~20″ later (ΒΑΡΔΑΚΩΣΤΑΣ 14/11 11:08) —
            # never a 2nd click while one may still be starting. Seen by EITHER message check (the 1st try of
            # this missed it: the 2nd check caught the «Όχι» → no re-click, ΔΑΦΝΗΣ 12:45).
            nonlocal reclick_at
            if answered == "status" and not seen_submit and reclicked is None:
                reclick_at = now + 60

        for second in range(UPLOAD_SECONDS):
            now = f.now_epoch() - t0
            if now > UPLOAD_SECONDS:
                break
            f.close_font_capture(on_error="continue")
            answered = known_send_messages(f, "μετά την «Αποστολή ΑΠΔ»")
            if answered:
                on_known(answered, now)
                ok_since = None
                continue
            page = submit_page(f, f"send_page_{kind}", timeout=1)
            seen_submit = seen_submit or bool(page)
            now = f.now_epoch() - t0
            if not seen_submit and reclick_at is None and now >= (reclicked or 0) + NO_SUBMIT_SECONDS:
                f.log(f"το EpsilonSubmit δεν άνοιξε σε {NO_SUBMIT_SECONDS}″ (προσπάθεια {attempt})", level="WARNING")
                break
            if reclick_at is not None and now >= reclick_at and not seen_submit:
                reclick_at, reclicked = None, now
                bring(f, APD, title, timeout=30)
                f.click("Αποστολή ΑΠΔ", class_name="TBitBtn", click_method="message")
                f.log("«Αποστολή ΑΠΔ» ξανά: το «Όχι» στην «κατάσταση → Υποβλήθηκε;» ακύρωσε την αποστολή", level="WARNING")
                continue
            if f.window_exists("TMessageForm", "Σφάλμα"):
                sent = any("εχετε υποβαλλει ηδη" in norm(el.get("name")) and has_period([el.get("name")], year, month)
                           for el in page)
                if sent and len(f.dialog_buttons("TMessageForm", "Σφάλμα", store_as="sent_error_buttons")) == 1:
                    shot = f.screenshot(store_as=f"already_sent_shot_{kind}")
                    answer(f, "Σφάλμα", 0)
                    for _ in range(5):                  # «κατάσταση → Υποβλήθηκε;» → Όχι
                        if not known_send_messages(f, "μετά το «Έχετε υποβάλλει ήδη»"):
                            f.wait(1)
                    return {"state": "already", "shot": shot}
                # «Open period button not found (j_idt210) - openPeriodForm» [ΟΚ] on «Δημιουργία Περιόδου
                # Υποβολής»: EpsilonSubmit could not open the period, NOTHING was uploaded (13/11/2026: ΠΕΤΣΑΣ
                # ΕΕ, ΑΛΙΑΚΙΟΖΟΓΛΟΥ, ΑΦΕΝΤΑΚΗ, ΒΑΡΔΑΚΩΣΤΑΣ; a 2nd try passed). Petros 13/11/2026: «να πατάει ΟΚ
                # και να ξαναπροσπαθεί» → photo, ΟΚ, close EpsilonSubmit, «Αποστολή ΑΠΔ» again (SEND_TRIES)
                no_period = any("open period button not found" in norm(el.get("name")) for el in page) or (
                    page_has(page, "Δημιουργία Περιόδου Υποβολής")
                    and not any(has_period(r["texts"], year, month) for r in submit_rows(page)))
                if no_period and len(f.dialog_buttons("TMessageForm", "Σφάλμα", store_as="period_error_buttons")) == 1:
                    shot = f.screenshot(store_as=f"open_period_error_shot_{kind}_{attempt}")
                    answer(f, "Σφάλμα", 0)
                    f.log(f"«Open period button not found» (προσπάθεια {attempt}) → ΟΚ, κλείσιμο και ξανά "
                          f"(φωτογραφία {shot})", level="WARNING")
                    break
            # A known message can appear WHILE the page is read (~1 s): look again before calling it
            # unexpected (13/11/2026 Α & Μ ΔΕΛΗΓΙΩΡΓΗ: «Πληροφορία» «Ολοκληρώθηκε η αυτόματη συμπλήρωση
            # ΑΠΔ…» [ΟΚ] came during uia_page → the flow stopped although the CSL01 had been uploaded).
            answered = known_send_messages(f, "μετά την «Αποστολή ΑΠΔ»") or late_known(
                f, "μετά την «Αποστολή ΑΠΔ»", error_handled=True)
            if answered:
                on_known(answered, f.now_epoch() - t0)
                ok_since = None
                continue
            names = [norm(el.get("name")) for el in page]
            # ΕΦΚΑ codes of the company wrong: «Αποτυχία στην Αυθεντικοποίηση του χρήστη! Η σύνδεσή σας δεν
            # ήταν δυνατή.» (14/11/2026 ΓΕΩΡΓΟΥΛΗΣ — as his TAXISnet codes in ΦΜΥ 09/11) → stop at once
            # (Petros: as an error), nothing was uploaded; never retried, never typed by the flow
            if any("αποτυχια στην αυθεντικοποιηση" in n for n in names):
                shot = f.screenshot(store_as=f"bad_login_shot_{kind}")
                problem(f, "ΛΑΘΟΣ ΚΩΔΙΚΟΙ ΕΦΚΑ — «Αποτυχία στην Αυθεντικοποίηση του χρήστη» στο EpsilonSubmit: οι "
                           "κωδικοί ΕΦΚΑ της εταιρίας στην Epsilon πρέπει να διορθωθούν. Δεν στάλθηκε τίποτα",
                        known=True, shot=shot)
            if any("ανεβηκε επιτυχως" in n for n in names):      # norm keeps the final «ς»
                ok_since = second if ok_since is None else ok_since
                if second - ok_since >= 3:
                    return {"state": "uploaded"}
                continue
            pending = [r for r in submit_rows(page) if has_period(r["texts"], year, month)
                       and any(norm(t) in (norm("ΠΡΟΣ ΕΛΕΓΧΟ ΥΠΟΒΟΛΗΣ"), norm("ΜΕ ΔΥΝΑΤΟΤΗΤΑ ΥΠΟΒΟΛΗΣ"))
                               for t in r["texts"])]
            if pending:
                return {"state": "pending", "row": pending[0]}
            if page_has(page, "Επιλογή Αρχείου προς Υποβολή"):
                empty_since = second if empty_since is None else empty_since
                if second - empty_since >= EMPTY_PAGE_SECONDS:
                    break
        # ACTION ADVANCED 17/11/2026: «Αποτυχία στην Αυθεντικοποίηση», but the EpsilonSubmit window could not
        # be read (empty title → uia_page «not found»), so the check above never saw it and 10 «Αποστολή ΑΠΔ»
        # took ~10′. Petros 17/11/2026 «μέχρι 2 φορές»: EpsilonSubmit never readable in 2 attempts → stop as
        # (probably) wrong ΕΦΚΑ codes, never run again by itself. («Open period button not found» shows
        # EpsilonSubmit's page, so those retries still go up to SEND_TRIES.)
        unread = unread + 1 if not seen_submit else 0
        if unread >= UNREAD_TRIES:
            shot = f.screenshot(store_as=f"send_unread_shot_{kind}")
            problem(f, f"ΛΑΘΟΣ ΚΩΔΙΚΟΙ ΕΦΚΑ; — το EpsilonSubmit δεν φάνηκε σε {unread} «Αποστολή ΑΠΔ» της «{title}» "
                       f"(έτσι ήταν η «Αποτυχία στην Αυθεντικοποίηση» της ACTION 17/11) — ελέγξτε τους κωδικούς ΕΦΚΑ "
                       f"της εταιρίας στην Epsilon. Δεν στάλθηκε τίποτα", known=True, shot=shot)
        if attempt == SEND_TRIES:
            shot = f.screenshot(store_as=f"send_shot_{kind}")
            problem(f, f"το CSL01 της «{title}» δεν ανέβηκε στο EpsilonSubmit μετά από {SEND_TRIES} «Αποστολή ΑΠΔ»",
                    shot=shot)
        f.log(f"το CSL01 δεν φορτώθηκε στο EpsilonSubmit (προσπάθεια {attempt}) — κλείσιμο και ξανά", level="WARNING")
        close_submit(f)


def close_submit(f):
    """EpsilonSubmit (its own program, nothing in progress) and «Υποβολή ΑΠΔ» (TfESubmitDocs) close; the
    late «Πληροφορία» / «Επιβεβαίωση» of the sending are answered; «Α.Π.Δ.» stays."""
    f.kill_process(["EpsilonSubmit.exe"], timeout=10, ignore_missing=True, on_error="continue")
    for _ in range(15):
        if known_send_messages(f, "στο κλείσιμο του EpsilonSubmit"):
            continue
        if f.window_exists("TfESubmitDocs"):
            f.connect_window(window_class="TfESubmitDocs", timeout=5)
            f.close_app()
            f.wait(1)
            continue
        break
    epsilon_message(f, "το κλείσιμο του EpsilonSubmit")


def dump_page(f, page, what):
    """Everything an EpsilonSubmit page shows (name = value), to learn what can be checked."""
    f.log(f"σελίδα «{what}» ({len(page)} στοιχεία):")
    for el in page:
        name_, value_ = str(el.get("name") or "").strip(), str(el.get("value") or "").strip()
        if name_ or value_:
            f.log(f"   [{el.get('type')}] {name_}" + (f" = {value_}" if value_ else ""))


def settle_messages(f, where, seconds=8):
    """The sending's messages come late (up to ~30 s after the page changed): answer the known ones
    («Πληροφορία» → ΟΚ, «Επιβεβαίωση» κατάστασης → Όχι) until `seconds` pass without one; anything else
    → photo and stop (epsilon_message)."""
    quiet = 0
    while quiet < seconds:
        if known_send_messages(f, where) or late_known(f, where):
            quiet = 0
            continue
        f.wait(1)
        quiet += 1


def late_known(f, where, error_handled=False) -> bool:
    """A TMessageForm that is not (yet) a known one: Epsilon gives it its title a moment later (13/11/2026
    ΠΕΤΣΑΣ ΕΕ: «Πληροφορία» → False, TMessageForm → True in the same second, then stopped as unexpected).
    Look again for 3 s; still unknown → photo and stop (epsilon_message)."""
    if not f.window_exists("TMessageForm"):
        return False
    for _ in range(3):
        f.wait(1)
        known = known_send_messages(f, where)
        if known:
            return known                        # "status" passes through (send_apd re-clicks on it)
    if error_handled and f.window_exists("TMessageForm", "Σφάλμα"):
        return True                                 # the caller's «Σφάλμα» handling looks at it next
    epsilon_message(f, where)
    return False


DETAILS = "Λοιπά Στοιχεία Υποβολής"
# «μερών Ασφάλισης»: the page writes «Hμερών» with a LATIN H (14/11/2026 ΠΕΤΣΑΣ ΕΕ)
DETAIL_LABELS = (("days", "μερών Ασφάλισης"), ("pay", "Αποδοχών"), ("total", "Σύνολο Εισφορών"))


def submission_details(f, kind, number) -> dict:
    """Click the submission's number in the ΕΦΚΑ list → the popup «Λοιπά Στοιχεία Υποβολής» (13/11/2026,
    ΔΕΛΗΓΙΩΡΓΗ 83748793: «14 Ημερών Ασφάλισης: 54 · 15 Αποδοχών: 2.327,31 · 16 Σύνολο Εισφορών: 818,28»)
    → {days, pay, total} (None = not read); the popup closes with «Κλείσιμο»."""
    settle_messages(f, "πριν τον αριθμό της εκκρεμούς")
    f.uia_click(number, control_type="Hyperlink", exact=True, timeout=30, **SUBMIT_WIN)
    page = wait_page(f, lambda p: any(DETAILS in str(el.get("name") or "") for el in p), 30,
                     f"το «{DETAILS}»", f"details_page_{kind}")
    dump_page(f, page, DETAILS)
    out = {}
    start = next(i for i, el in enumerate(page) if DETAILS in str(el.get("name") or ""))
    # Each field: [DataItem] «14 Hμερών Ασφάλισης:» · [Text] «14» · [Text] label · [Edit] «14 Hμερών…» whose
    # VALUE is the number. Only that Edit's value counts: the «14»/«15»/«16» texts are row numbers
    # (14/11/2026 ΠΕΤΣΑΣ ΕΕ: read as amounts «15.00» / «16.00» while the fields were still empty)
    for key, label in DETAIL_LABELS:               # only inside the popup (after its title)
        edit = next((el for i, el in enumerate(page) if i > start and el.get("type") == "Edit"
                     and label in str(el.get("name") or "")), None)
        m = re.fullmatch(r"\s*([\d.]+(?:,\d+)?)\s*", str((edit or {}).get("value") or ""))
        if m:
            out[key] = int(m.group(1).replace(".", "")) if key == "days" else amount(m.group(1))
    f.uia_click("Κλείσιμο", control_type="Button", exact=True, timeout=15, on_error="continue", **SUBMIT_WIN)
    f.log(f"ΕΦΚΑ αρ. {number}: {out}")
    return out


def same_as_csl(f, kind, number, csl, contrib):
    """A PENDING ΑΠΔ in ΕΦΚΑ is submitted ONLY if it is what we would upload now (Petros 13/11/2026): days,
    pay and contributions = the new CSL01 («Σύνολο Εισφορών» = καταβλητέες, or + επιδοτήσεις as on the
    receipt). Otherwise stop: deleting it («Ακύρωση Περιόδου Υποβολής»?) is not known yet."""
    d = submission_details(f, kind, number)
    if not csl:
        problem(f, f"εκκρεμής ΑΠΔ {number}: δεν υπάρχει CSL01 για σύγκριση — δεν υποβλήθηκε")
    days, pay, due = csl["header"]
    totals = (due, due + contrib["επιδοτήσεις"].get(kind, ZERO))
    ok = d.get("days") == days and d.get("pay") == pay and d.get("total") in totals
    text = (f"ΕΦΚΑ: ημέρες {d.get('days')} · αποδοχές {d.get('pay')} · εισφορές {d.get('total')} — CSL01: "
            f"ημέρες {days} · αποδοχές {pay} · καταβλητέες {due}")
    if len(d) < 3:
        shot = f.screenshot(store_as=f"pending_unread_shot_{kind}")
        problem(f, f"η εκκρεμής ΑΠΔ {number}: τα «{DETAILS}» δεν διαβάστηκαν ({text}) — ΔΕΝ υποβλήθηκε "
                   f"(η σελίδα είναι στο log)", shot=shot)
    if not ok:
        shot = f.screenshot(store_as=f"pending_differs_shot_{kind}")
        problem(f, f"η εκκρεμής ΑΠΔ {number} στον ΕΦΚΑ ΔΙΑΦΕΡΕΙ από το νέο CSL01 ({text}) — ΔΕΝ υποβλήθηκε· "
                   f"να ακυρωθεί στον ΕΦΚΑ", known=True, shot=shot)
    f.log(f"✔ εκκρεμής {number} = νέο CSL01 ({text})")
    f.set_result(f"εκκρεμής {kind}", f"ίδια με CSL01 ({text})")


def submit_apd(f, kind, title, year, month, contrib, folder, name, downloads, afm, csl=None,
               stop_final="") -> str:
    """From the uploaded / pending ΑΠΔ to «Υποβολή & Έκδοση Αποδεικτικού» (once). The result text."""
    sent = send_apd(f, kind, title, year, month)
    if sent["state"] == "already":
        return f"ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ στον ΕΦΚΑ («Έχετε υποβάλλει ήδη», φωτογραφία {sent['shot']})"
    if sent["state"] == "uploaded":
        f.log("✔ το CSL01 ανέβηκε επιτυχώς")
        # The «Πληροφορία» «Ολοκληρώθηκε η αυτόματη συμπλήρωση ΑΠΔ…» can come AFTER «ανέβηκε επιτυχώς» (13/11/2026
        # ΠΕΤΣΑΣ ΕΕ: 30 s later): it blocks «Επιστροφή» → answer it (ΟΚ) and press «Επιστροφή» again while
        # the upload page is still there, until the list shows
        page, last_back = None, -99
        for second in range(0, 90, 2):
            if known_send_messages(f, "μετά το ανέβασμα") or late_known(f, "μετά το ανέβασμα του CSL01"):
                continue
            page = submit_page(f, f"list_page_{kind}")
            if any(r["actions"] for r in submit_rows(page)):
                break
            if second - last_back >= 10 and page_has(page, "Επιστροφή", ("Button",)):
                f.uia_click("Επιστροφή", control_type="Button", exact=True, timeout=10, on_error="continue",
                            **SUBMIT_WIN)
                last_back = second
            f.wait(1)
        else:
            shot = f.screenshot(store_as=f"list_shot_{kind}")
            problem(f, "EpsilonSubmit: δεν φάνηκε η λίστα υποβολών σε 90″ μετά το ανέβασμα", shot=shot)
        rows = submit_rows(page)
        new = rows[0]
        if "Έλεγχος Υποβολής" not in new["actions"] or not has_period(new["texts"], year, month):
            shot = f.screenshot(store_as=f"list_shot_{kind}")
            problem(f, f"η 1η γραμμή της λίστας ({new['number']}: {' | '.join(new['texts'][:4])}) δεν είναι η ΑΠΔ "
                       f"{month}/{year} με «Έλεγχος Υποβολής» — δεν συνεχίζω", shot=shot)
    else:
        rows = submit_rows(submit_page(f, f"list_page_{kind}"))
        new = sent["row"]
        f.log(f"εκκρεμής {title} {month:02d}/{year} αρ. {new['number']} στον ΕΦΚΑ → υποβάλλεται αυτή", level="WARNING")
        f.set_result(f"ΑΠΔ {kind} εκκρεμής", new["number"])
    number = new["number"]
    for r in rows[:5]:
        f.log(f"   γραμμή {r['number']}: {' | '.join(r['texts'])} · ενέργειες: {', '.join(r['actions'])}")
    if stop_final == "list":
        shot = f.screenshot(store_as=f"list_stop_shot_{kind}")
        f.halt(f"ΣΤΟΠ στη λίστα του ΕΦΚΑ (stop.before_final=list) — {title} {month:02d}/{year} αρ. {number}: "
               f"τίποτα δεν πατήθηκε, όλα ανοιχτά. Φωτογραφία {shot}")
    done = [r for r in rows if r["number"] != number and "Επανέκδοση Αποδεικτικού" in r["actions"]
            and has_period(r["texts"], year, month)]
    if done:
        shot = f.screenshot(store_as=f"already_submitted_shot_{kind}")
        return f"ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ στον ΕΦΚΑ αρ. {done[0]['number']} — η νέα ({number}) ΔΕΝ υποβλήθηκε (φωτογραφία {shot})"
    # «Έλεγχος Υποβολής» → OK → until «ΜΕ ΔΥΝΑΤΟΤΗΤΑ ΥΠΟΒΟΛΗΣ» (the row gets «Υποβολή»)
    def row_of(page):
        return next((r for r in submit_rows(page) if r["number"] == number), None)
    row = row_of(submit_page(f, f"row_page_{kind}")) or new
    if "Υποβολή" not in row["actions"] and "Έλεγχος Υποβολής" in row["actions"]:
        row_click(f, "Έλεγχος Υποβολής", number)
        # 13/11/2026 (ΔΕΛΗΓΙΩΡΓΗ, pending ΑΠΔ): instead of the page's «OK» an Epsilon «Πληροφορία» came
        # and the flow failed on the missing OK (Epsilon closed, message unread) → wait for either; a
        # message → photo and HALT (everything stays open for a person to read it), never answered
        for _ in range(30):
            if known_send_messages(f, "μετά το «Έλεγχος Υποβολής»"):   # the sending's late «Πληροφορία» → ΟΚ
                continue
            if f.window_exists("TMessageForm"):        # its title may come a moment later (late_known)
                f.wait(2)
                if known_send_messages(f, "μετά το «Έλεγχος Υποβολής»"):
                    continue
            if f.window_exists("TMessageForm"):
                buttons = f.dialog_buttons("TMessageForm", store_as="check_message_buttons")
                shot = f.screenshot(store_as=f"check_message_shot_{kind}")
                f.halt(f"μήνυμα της Epsilon μετά το «Έλεγχος Υποβολής» ({number}): {buttons} — ΔΕΝ απαντήθηκε, "
                       f"όλα μένουν ανοιχτά. Φωτογραφία {shot}")
            if page_has(submit_page(f, f"check_ok_page_{kind}", timeout=1), "OK", ("Button",)):
                break
        f.uia_click("OK", control_type="Button", exact=True, timeout=30, **SUBMIT_WIN)
    last_refresh = 0
    for second in range(0, CHECK_SECONDS, 2):
        row = row_of(submit_page(f, f"check_page_{kind}"))
        bad = next((t for t in (row or {}).get("texts", [])
                    if any(w in norm(t) for w in ("σφαλμ", "απορρ", "λαθ"))), None)
        if bad:
            # Petros 17/11/2026: only stop and say so — the errors themselves are NOT read
            shot = f.screenshot(store_as=f"check_errors_shot_{kind}")
            earlier = " (εκκρεμής από προηγούμενη αποστολή)" if sent["state"] == "pending" else ""
            problem(f, f"ΣΤΑΜΑΤΗΣΕ — ΛΑΘΗ ΣΤΟΝ ΕΛΕΓΧΟ ΠΡΙΝ ΤΗΝ ΥΠΟΒΟΛΗ: η ΑΠΔ {kind} {month:02d}/{year} "
                       f"(αρ. {number}){earlier} βγήκε «{bad}» — ΔΕΝ υποβλήθηκε", known=True, shot=shot)
        if row and "Υποβολή" in row["actions"]:
            break
        if row and "Ενημέρωση Προόδου Ελέγχου Υποβολής" in row["actions"] and second - last_refresh >= 6:
            row_click(f, "Ενημέρωση Προόδου Ελέγχου Υποβολής", number)
            last_refresh = second
            f.uia_click("OK", control_type="Button", exact=True, timeout=5, on_error="continue", **SUBMIT_WIN)
        f.wait(1)
    else:
        shot = f.screenshot(store_as=f"check_timeout_shot_{kind}")
        problem(f, f"η υποβολή {number} δεν πέρασε τον έλεγχο σε {CHECK_SECONDS}″", shot=shot)
    f.log(f"✔ έλεγχος υποβολής {number}: ΜΕ ΔΥΝΑΤΟΤΗΤΑ ΥΠΟΒΟΛΗΣ")
    # a PENDING one: its «Λοιπά Στοιχεία Υποβολής» AFTER the check (Petros 13/11/2026: «πρώτα έλεγχο
    # υποβολής και μετά το νούμερο») and only if = the new CSL01 it goes on to «Υποβολή»
    if sent["state"] == "pending":
        same_as_csl(f, kind, number, csl, contrib)

    # «Υποβολή» → «Συνέχεια» (its «Περίοδος Υποβολής» = the month) → «Υποβολή & Έκδοση Αποδεικτικού»
    row_click(f, "Υποβολή", number)
    page = wait_page(f, lambda p: page_has(p, "Συνέχεια", ("Button",)), 30, "η «Συνέχεια»", f"continue_page_{kind}")
    dump_page(f, page, "Υποβολή (πριν τη Συνέχεια)")
    for el in page:
        if el.get("type") == "Edit" and norm(el.get("name")).startswith(norm("Περίοδος Υποβολής")):
            value = str(el.get("value") or "")
            f.log(f"σελίδα υποβολής — Περίοδος Υποβολής: «{value}»")
            f.check(not value or has_period([value], year, month),
                    f"η σελίδα υποβολής γράφει περίοδο «{value}» αντί για {month}/{year}")
    f.uia_click("Συνέχεια", control_type="Button", exact=True, timeout=30, **SUBMIT_WIN)
    page = wait_page(f, lambda p: page_has(p, FINAL, ("Button",)), 30, f"το «{FINAL}»", f"final_page_{kind}")
    f.log("τελική σελίδα πριν την οριστική υποβολή: " + " | ".join(
        f"{el.get('name')}={el.get('value')}" if el.get("value") else str(el.get("name"))
        for el in page if el.get("type") in ("Text", "Edit", "DataItem")))
    if stop_final == "yes":
        dump_page(f, page, FINAL)
        shot = f.screenshot(store_as=f"final_page_shot_{kind}")
        mine = (f"ημέρες {csl['header'][0]} · αποδοχές {csl['header'][1]} · καταβλητέες {csl['header'][2]}"
                if csl else "-")
        f.halt(f"ΣΤΟΠ ΠΡΙΝ το «{FINAL}» (stop.before_final=yes) — {title} {month:02d}/{year} αρ. {number}, "
               f"ΔΕΝ υποβλήθηκε. Το CSL01 μας: {mine}. Φωτογραφία {shot} — η σελίδα είναι στο log")
    started = f.now_epoch()
    with f.time_limit(90):                      # ONCE: never pressed a second time
        f.uia_click(FINAL, control_type="Button", exact=True, timeout=30, fallback_click=False, **SUBMIT_WIN)
    f.committed(f"ΥΠΟΒΛΗΘΗΚΕ {title} {month:02d}/{year} αρ. {number} — να ΜΗΝ ξανατρέξει")
    f.set_result(f"ΑΠΔ {kind}", f"ΥΠΟΒΛΗΘΗΚΕ αρ. {number}")
    f.log(f"ΥΠΟΒΛΗΘΗΚΕ {title} {month:02d}/{year} (αρ. {number}) — αναμονή αποδεικτικού")
    return f"ΥΠΟΒΛΗΘΗΚΕ αρ. {number} · " + receipt(f, kind, title, year, month, number, started, contrib, folder,
                                                   name, downloads, afm)


def receipt(f, kind, title, year, month, number, started, contrib, folder, name, downloads, client_afm) -> str:
    """The receipt PDF (EpsilonSubmit downloads it, Acrobat opens it): number, ΑΦΜ, period and fund checked,
    amounts vs the PDF εισφορών (Petros 23/10/2026), moved to the company's folder. Never fails the client:
    the ΑΠΔ is already submitted."""
    if not downloads:
        f.log("αποδεικτικό: έμεινε στα Downloads (χωρίς downloads.folder)", level="WARNING")
        return "αποδεικτικό στα Downloads"
    pdf = f.find_file(downloads, "*.pdf", newest=True, newer_than=started - 1, timeout=120, required=False,
                      store_as=f"receipt_{kind}", on_error="continue")
    if not pdf or pdf.get("found") not in (True, "true"):
        f.log("αποδεικτικό: δεν βρέθηκε νέο PDF στα Downloads σε 120″", level="WARNING")
        return "ΧΩΡΙΣ ΑΠΟΔΕΙΚΤΙΚΟ (δεν κατέβηκε σε 120″)"
    f.wait_for_file(pdf["path"], timeout=30, settle=1, on_error="continue")
    # It opens by itself in the PC's PDF app (10/11/2026: Acrobat on the reference PC, Chrome on LENOVO-PC —
    # «report (394).pdf» stayed open in a tab): Acrobat & co. → closed; Chrome / Edge → ONLY its tab closes.
    # kill with on_error=continue: Acrobat's 2nd process can end by itself between listing and killing («PID
    # not found»: 13/11/2026 ΚΑΡΛΑΤΗΡΑΣ, after the submission → the receipt stayed in Downloads)
    tab_left = False
    for _ in range(10):
        if pdf_app_open(f):
            kill_pdf_apps(f)
            break
        if f.window_exists(PDF_BROWSER, pdf["name"]):
            tab_left = not close_pdf_tab(f, pdf["name"])
            break
        f.wait(1)
    if tab_left:
        f.set_result(f"αποδεικτικό {kind} στον browser", f"η καρτέλα «{pdf['name']}» έμεινε ανοιχτή")
    text = f.pdf_text(pdf["path"], store_as=f"receipt_text_{kind}", on_error="continue") or ""
    num = re.search(r"Αριθμ\. Υποβολής\s+(\d+)", text)
    afm = re.search(r"Α\.Φ\.Μ\.\s+(\d+)", text)
    per = re.search(r"Περίοδος Από\s+(\d+/\d+)", text)
    problems = []
    if not num or num.group(1) != number:
        problems.append(f"αριθμός {num.group(1) if num else '-'} αντί για {number}")
    if not afm or afm.group(1) != client_afm:
        problems.append(f"ΑΦΜ {afm.group(1) if afm else '-'}")
    if not per or per.group(1) != f"{month}/{year}":
        problems.append(f"περίοδος {per.group(1) if per else '-'}")
    if ("ΤΕΚΑ" in text) != (kind == "ΤΕΚΑ"):
        problems.append(f"δεν είναι αποδεικτικό {title}")
    if problems:                                # not ours: it stays in Downloads as it is
        f.log(f"το {pdf['name']} στα Downloads ΔΕΝ είναι το αποδεικτικό ({'; '.join(problems)})", level="WARNING")
        return f"ΧΩΡΙΣ ΑΠΟΔΕΙΚΤΙΚΟ (το {pdf['name']}: {'; '.join(problems)})"

    def value(label):
        m = re.search(rf"{label}\s+([\d.,]+)", text)
        return amount(m.group(1)) if m else None
    ours = (contrib[kind], contrib["επιδοτήσεις"].get(kind, ZERO))
    theirs = (value("Καταβλητέες Εισφορές"), value("Σύνολο Επιδοτήσεων"))
    same = ours == theirs
    amounts = f"αποδεικτικό: καταβλητέες {theirs[0]} · επιδοτήσεις {theirs[1]} — εμείς: {ours[0]} · {ours[1]}"
    f.log(("✔ ποσά ΣΥΜΦΩΝΟΥΝ — " if same else "✘ ΔΙΑΦΟΡΑ ΠΟΣΩΝ — ") + amounts, level="INFO" if same else "WARNING")
    # Petros 20/11/2026: «ΟΝΟΜΑ ΠΕΛΑΤΗ ΑΠΔ ΜΗΝΑΣ» / «ΟΝΟΜΑ ΠΕΛΑΤΗ ΑΠΔ ΤΕΚΑ ΜΗΝΑΣ» (was «ΑΠΔ 10 2026 ΟΝΟΜΑ_<time>»);
    # the time only when that name is already there (a 2nd run), so nothing is overwritten
    base = f"{name} {APD_KINDS[kind]['prefix']} {month:02d} {year}".rstrip(". ")
    there = f.find_file(folder, f"{base}.pdf", required=False, store_as=f"receipt_exists_{kind}", on_error="continue")
    file_name = base + (f"_{f.now()}" if there and there.get("found") in (True, "true") else "") + ".pdf"
    dest = f"{folder}\\{file_name}"
    f.move_file(pdf["path"], dest, on_error="continue")   # never fails the client: it is already submitted
    moved = f.find_file(folder, file_name, required=False, store_as=f"receipt_moved_{kind}", on_error="continue")
    if not moved or moved.get("found") not in (True, "true"):
        f.log(f"αποδεικτικό: δεν μετακινήθηκε — έμεινε {pdf['path']}", level="WARNING")
        dest = pdf["path"]
    f.set_result(f"αποδεικτικό {kind}", dest)
    attach_receipt(f, dest, f"αποδεικτικό {kind}")      # Lefteris v3: the receipt to the cloud («Υποβολές»)
    return f"{dest}" + ("" if same else f" · ΔΙΑΦΟΡΑ ΠΟΣΩΝ ({amounts})")


def close_apd(f, title):
    """EpsilonSubmit, «Υποβολή ΑΠΔ» and «Α.Π.Δ.» close (like their ×) — before the next ΑΠΔ / the end."""
    close_submit(f)
    if f.window_exists(APD, title):
        f.connect_window(window_class=APD, window_title=title, timeout=5)
        f.close_app()
    must_close(f, APD, title, 30)


def apd_one(f, kind, year, month, afm, name, folder, contrib, results, test, downloads,
            stop_final="") -> str:
    """One ΑΠΔ (ΕΦΚΑ or ΤΕΚΑ) up to the submission (or, in a test, up to the CSL01 check)."""
    k = APD_KINDS[kind]
    title = k["title"]
    f.log(f"── {title} {month:02d}/{year}")
    open_apd(f, kind)
    apd_settings(f, kind, title, year, month)
    state = calculate(f, kind, title, folder, year, month)
    if state == "saved":
        deleted, labels = delete_saved(f, kind, title, year, month)
        apd_settings(f, kind, title, year, month)
        state = calculate(f, kind, title, folder, year, month)
        if state == "saved":
            # «ήδη δημιουργήσει» again: after a deletion = something else → stop; with an EMPTY list = the
            # month's ΑΠΔ is «Υποβλήθηκε» in Epsilon (hidden, never deleted — ACTION ADVANCED 22/10/2026);
            # anything else (a toolbar we didn't recognise) → stop, never a guess
            shot = f.screenshot(store_as=f"saved_again_shot_{kind}")
            if deleted:
                problem(f, f"«Σφάλμα» ξανά μετά τη διαγραφή της αποθηκευμένης {title} — ΔΕΝ απαντήθηκε",
                        known=True, shot=shot)
            if len(labels) not in EMPTY_TOOLBAR_LABELS:
                problem(f, f"ΥΠΑΡΧΕΙ ΗΔΗ ΑΠΟΘΗΚΕΥΜΕΝΗ {title} {month:02d}/{year} που ΔΕΝ διαγράφηκε: η μπάρα των "
                           f"«Αποθηκευμένων» είχε {len(labels)} κουμπιά {labels} (αναμένονται {TOOLBAR_LABELS} για "
                           f"«Διαγραφή») — διαγραφή με το χέρι ή ρύθμιση TOOLBAR_LABELS. ΔΕΝ απαντήθηκε",
                        known=True, shot=shot)
            answer(f, "Σφάλμα", 0)
            close_apd(f, title)
            return f"ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ (στην Epsilon «Υποβλήθηκε», φωτογραφία {shot})"
    to_export_page(f, kind, title)
    _, lines = export_csl(f, kind, title)

    # CSL01 check: ΕΦΚΑ → everyone with κύριο contributions; ΤΕΚΑ → with επικουρικό (Petros 26/10/2026),
    # May / January with the 2nd Μισθοδοτική's gift (27/10/2026)
    c = read_csl(lines)
    column = "Σύν. Εισφ. Κύριου Ταμείου" if kind == "ΕΦΚΑ" else "Συν. Εισφ. Επίκ. Ταμείου"
    people = payroll_people(results[0]["rows"], column, skip=OWN_EXCLUDED.get(month, ()))
    if month in EXTRA_PERIODS and len(results) > 1 and results[1].get("rows"):
        for a, (n, v) in payroll_people(results[1]["rows"], column, only=EXTRA_PERIODS[month]).items():
            people[a] = (people.get(a, (n, ZERO))[0], people.get(a, (n, ZERO))[1] + v)
    if c:
        f.log(f"CSL01: περίοδος {c['period']} · {len(c['afm'])} εργαζόμενοι · ημέρες {c['header'][0]} · "
              f"αποδοχές {c['header'][1]} · καταβλητέες {c['header'][2]}")
    problems = csl_problems(c, kind, year, month, contrib[kind], people, contrib["άλλα"] if kind == "ΤΕΚΑ" else {})
    if problems:
        # Petros 20/11/2026: an ΑΣΥΜΦΩΝΙΑ is a PROBLEM (it was f.done = success in the cloud), with a photo,
        # never run again by itself
        problem(f, f"ΑΣΥΜΦΩΝΙΑ CSL01 — {title} {month:02d}/{year}: {' · '.join(problems)} — δεν στάλθηκε", known=True)
    f.log(f"✔ CSL01 συμφωνεί: καταβλητέες = PDF εισφορών {kind} · όλοι οι εργαζόμενοι της Μισθοδοτικής μέσα")
    f.set_result(f"CSL01 {kind}", f"ΣΥΜΦΩΝΕΙ · καταβλητέες {c['header'][2]} · {len(c['afm'])} εργαζόμενοι")
    if test:
        close_apd(f, title)
        return f"CSL01 OK (καταβλητέες {c['header'][2]})"
    out = submit_apd(f, kind, title, year, month, contrib, folder, name, downloads, afm, csl=c,
                     stop_final=stop_final)
    close_apd(f, title)
    return out


def epsilon_app(f) -> str:
    """Το exe της Μισθοδοσίας Epsilon σε αυτό το PC: Hyper (HyperM.exe) ή Extra (ExtraM.exe, Lefteris v4 23/11/2026:
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
    # submit: the cloud's choice for the job — only «yes» submits, anything else is a test. No submit at
    # all (a local run): submit. test=yes wins (as flows\ΦΜΥ.py).
    submit = ctx.get("submit")
    test = ctx.get("test") == "yes" or (submit is not None and str(submit).strip().lower() != "yes")
    why_test = "test=yes" if ctx.get("test") == "yes" else f"submit={submit}"
    f.log(f"ΑΠΔ {month:02d}/{year} — υποβολή στον ΕΦΚΑ: {'ΟΧΙ (' + why_test + ')' if test else 'ΝΑΙ'}")

    # ── 0. Epsilon as «logimate» (as flows\ΦΜΥ.py) ──
    with f.phase("launch_epsilon"):
        f.minimize_all()
        f.set_keyboard_layout("en")
        app = epsilon_app(f)

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

    # ── 1. The company (as flows\ΦΜΥ.py) ──
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

    # ── 2. Μισθοδοτικές → Excel: the ΑΠΔ month (+ April / December for May / January) ──
    with f.phase("export_excel"):
        f.create_folder(folder)
        months = apd_months(year, month)
        f.log(f"Μισθοδοτικές: {', '.join(f'{m:02d}/{y}' for y, m in months)}")
        results = []
        for n, (y, m) in enumerate(months, start=1):
            retry(f, lambda f: f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=60))
            results.append(export_payroll(f, afm, folder, y, m, n))

    # ── the Μισθοδοτικές → ΚΥΡ / ΕΠΙΚ for the ΑΠΔ ──
    with f.phase("payroll_totals"):
        first = results[0]
        if first["status"] == "0 ΕΓΓΡΑΦΕΣ":
            return finish(f, f"0 ΕΓΓΡΑΦΕΣ — δεν έχει εγγραφές μισθοδοσίας για {month_name} {year} "
                             f"(δεν βρέθηκε το «Επιλογή Όλων») — Φωτογραφία: {first['shot']}")
        if first["status"] == "ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ":
            return finish(f, f"ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ — δεν έχει εργαζόμενους για {month_name} {year}")
        notes = [f"{m:02d}/{y}: {r['status']}" for (y, m), r in zip(months[1:], results[1:]) if r["status"] != "ok"]
        kyr, epik, parts = apd_totals(month, [r["periods"] for r in results])
        f.set_result("ΚΥΡ (Μισθοδοτική)", str(kyr))
        f.set_result("ΕΠΙΚ (Μισθοδοτική)", str(epik))
        f.log(f"ΑΠΔ {month:02d}/{year}: ΚΥΡ {kyr} · ΕΠΙΚ {epik} — {' · '.join(parts) or '-'}")

    # ── 3. Εισφορές ανά Ταμείο → PDF: ONLY the ΑΠΔ month (the 2nd one is needed only as a
    # Μισθοδοτική, for its gift) ──
    with f.phase("contributions"):
        retry(f, lambda f: f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=60))
        contrib = contributions_pdf(f, afm, folder, year, month)
        f.set_result("PDF εισφορών", contrib["pdf"])
        f.set_result("ΕΦΚΑ (PDF)", str(contrib["ΕΦΚΑ"]))
        f.set_result("ΤΕΚΑ (PDF)", str(contrib["ΤΕΚΑ"]))

    # ── 4. Έλεγχος ταμείων: ΚΥΡ = ΕΦΚΑ, ΕΠΙΚ = ΤΕΚΑ + the other auxiliary funds of the PDF (e.g.
    # ΤΕΑΥΦΕ in pharmacies, 23/10/2026). Only when they agree the ΑΠΔ are made. ──
    with f.phase("check_funds"):
        aux = contrib["ΤΕΚΑ"] + sum(contrib["άλλα"].values(), ZERO)
        aux_txt = " + ".join([f"ΤΕΚΑ {contrib['ΤΕΚΑ']}"] + [f"{k} {v}" for k, v in contrib["άλλα"].items()])
        ok = kyr == contrib["ΕΦΚΑ"] and epik == aux
        line = (f"ΚΥΡ {kyr} vs ΕΦΚΑ {contrib['ΕΦΚΑ']} · ΕΠΙΚ {epik} vs {aux_txt} "
                f"({' · '.join(parts) or '-'}){' · ' + '; '.join(notes) if notes else ''}")
        f.log(("έλεγχος ταμείων ΣΥΜΦΩΝΕΙ — " if ok else "⚠ έλεγχος ταμείων ΔΕΝ ΣΥΜΦΩΝΕΙ — ") + line,
              level="INFO" if ok else "WARNING")
        f.set_result("έλεγχος ταμείων", "ΣΥΜΦΩΝΕΙ" if ok else "ΔΕΝ ΣΥΜΦΩΝΕΙ")
        if not ok:
            # Petros 20/11/2026: a PROBLEM, not a success (was finish → f.done), never run again by itself
            problem(f, f"ΑΣΥΜΦΩΝΙΑ ΤΑΜΕΙΩΝ — ΑΠΔ {month:02d}/{year}: {line} — δεν υποβλήθηκε", known=True)

    # ── 5. ΑΠΔ ΕΦΚΑ, and ΤΕΚΑ when the PDF εισφορών has ΤΕΚΑ (as fmy_python\ΑΠΔ.py) ──
    kinds = ["ΕΦΚΑ"] + (["ΤΕΚΑ"] if contrib["ΤΕΚΑ"] > 0 else [])
    downloads = ctx.get("downloads.folder") or f.downloads_folder()
    outcomes = {}
    for kind in kinds:
        with f.phase("apd_efka" if kind == "ΕΦΚΑ" else "apd_teka"):
            retry(f, lambda f: f.connect_window(window_class="TfMain", window_title="Μισθοδοσία", timeout=60))
            outcomes[kind] = apd_one(f, kind, year, month, afm, name, folder, contrib, results, test, downloads,
                                     stop_final=ctx.get("stop.before_final") or "")
            f.set_result(f"ΑΠΔ {kind}", outcomes[kind])
    summary = " · ".join(f"{k}: {v}" for k, v in outcomes.items())
    if test:
        shot = f.screenshot(store_as="test_shot")
        return finish(f, f"ΔΟΚΙΜΗ — {why_test}: ΑΠΔ {month:02d}/{year} {summary} — σταμάτησε ΠΡΙΝ την «Αποστολή ΑΠΔ», "
                         f"τίποτα δεν στάλθηκε ({line}). Φωτογραφία: {shot}")
    if all(v.startswith("ΗΔΗ") for v in outcomes.values()):
        return finish(f, f"ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ — ΑΠΔ {month:02d}/{year} {summary}")
    # Petros 20/11/2026: only the receipt(s) stay — the Excel(s) and the εισφορές PDF go, but only when every
    # receipt was found and moved (otherwise they stay as the trace of what was submitted)
    if "ΧΩΡΙΣ ΑΠΟΔΕΙΚΤΙΚΟ" not in summary and "στα Downloads" not in summary:
        delete_work_files(f, [r.get("xlsx") for r in results] + [contrib.get("pdf")], folder,
                          [f"{afm}_ΜΙΣΘΟΔΟΤΙΚΗ_{year}_{month:02d}_*.xlsx", f"{afm}_ΕΙΣΦΟΡΕΣ_{year}_{month:02d}*.pdf"])
    return finish(f, f"ΥΠΟΒΛΗΘΗΚΕ — ΑΠΔ {month:02d}/{year} {summary}", acrobat_wait=20)
