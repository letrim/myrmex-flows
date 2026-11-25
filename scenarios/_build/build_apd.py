"""Rebuild scenarios\\ΑΠΔ\\*.json (the originals were lost on 25/11/2026) from the flow's branches.
Every scenario = BASE (a whole ΑΠΔ 10/2026: Excel, εισφορές PDF, ΑΠΔ ΕΦΚΑ + ΤΕΚΑ up to the receipts) + what the
branch changes. Change the flow → change this file → run it again (it rewrites every .json).
usage (any folder): python scenarios\\_build\\build_apd.py
then:               python scenarios\\_build\\run_all.py flows\\ΑΠΔ.py scenarios\\ΑΠΔ"""
import json, copy, os, calendar

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ΑΠΔ")
AFM = "123456789"
E1, E2, E3 = "111111111", "222222222", "333333333"
FOLDER = "C:\\Myrmex\\exports\\ΤΕΣΤ ΕΤΑΙΡΙΑ - " + AFM
DOWN = "C:\\Users\\test\\Downloads"
KYR, EPIK = "Σύν. Εισφ. Κύριου Ταμείου", "Συν. Εισφ. Επίκ. Ταμείου"
MONTHS = {1: "Ιανουάριος", 2: "Φεβρουάριος", 3: "Μάρτιος", 4: "Απρίλιος", 5: "Μάιος", 6: "Ιούνιος", 7: "Ιούλιος",
          8: "Αύγουστος", 9: "Σεπτέμβριος", 10: "Οκτώβριος", 11: "Νοέμβριος", 12: "Δεκέμβριος"}
NUM = {"ΕΦΚΑ": "84000001", "ΤΕΚΑ": "84000002"}
W = lambda cls, title=None, results=(True,): {k: v for k, v in
                                               {"class_name": cls, "title": title, "results": list(results)}.items()
                                               if v is not None}


def excel(rows):
    """Μισθοδοτική Excel (as_lists): header with «Περίοδος», employee rows, a totals row (no «Εταιρεία»)."""
    out = [["Μισθοδοτική Κατάσταση"], ["Εταιρεία", "Α.Φ.Μ.", "Επώνυμο", "Όνομα", "Περίοδος", KYR, EPIK]]
    for a, p, k, e in rows:
        out.append(["ΤΕΣΤ ΕΤΑΙΡΙΑ", a, "ΕΡΓΑΖΟΜΕΝΟΣ", a[:1], p, k, e])
    out.append(["", "", "", "", "Σύνολα", sum(r[2] for r in rows), sum(r[3] for r in rows)])
    return out


def contrib_text(efka=(100, 200, 0, 300), teka=(10, 20, 0, 30), others=(), month=10, year=2026, afm=AFM):
    """The PDF «Ασφαλιστικές Εισφορές ανά Ταμείο»: per fund «Σύνολο : εργαζ. εργοδ. επιδότηση σύνολο»."""
    end = calendar.monthrange(year, month)[1]
    f = lambda v: f"{v:.2f}".replace(".", ",")
    lines = [f"Περίοδος: 1/{month}/{year} - {end}/{month}/{year}", f"Α.Φ.Μ. : {afm}", "Εργαζομένου Εργοδότη Επιδότηση Σύνολο"]
    for fund, vals in (("ΕΦΚΑ", efka), ("TEKA", teka)) + tuple(others):
        if vals:
            lines += [fund, "Σύνολο : " + " ".join(f(v) for v in vals)]
    lines.append("Γενικό Σύνολο")
    return "\n".join(lines)


def csl(period, emps, due=None):
    """CSL01: header «1…MMYYYY days(8) pay(12) due(12) date(8)», per employee a record 2 (ΑΦΜ last) and a
    record 3 (days [49:52], pay [62:72], due [123:134]); amounts in cents. due= a different header due."""
    days, pay = sum(e[1] for e in emps), sum(e[2] for e in emps)
    due = sum(e[3] for e in emps) if due is None else due
    lines = ["1" + " " * 20 + period + str(days).zfill(8) + str(pay).zfill(12) + str(due).zfill(12) + "26112026"]
    for a, d, p, c in emps:
        lines.append(("2" + " " * 20 + "ΕΡΓΑΖΟΜΕΝΟΣ").ljust(80) + a)
        r = list("3".ljust(140))
        r[49:52] = list(str(d).zfill(3))
        r[62:72] = list(str(p).zfill(10))
        r[123:134] = list(str(c).zfill(11))
        lines.append("".join(r))
    return lines + ["EOF"]


def row(number, status="ΠΡΟΣ ΕΛΕΓΧΟ ΥΠΟΒΟΛΗΣ", actions=("Έλεγχος Υποβολής",), period="10/2026", kind="01 ΚΑΝΟΝΙΚΗ"):
    """One row of the ΕΦΚΑ list of submissions as uia_page gives it."""
    return ([{"type": "Hyperlink", "name": number}, {"type": "Text", "name": kind},
             {"type": "Text", "name": f"{period} - {period}"}, {"type": "Text", "name": status}]
            + [{"type": "Hyperlink", "name": a} for a in actions])


def page(*rows):
    out = [{"type": "Text", "name": "Αριθμός"}]
    for r in rows:
        out += r
    return out + [{"type": "Text", "name": "Περίοδος Υποβολής"}]


def receipt_text(kind, number, period, due, afm=AFM, sub="0,00"):
    return (f"Αποδεικτικό Υποβολής ΑΠΔ{' ΤΕΚΑ' if kind == 'ΤΕΚΑ' else ''}\nΑριθμ. Υποβολής {number}\n"
            f"Α.Φ.Μ. {afm}\nΠερίοδος Από {period}\nΚαταβλητέες Εισφορές {due}\nΣύνολο Επιδοτήσεων {sub}")


def money(cents):
    return f"{cents / 100:.2f}".replace(".", ",")


def apd_store(kind, due_cents, month=10, year=2026, emps=None, number=None):
    """Everything one ΑΠΔ (ΕΦΚΑ or ΤΕΚΑ) reads, from its settings to the receipt."""
    n, k, per = number or NUM[kind], kind, f"{month}/{year}"
    emps = emps or ((E1, 25, 100000, due_cents // 3), (E2, 25, 120000, due_cents - due_cents // 3))
    return {
        f"apd_combo_{k}_1": "01 Κανονική", f"apd_combo_{k}_2": MONTHS[month], f"apd_combo_{k}_3": MONTHS[month],
        f"apd_year_{k}_1": str(year), f"apd_year_{k}_2": str(year),
        f"calc_bar_runs_{k}": [[10, 90], [120, 170]],
        f"calc_done_{k}": True,
        f"export_page_{k}": True,
        f"apd_edit_{k}_1": "C:\\Program Files (x86)\\epsilon net\\hmisthos\\",
        f"csl_{k}": csl(f"{month:02d}{year}", emps),
        f"send_page_{k}": [{"type": "Text", "name": "Το αρχείο ανέβηκε επιτυχώς"}],
        f"list_page_{k}": page(row(n, period=per), row("83000000", "ΥΠΟΒΟΛΗ", ("Επανέκδοση Αποδεικτικού",), period="9/2026")),
        f"row_page_{k}": page(row(n, period=per)),
        f"check_ok_page_{k}": [{"type": "Button", "name": "OK"}],
        f"check_page_{k}": page(row(n, "ΜΕ ΔΥΝΑΤΟΤΗΤΑ ΥΠΟΒΟΛΗΣ", ("Υποβολή",), period=per)),
        f"continue_page_{k}": [{"type": "Edit", "name": "Περίοδος Υποβολής", "value": per},
                               {"type": "Button", "name": "Συνέχεια"}],
        f"final_page_{k}": [{"type": "Text", "name": "Σύνολο Εισφορών"},
                            {"type": "Button", "name": "Υποβολή & Έκδοση Αποδεικτικού"}],
        f"receipt_{k}": {"found": True, "name": f"report_{k}.pdf", "path": f"{DOWN}\\report_{k}.pdf"},
        f"receipt_text_{k}": receipt_text(kind, n, per, money(due_cents)),
        f"receipt_exists_{k}": {"found": False},
        f"receipt_moved_{k}": {"found": True, "name": "moved.pdf", "path": FOLDER + "\\moved.pdf"},
    }


def month_base(month, year=2026, rows=None, rows2=None, efka=300, teka=30):
    """A whole scenario for ΑΠΔ month/year: Excel(s), εισφορές PDF, ΑΠΔ ΕΦΚΑ (+ ΤΕΚΑ when teka)."""
    end = calendar.monthrange(year, month)[1]
    rows = rows or ((E1, MONTHS[month], 100.0, 10.0), (E2, MONTHS[month], 200.0, 20.0))
    store = {
        "company_search": AFM,
        "select_all_runs_1": [[10, 60]],
        "exported_file_1": {"found": True, "name": "export.xlsx", "path": FOLDER + "\\export.xlsx"},
        "payroll_rows_1": excel(rows),
        "date_1": f"1/{month}/{year}", "date_2": f"{end}/{month}/{year}",
        "print_bar_runs": [[10, 80], [100, 160], [180, 220]],
        "screen": [1920, 1080],
        "contrib_pdf_saved": {"found": True, "name": f"{AFM}_ΕΙΣΦΟΡΕΣ_{year}_{month:02d}.pdf",
                              "path": f"{FOLDER}\\{AFM}_ΕΙΣΦΟΡΕΣ_{year}_{month:02d}.pdf"},
        "contrib_pdf_text": contrib_text(efka=(efka / 3, efka - efka / 3, 0, efka),
                                         teka=(teka / 3, teka - teka / 3, 0, teka) if teka else None,
                                         month=month, year=year),
        **apd_store("ΕΦΚΑ", efka * 100, month, year),
    }
    if teka:
        store.update(apd_store("ΤΕΚΑ", teka * 100, month, year))
    if rows2 is not None:
        store.update({"select_all_runs_2": [[10, 60]],
                      "exported_file_2": {"found": True, "name": "export2.xlsx", "path": FOLDER + "\\export2.xlsx"},
                      "payroll_rows_2": excel(rows2)})
    return {
        "description": "",
        "vars": {"period.year": str(year), "period.month": str(month), "client.afm": AFM,
                 "client.name": "ΤΕΣΤ ΕΤΑΙΡΙΑ", "output.folder": FOLDER, "downloads.folder": DOWN},
        "windows": [
            W("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση", [False, True]),
            W("#32770", "Αποθήκευση"),
            W("TfChooseCompanies", "Επιλογή Εταιριών"),
            W("TfrxPreviewForm", None, [False, True]),
            W("TfrxPDFExportDialog"),
            W("TfAPD", "Α.Π.Δ. ΤΕΚΑ"),
            W("TfAPD", "Α.Π.Δ."),
            W("TfAPD", "ΤΕΚΑ", [False]), W("TfAPD", "Δημοσ", [False]), W("TfAPD", "Οικοδομ", [False]),
            W("TfAPD", "Ναυτ", [False]),
            W("TfAPD", None, [False, True]),
        ],
        "store": store,
        "fail": [],
    }


BASE = month_base(10)
EFKA_ONLY = month_base(10, rows=((E1, "Οκτώβριος", 100.0, 0.0), (E2, "Οκτώβριος", 200.0, 0.0)), teka=0)
SCEN = {}


def sc(name, desc, vars=None, windows=None, store=None, fail=None, expect=None, drop_windows=(), base=None):
    s = copy.deepcopy(base or BASE)
    s["description"] = desc
    s["vars"].update(vars or {})
    s["windows"] = [w for w in s["windows"] if (w["class_name"], w.get("title")) not in drop_windows]
    s["windows"] = (windows or []) + s["windows"]
    s["store"].update(store or {})
    s["fail"] = (fail or []) + s["fail"]
    if expect:
        s["expect"] = expect
    SCEN[name] = s
    return s


DONE = lambda text: {"status": "success", "message_contains": text}
FAILED = lambda text, retry=None: {"status": "failed", "message_contains": text,
                                   **({} if retry is None else {"retryable": retry})}
HALT = lambda text: {"status": "halted", "message_contains": text}
TEST = {"test": "yes"}
# TMessageForm titles that must NOT be found (an untitled rule would answer them too)
NO_MSG = [W("TMessageForm", t, [False]) for t in ("Σφάλμα", "Πληροφορία", "Επιβεβαίωση", "Διαγραφή")]
SAVED_ONCE = {"calc_error_buttons_ΕΦΚΑ": ["OK"], "saved_bar_runs_ΕΦΚΑ": [[10, 90], [120, 170]],
              "saved_grid_ΕΦΚΑ": True, "back_bar_runs_ΕΦΚΑ": [[10, 90], [120, 170], [200, 250]],
              "back_page1_ΕΦΚΑ": True}
TOOLBAR7 = [[6, 89], [100, 159], [171, 236], [248, 312], [336, 486], [511, 617], [629, 746]]

# ── the normal run, test / submit switches ──
sc("apd_submit", "ΑΠΔ ΕΦΚΑ + ΤΕΚΑ 10/2026 μέχρι την υποβολή, αποδεικτικά «ΤΕΣΤ ΕΤΑΙΡΙΑ ΑΠΔ 10 2026» / «… ΑΠΔ ΤΕΚΑ …»",
   expect=DONE("ΤΕΚΑ: ΥΠΟΒΛΗΘΗΚΕ αρ. 84000002"))
sc("normal_month", "test=yes: ταμεία και CSL01 συμφωνούν, σταματά ΠΡΙΝ την «Αποστολή ΑΠΔ»", vars=TEST,
   expect=DONE("ΔΟΚΙΜΗ — test=yes"))
sc("apd_submit_cloud_no", "cloud submit=no → ΔΟΚΙΜΗ, τίποτα δεν στάλθηκε", vars={"submit": "no"},
   expect=DONE("ΔΟΚΙΜΗ — submit=no"))
sc("apd_efka_only", "Χωρίς ΤΕΚΑ στο PDF εισφορών → μόνο ΑΠΔ ΕΦΚΑ", base=EFKA_ONLY,
   expect=DONE("ΑΠΔ 10/2026 ΕΦΚΑ: ΥΠΟΒΛΗΘΗΚΕ αρ. 84000001"))

# ── login / company / Μισθοδοτική (as ΦΜΥ) ──
sc("login_error", "«Σφάλμα» της Epsilon στη σύνδεση → πρόβλημα", windows=[W("TMessageForm", "Σφάλμα")],
   expect=FAILED("«Σφάλμα» στη σύνδεση"))
sc("wrong_company_search", "Η αναζήτηση εταιριών δεν γράφει το ΑΦΜ → σταματά", store={"company_search": "1234"},
   expect=FAILED("η αναζήτηση εταιριών γράφει"))
sc("kad_question", "«Επιβεβαίωση» ΚΑΔ στο άνοιγμα → «Όχι», φωτογραφία, συνέχεια", vars=TEST,
   windows=[W("TMessageForm", "Επιβεβαίωση", [True, True, False])], expect=DONE("ΔΟΚΙΜΗ"))
sc("payroll_three_rows", "Η Μισθοδοτική ανοίγει με 3 γραμμές (LENOVO-PC), όχι με 4", vars=TEST,
   windows=[W("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση", [False, False, True])],
   drop_windows=[("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση")], expect=DONE("ΔΟΚΙΜΗ"))
sc("payroll_unknown_rows", "Καμία γνωστή γραμμή δεν ανοίγει τη Μισθοδοτική → πρόβλημα (PAYROLL_ROWS)",
   drop_windows=[("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση")], expect=FAILED("PAYROLL_ROWS"))
sc("no_records", "Χωρίς «Επιλογή Όλων» → 0 ΕΓΓΡΑΦΕΣ (επιτυχία, δεν υποβλήθηκε)", store={"select_all_runs_1": []},
   expect=DONE("0 ΕΓΓΡΑΦΕΣ"))
sc("no_employees", "«Πληροφορία» στο «Εξαγωγή σε EXCEL» → ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ",
   windows=[W("TMessageForm", "Πληροφορία")], drop_windows=[("#32770", "Αποθήκευση")],
   expect=DONE("ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ"))

# ── gifts: April / May / December / January ──
sc("april", "ΑΠΔ Απριλίου: το Δώρο Πάσχα της Μισθοδοτικής ΔΕΝ μετράει (πάει στη Μαΐου)", vars=TEST,
   base=month_base(4, rows=((E1, "Απρίλιος", 100.0, 10.0), (E2, "Απρίλιος", 200.0, 20.0),
                            (E1, "Δώρο Πάσχα", 50.0, 5.0))), expect=DONE("ΔΟΚΙΜΗ"))
sc("may", "ΑΠΔ Μαΐου: + το Δώρο Πάσχα από τη Μισθοδοτική Απριλίου (2 Μισθοδοτικές)", vars=TEST,
   base=month_base(5, rows=((E1, "Μάιος", 100.0, 10.0), (E2, "Μάιος", 200.0, 20.0)),
                   rows2=((E1, "Απρίλιος", 100.0, 10.0), (E1, "Δώρο Πάσχα", 50.0, 5.0)), efka=350, teka=35),
   expect=DONE("ΔΟΚΙΜΗ"))
sc("december", "ΑΠΔ Δεκεμβρίου: το Δώρο Χριστουγέννων ΔΕΝ μετράει (πάει στη Ιανουαρίου)", vars=TEST,
   base=month_base(12, rows=((E1, "Δεκέμβριος", 100.0, 10.0), (E2, "Δεκέμβριος", 200.0, 20.0),
                             (E2, "Δώρο Χριστουγέννων", 80.0, 8.0))), expect=DONE("ΔΟΚΙΜΗ"))
sc("january", "ΑΠΔ Ιανουαρίου: + το Δώρο Χριστουγέννων από τη Μισθοδοτική Δεκεμβρίου του προηγούμενου έτους",
   vars=TEST, base=month_base(1, rows=((E1, "Ιανουάριος", 100.0, 10.0), (E2, "Ιανουάριος", 200.0, 20.0)),
                              rows2=((E2, "Δεκέμβριος", 200.0, 20.0), (E2, "Δώρο Χριστουγέννων", 80.0, 8.0)),
                              efka=380, teka=38), expect=DONE("ΔΟΚΙΜΗ"))
s = sc("january_december_no_records", "ΑΠΔ Ιανουαρίου, ο Δεκέμβριος χωρίς εγγραφές → μόνο ο Ιανουάριος, συνέχεια",
       vars=TEST, base=month_base(1, rows=((E1, "Ιανουάριος", 100.0, 10.0), (E2, "Ιανουάριος", 200.0, 20.0)),
                                  rows2=()), store={"select_all_runs_2": []}, expect=DONE("ΔΟΚΙΜΗ"))

# ── εισφορές PDF / έλεγχος ταμείων ──
sc("funds_mismatch", "ΚΥΡ Μισθοδοτικής ≠ ΕΦΚΑ του PDF → ΑΣΥΜΦΩΝΙΑ ΤΑΜΕΙΩΝ (πρόβλημα, δεν ξανατρέχει)",
   store={"contrib_pdf_text": contrib_text(efka=(100, 210, 0, 310))}, expect=FAILED("ΑΣΥΜΦΩΝΙΑ ΤΑΜΕΙΩΝ", retry=False))
o = month_base(10, rows=((E1, "Οκτώβριος", 100.0, 10.0), (E2, "Οκτώβριος", 200.0, 20.0), (E3, "Οκτώβριος", 50.0, 5.0)),
               efka=350, teka=30)
o["store"]["contrib_pdf_text"] = contrib_text(efka=(100, 250, 0, 350), others=(("ΤΕΑΥΦΕ", (2, 3, 0, 5)),))
o["store"]["csl_ΕΦΚΑ"] = csl("102026", ((E1, 25, 100000, 10000), (E2, 25, 120000, 20000), (E3, 25, 50000, 5000)))
sc("other_aux_fund", "Φαρμακείο: ΕΠΙΚ = ΤΕΚΑ + ΤΕΑΥΦΕ· η ΑΠΔ ΤΕΚΑ χωρίς τον εργαζόμενο του ΤΕΑΥΦΕ γίνεται δεκτή",
   vars=TEST, base=o, expect=DONE("ΔΟΚΙΜΗ"))
sc("contrib_wrong_afm", "Το PDF εισφορών είναι άλλου ΑΦΜ → σταματά",
   store={"contrib_pdf_text": contrib_text(afm="999999999")}, expect=FAILED("για ΑΦΜ 999999999"))
sc("contrib_pdf_no_name", "Epsilon Extra (ΦΡΑΝΤΖΗΣ 24/11): το PDF εισφορών σώθηκε ως «.pdf» → βρίσκεται και μετονομάζεται",
   vars=TEST, store={"contrib_pdf_saved": {"found": True, "name": ".pdf", "path": FOLDER + "\\.pdf"}},
   expect=DONE("ΔΟΚΙΜΗ"))
sc("print_bar_two_runs", "Η κάτω μπάρα των Εισφορών έχει 2 κείμενα αντί για 3/4 → δεν πατάει, σταματά",
   store={"print_bar_runs": [[10, 80], [100, 160]]}, windows=[W("TfrxPreviewForm", None, [False])],
   drop_windows=[("TfrxPreviewForm", None)], expect=FAILED("έχει 2 κείμενα"))
sc("pdf_button_no_dialog", "Το κλικ στο «Εξαγωγή σε PDF» δεν άνοιξε τον διάλογο → σταματά",
   drop_windows=[("TfrxPDFExportDialog", None)], expect=FAILED("δεν άνοιξε το «Εξαγωγή σε PDF»"))

# ── the ΑΠΔ window / calculation ──
sc("apd_menu_not_opened", "Η «Α.Π.Δ.» δεν ανοίγει από την 1η γραμμή της αναζήτησης → σταματά (v6: χωρίς rows)",
   windows=[W("TfAPD", None, [False])], drop_windows=[("TfAPD", None)], expect=FAILED("δεν άνοιξε η «Α.Π.Δ.»"))
sc("apd_wrong_window", "Η 1η γραμμή άνοιξε άλλη ΑΠΔ (Δημοσίου) → σταματά",
   windows=[W("TfAPD", "Δημοσ")], drop_windows=[("TfAPD", "Δημοσ")], expect=FAILED("άνοιξε άλλο παράθυρο ΑΠΔ"))
sc("apd_calc_bar_three", "Η κάτω μπάρα έχει 3 κείμενα αντί για 2 → δεν πατάει «Υπολογισμός»",
   store={"calc_bar_runs_ΕΦΚΑ": [[10, 90], [120, 170], [200, 250]]}, expect=FAILED("δεν πατάω «Υπολογισμός»"))
review = [W("TuAPDCaseReview", None, [True, False])]
sc("apd_incomplete", "«Ελλιπή στοιχεία» μετά τον Υπολογισμό → ΕΡΓΑΖΟΜΕΝΟΙ ΕΚΤΟΣ ΑΠΔ, ΔΕΝ πατιέται OK, δεν ξανατρέχει",
   windows=review, store={"review_missing_ΕΦΚΑ": True, "review_reason_ΕΦΚΑ": "Μηδενική μισθοδοσία λόγω λοχείας"},
   expect=FAILED("ΕΡΓΑΖΟΜΕΝΟΙ ΕΚΤΟΣ ΑΠΔ", retry=False))
sc("apd_incomplete_accepted", "«Ελλιπή στοιχεία» με review.accept=yes → OK με έγκριση χρήστη και συνέχεια",
   vars={**TEST, "review.accept": "yes"}, windows=review,
   store={"review_missing_ΕΦΚΑ": True, "review_reason_ΕΦΚΑ": "Μηδενική μισθοδοσία λόγω λοχείας"},
   expect=DONE("ΔΟΚΙΜΗ"))
sc("apd_not_recalculated", "Μόνο «Μη επανυπολογισμένοι περίοδοι» (το φύλλο Ελλιπή κρυφό) → OK και συνέχεια",
   vars=TEST, windows=review, store={"review_missing_ΕΦΚΑ": False}, expect=DONE("ΔΟΚΙΜΗ"))
saved = [W("TMessageForm", "Σφάλμα", [False, True, False])]
sc("apd_saved_deleted", "«Έχετε ήδη δημιουργήσει Α.Π.Δ.» → Αποθηκευμένα → «Διαγραφή» (7 κουμπιά) → Ναι → νέος Υπολογισμός",
   vars=TEST, windows=saved,
   store={**SAVED_ONCE, "saved_toolbar_runs_ΕΦΚΑ": TOOLBAR7, "delete_buttons_ΕΦΚΑ": ["Ναι", "Όχι"]},
   expect=DONE("ΔΟΚΙΜΗ"))
sc("apd_saved_nothing_to_delete", "«ήδη δημιουργήσει» αλλά η λίστα κενή (5 κουμπιά) → δεν διαγράφει, ο νέος Υπολογισμός περνά",
   vars=TEST, windows=saved, store={**SAVED_ONCE, "saved_toolbar_runs_ΕΦΚΑ": TOOLBAR7[:2] + TOOLBAR7[4:]},
   expect=DONE("ΔΟΚΙΜΗ"))
sc("apd_saved_again_after_delete", "«ήδη δημιουργήσει» ΞΑΝΑ μετά τη διαγραφή → σταματά χωρίς απάντηση",
   windows=[W("TMessageForm", "Σφάλμα", [False, True, True, False])],
   store={**SAVED_ONCE, "saved_toolbar_runs_ΕΦΚΑ": TOOLBAR7, "delete_buttons_ΕΦΚΑ": ["Ναι", "Όχι"]},
   expect=FAILED("«Σφάλμα» ξανά μετά τη διαγραφή", retry=False))
sc("apd_saved_not_deleted", "«ήδη δημιουργήσει» ξανά χωρίς διαγραφή (άγνωστη μπάρα) → ΥΠΑΡΧΕΙ ΗΔΗ ΑΠΟΘΗΚΕΥΜΕΝΗ, σταματά",
   windows=[W("TMessageForm", "Σφάλμα", [False, True, True, False])],
   store={**SAVED_ONCE, "saved_toolbar_runs_ΕΦΚΑ": TOOLBAR7[:5]},
   expect=FAILED("ΥΠΑΡΧΕΙ ΗΔΗ ΑΠΟΘΗΚΕΥΜΕΝΗ", retry=False))
sc("apd_delete_wrong_question", "Μετά το κλικ στη «Διαγραφή» δεν βγήκε «Διαγραφή» [Ναι / Όχι] → σταματά, τίποτα δεν απαντήθηκε",
   windows=saved, store={**SAVED_ONCE, "saved_toolbar_runs_ΕΦΚΑ": TOOLBAR7, "delete_buttons_ΕΦΚΑ": ["OK"]},
   expect=FAILED("δεν βγήκε η ερώτηση «Διαγραφή»"))
sc("apd_unexpected_message", "Μήνυμα της Epsilon που κανείς δεν περίμενε μετά τον Υπολογισμό → HALT, δεν απαντιέται",
   windows=NO_MSG + [W("TMessageForm", "Προσοχή"), W("TMessageForm", None, [False, True])],
   store={"unexpected_buttons": ["Ναι", "Όχι"]}, expect=HALT("μήνυμα της Epsilon μετά τον «Υπολογισμό»"))
sc("apd_info_message", "«Πληροφορία» με μόνο [ΟΚ] μετά τον Υπολογισμό = ενημέρωση → ΟΚ, σημείωση, συνέχεια",
   vars=TEST, windows=[W("TMessageForm", "Πληροφορία", [False, False, True, False])]
   + [w for w in NO_MSG if w["title"] != "Πληροφορία"] + [W("TMessageForm", None, [False, True, False])],
   store={"unexpected_buttons": ["OK"], "unexpected_info_buttons": ["OK"]}, expect=DONE("ΔΟΚΙΜΗ"))

# ── CSL01 check ──
sc("apd_csl_wrong_period", "Το CSL01 γράφει άλλη περίοδο → ΑΣΥΜΦΩΝΙΑ CSL01, δεν στάλθηκε",
   store={"csl_ΕΦΚΑ": csl("092026", ((E1, 25, 100000, 10000), (E2, 25, 120000, 20000)))},
   expect=FAILED("περίοδος 092026", retry=False))
sc("apd_csl_wrong_due", "Καταβλητέες CSL01 ≠ ΕΦΚΑ του PDF (και ≠ γραμμές) → ΑΣΥΜΦΩΝΙΑ CSL01",
   store={"csl_ΕΦΚΑ": csl("102026", ((E1, 25, 100000, 10000), (E2, 25, 120000, 20000)), due=29000)},
   expect=FAILED("καταβλητέες 290.00 αντί για 300.00", retry=False))
sc("apd_csl_missing_employee", "Εργαζόμενος με εισφορές στη Μισθοδοτική λείπει από το CSL01 (TRAVEL SENSE) → ΑΣΥΜΦΩΝΙΑ",
   store={"csl_ΕΦΚΑ": csl("102026", ((E1, 25, 220000, 30000),))}, expect=FAILED("λείπουν από το CSL01", retry=False))

# ── EpsilonSubmit / ΕΦΚΑ ──
sc("apd_open_period_retry", "«Open period button not found» [ΟΚ] στην 1η «Αποστολή ΑΠΔ» (13/11) → ΟΚ, κλείσιμο, ξανά → υποβολή",
   windows=[W("TMessageForm", "Σφάλμα", [False, False, True, False])],
   store={"period_error_buttons": ["OK"],
          "send_page_ΕΦΚΑ": [{"type": "Text", "name": "Open period button not found (j_idt210) - openPeriodForm"},
                             {"type": "Text", "name": "Το αρχείο ανέβηκε επιτυχώς"}]},
   expect=DONE("ΕΦΚΑ: ΥΠΟΒΛΗΘΗΚΕ"))
sc("apd_submit_already_in_efka", "«Σφάλμα: Έχετε υποβάλλει ήδη … 10/2026» (ΨΑΛΛΙΔΑ) → ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ, τίποτα ξανά",
   base=EFKA_ONLY, windows=[W("TMessageForm", "Σφάλμα", [False, False, True, False])],
   store={"sent_error_buttons": ["OK"],
          "send_page_ΕΦΚΑ": [{"type": "Text", "name": "Έχετε υποβάλλει ήδη ΑΠΔ για την περίοδο 10/2026"}]},
   expect=DONE("ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ"))
sc("apd_already_submitted_list", "Η λίστα του ΕΦΚΑ έχει ήδη υποβληθείσα ΑΠΔ 10/2026 (Επανέκδοση) → η νέα ΔΕΝ υποβάλλεται",
   base=EFKA_ONLY, store={"list_page_ΕΦΚΑ": page(row(NUM["ΕΦΚΑ"]), row("83990000", "ΥΠΟΒΟΛΗ",
                                                                          ("Επανέκδοση Αποδεικτικού",)))},
   expect=DONE("ΗΔΗ ΥΠΟΒΛΗΘΕΙΣΑ"))
pending = {"send_page_ΕΦΚΑ": page(row(NUM["ΕΦΚΑ"])),
           "details_page_ΕΦΚΑ": [{"type": "Text", "name": "Λοιπά Στοιχεία Υποβολής"},
                                 {"type": "Edit", "name": "14 Hμερών Ασφάλισης:", "value": "50"},
                                 {"type": "Edit", "name": "15 Αποδοχών:", "value": "2.200,00"},
                                 {"type": "Edit", "name": "16 Σύνολο Εισφορών:", "value": "300,00"}]}
sc("apd_submit_pending", "Εκκρεμής ΑΠΔ του μήνα στον ΕΦΚΑ (801646253): έλεγχος, τα «Λοιπά Στοιχεία» = νέο CSL01 → υποβάλλεται αυτή",
   base=EFKA_ONLY, store=pending, expect=DONE("ΥΠΟΒΛΗΘΗΚΕ αρ. 84000001"))
sc("apd_pending_differs", "Εκκρεμής ΑΠΔ με άλλα ποσά από το νέο CSL01 → ΔΕΝ υποβάλλεται, να ακυρωθεί στον ΕΦΚΑ",
   base=EFKA_ONLY, store={**pending, "details_page_ΕΦΚΑ": pending["details_page_ΕΦΚΑ"][:3]
                          + [{"type": "Edit", "name": "16 Σύνολο Εισφορών:", "value": "310,00"}]},
   expect=FAILED("ΔΙΑΦΕΡΕΙ από το νέο CSL01", retry=False))
sc("apd_submit_check_errors", "Ο έλεγχος του ΕΦΚΑ βγάζει «ΜΕ ΛΑΘΗ ΥΠΟΒΟΛΗΣ» → ΣΤΑΜΑΤΗΣΕ, δεν υποβλήθηκε, δεν ξανατρέχει",
   base=EFKA_ONLY, store={"check_page_ΕΦΚΑ": page(row(NUM["ΕΦΚΑ"], "ΜΕ ΛΑΘΗ ΥΠΟΒΟΛΗΣ", ()))},
   expect=FAILED("ΛΑΘΗ ΣΤΟΝ ΕΛΕΓΧΟ", retry=False))
sc("apd_bad_efka_login", "«Αποτυχία στην Αυθεντικοποίηση του χρήστη» στο EpsilonSubmit (ΓΕΩΡΓΟΥΛΗΣ 14/11) → ΛΑΘΟΣ ΚΩΔΙΚΟΙ ΕΦΚΑ",
   base=EFKA_ONLY, store={"send_page_ΕΦΚΑ": [{"type": "Text", "name": "Αποτυχία στην Αυθεντικοποίηση του χρήστη! "
                                                                      "Η σύνδεσή σας δεν ήταν δυνατή."}]},
   expect=FAILED("ΛΑΘΟΣ ΚΩΔΙΚΟΙ ΕΦΚΑ —", retry=False))
sc("apd_send_unread", "Το EpsilonSubmit δεν διαβάζεται σε 2 «Αποστολή ΑΠΔ» (ACTION 17/11) → ΛΑΘΟΣ ΚΩΔΙΚΟΙ ΕΦΚΑ;",
   base=EFKA_ONLY, store={"send_page_ΕΦΚΑ": []}, expect=FAILED("ΛΑΘΟΣ ΚΩΔΙΚΟΙ ΕΦΚΑ;", retry=False))
sc("apd_stop_at_list", "stop.before_final=list → ΣΤΟΠ στη λίστα του ΕΦΚΑ, τίποτα δεν πατήθηκε", base=EFKA_ONLY,
   vars={"stop.before_final": "list"}, expect=HALT("ΣΤΟΠ στη λίστα του ΕΦΚΑ"))
sc("apd_stop_before_final", "stop.before_final=yes → μέχρι την τελευταία σελίδα, HALT χωρίς «Υποβολή & Έκδοση Αποδεικτικού»",
   base=EFKA_ONLY, vars={"stop.before_final": "yes"}, expect=HALT("ΣΤΟΠ ΠΡΙΝ το «Υποβολή & Έκδοση Αποδεικτικού»"))
sc("apd_submit_wrong_period_page", "Η σελίδα υποβολής γράφει άλλη «Περίοδος Υποβολής» → σταματά πριν τη Συνέχεια",
   base=EFKA_ONLY, store={"continue_page_ΕΦΚΑ": [{"type": "Edit", "name": "Περίοδος Υποβολής", "value": "9/2026"},
                                                 {"type": "Button", "name": "Συνέχεια"}]},
   expect=FAILED("γράφει περίοδο «9/2026»"))
sc("apd_submit_no_downloads", "Υποβλήθηκε αλλά δεν κατέβηκε νέο PDF σε 120″ → ΧΩΡΙΣ ΑΠΟΔΕΙΚΤΙΚΟ, τα αρχεία εργασίας μένουν",
   base=EFKA_ONLY, store={"receipt_ΕΦΚΑ": {"found": False}}, expect=DONE("ΧΩΡΙΣ ΑΠΟΔΕΙΚΤΙΚΟ"))
sc("apd_submit_receipt_amounts", "Το αποδεικτικό γράφει άλλες καταβλητέες → υποβλήθηκε, με «ΔΙΑΦΟΡΑ ΠΟΣΩΝ» στο αποτέλεσμα",
   base=EFKA_ONLY, store={"receipt_text_ΕΦΚΑ": receipt_text("ΕΦΚΑ", NUM["ΕΦΚΑ"], "10/2026", "290,00")},
   expect=DONE("ΔΙΑΦΟΡΑ ΠΟΣΩΝ"))

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for name, s in SCEN.items():
        s = {k: v for k, v in s.items() if v not in ([], None)}
        with open(os.path.join(OUT, name + ".json"), "w", encoding="utf-8", newline="\n") as fh:
            json.dump(s, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
    print("written:", len(SCEN))
