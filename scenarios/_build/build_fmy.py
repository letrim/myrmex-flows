"""Rebuild scenarios\\ΦΜΥ\\*.json (the originals were lost on 25/11/2026) from the flow's branches.
Every scenario = BASE (a whole ΦΜΥ 09/2026 up to the PDF) + what the branch changes. Change the flow →
change this file → run it again (it rewrites every .json).
usage (any folder): python scenarios\\_build\\build_fmy.py
then:               python scenarios\\_build\\run_all.py flows\\ΦΜΥ.py scenarios\\ΦΜΥ"""
import json, copy, sys, os

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ΦΜΥ")
AFM = "123456789"
E1, E2 = "111111111", "222222222"


def jl10(employer=AFM, emps=((E1, 100000, 5000), (E2, 120000, 7000))):
    """JL10 lines: 0 header · 1 employer (ΑΦΜ at 36–45) · 2 totals · 3 one per employee
    (ΑΦΜ 1–10, gross 56–67, ΦΜΥ 98–108; amounts in cents)."""
    lines = ["0JL10" + " " * 60, ("1" + " " * 35 + employer).ljust(80), "2".ljust(80)]
    for a, gross, tax in emps:
        ln = list(("3" + a).ljust(120))
        ln[56:67] = list(str(gross).zfill(11))
        ln[98:108] = list(str(tax).zfill(10))
        lines.append("".join(ln))
    return lines


def excel(rows=((E1, 1000, 50.0, 0), (E2, 1200, 70.0, 0)), total=None):
    """The Μισθοδοτική's Excel (as_lists): header with «Φ.Μ.Υ», one row per employee, a totals row."""
    out = [["Μισθοδοτική Κατάσταση"], ["Εταιρεία", "Επώνυμο", "ΑΦΜ", "Συν.Αποδ.", "Φ.Μ.Υ", "Ψηφ.Τέλος", "Έτος"]]
    for a, gross, tax, psif in rows:
        out.append(["ΤΕΣΤ ΕΤΑΙΡΙΑ", "ΕΡΓΑΖΟΜΕΝΟΣ " + a[:1], a, gross, tax, psif, 2026])
    t = total or (sum(r[1] for r in rows), sum(r[2] for r in rows), sum(r[3] for r in rows))
    out.append(["", "Σύνολα", "", t[0], t[1], t[2], ""])
    return out


W = lambda cls, title=None, results=(True,): {k: v for k, v in
                                               {"class_name": cls, "title": title, "results": list(results)}.items()
                                               if v is not None}

BASE = {
    "description": "",
    "vars": {"period.year": "2026", "period.month": "9", "client.afm": AFM,
             "client.name": "ΤΕΣΤ ΕΤΑΙΡΙΑ", "output.folder": "C:\\Myrmex\\exports\\ΤΕΣΤ ΕΤΑΙΡΙΑ - " + AFM},
    "windows": [
        W("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση", [False, True]),   # opens at the 1st pick (4 rows)
        W("#32770", "Αποθήκευση"),
        W("TfJL10Edit", "Αποθηκευμένα στοιχεία"),
        W("#32770", "Πληροφορία"),
        W("#32770", "Εκτύπωση"),
        W("#32770", "Αποθήκευση εξόδου εκτύπωσης ως"),
    ],
    "store": {
        "company_search": AFM,
        "select_all_runs": [[10, 60]],
        "exported_file": {"found": True, "name": "export.xlsx", "path": "C:\\Myrmex\\exports\\ΤΕΣΤ ΕΤΑΙΡΙΑ - " + AFM + "\\export.xlsx"},
        "fmy_amount": "120,00",
        "digital_signature": "0,00",
        "payroll_rows": excel(),
        "jl10_lines": jl10(),
    },
    "fail": [],
}

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


DONE = lambda text, **kw: {"status": "success", "message_contains": text, **kw}
FAILED = lambda text, retry=None: {"status": "failed", "message_contains": text,
                                   **({} if retry is None else {"retryable": retry})}
HALT = lambda text: {"status": "halted", "message_contains": text}
SUBMIT = ["Υποβολή"]

# ── the normal run and the submit / test switches ──
sc("normal_month", "Κανονικός μήνας, τοπικά (χωρίς submit): ΦΜΥ οθόνης = Excel = JL10 → υποβολή, PDF «Ταυτότητα Οφειλής»",
   expect=DONE("ΥΠΟΒΛΗΘΗΚΕ + PDF"))
sc("submit_amount_due", "Υποβολή με ΦΜΥ προς πληρωμή (120,00): «Ταυτότητα Οφειλής» → Εκτύπωση → PDF «ΤΕΣΤ ΕΤΑΙΡΙΑ ΦΜΥ 09 2026»",
   vars={"submit": "yes"}, expect=DONE("ΤΕΣΤ ΕΤΑΙΡΙΑ ΦΜΥ 09 2026"))
sc("submit_zero_return", "Μηδενική (ΦΜΥ 0 και Ψηφ.Τέλος 0): window.print() → PDF «ΤΕΣΤ ΕΤΑΙΡΙΑ ΜΗΔΕΝΙΚΟ ΦΜΥ 09 2026»",
   vars={"submit": "yes"}, store={"fmy_amount": "0,00", "payroll_rows": excel(((E1, 0, 0, 0),)),
                                  "jl10_lines": jl10(emps=())},
   expect=DONE("ΜΗΔΕΝΙΚΟ ΦΜΥ 09 2026"))
sc("zero_print_fallback", "Μηδενική: η «Εκτύπωση» δεν ανοίγει ποτέ → PDF από φωτογραφία της σελίδας",
   vars={"submit": "yes"}, store={"fmy_amount": "0,00", "payroll_rows": excel(((E1, 0, 0, 0),)),
                                  "jl10_lines": jl10(emps=())},
   drop_windows=[("#32770", "Εκτύπωση")], expect=DONE("PDF φωτογραφία"))
sc("cloud_submit_yes", "cloud submit=yes → υποβάλλει", vars={"submit": "yes"}, expect=DONE("ΥΠΟΒΛΗΘΗΚΕ"))
sc("cloud_submit_no", "cloud submit=no → ΔΟΚΙΜΗ, σταματά πριν τη ΜΕΤΑΦΟΡΤΩΣΗ", vars={"submit": "no"},
   expect=DONE("ΔΟΚΙΜΗ — submit=no"))
sc("cloud_submit_other", "cloud submit=«κάτι άλλο» → ποτέ υποβολή, ΔΟΚΙΜΗ", vars={"submit": "maybe"},
   expect=DONE("ΔΟΚΙΜΗ — submit=maybe"))
sc("test_stops_before_upload", "test=yes → σταματά πριν τη ΜΕΤΑΦΟΡΤΩΣΗ, τίποτα στην ΑΑΔΕ", vars={"test": "yes"},
   expect=DONE("ΔΟΚΙΜΗ — test=yes"))
sc("test_wins_over_submit", "test=yes μαζί με submit=yes → κερδίζει το test", vars={"test": "yes", "submit": "yes"},
   expect=DONE("ΔΟΚΙΜΗ — test=yes"))

# ── Epsilon login / company ──
sc("login_error", "«Σφάλμα» της Epsilon στη σύνδεση (άλλοι χρήστες μέσα) → πρόβλημα",
   windows=[W("TMessageForm", "Σφάλμα")], expect=FAILED("«Σφάλμα» στη σύνδεση"))
sc("other_app_error_title", "Καρτέλα Chrome με «Σφάλμα» στον τίτλο (04/11): ΔΕΝ είναι μήνυμα της Epsilon → συνεχίζει",
   windows=[W("Chrome_WidgetWin_1", "Gmail — Σφάλμα με ΑΠΔ")], vars={"test": "yes"}, expect=DONE("ΔΟΚΙΜΗ"))
sc("wrong_company_search", "Η αναζήτηση εταιριών δεν γράφει το ΑΦΜ → σταματά πριν ανοίξει εταιρία",
   store={"company_search": "12345"}, expect=FAILED("η αναζήτηση εταιριών γράφει"))
sc("kad_question", "«Επιβεβαίωση» για ΚΑΔ στο άνοιγμα της εταιρίας → «Όχι», φωτογραφία, συνέχεια",
   windows=[W("TMessageForm", "Επιβεβαίωση", [True, True, False])], vars={"test": "yes"},
   expect=DONE("ΔΟΚΙΜΗ"))

# ── Μισθοδοτική → Excel ──
sc("payroll_three_rows", "Η «Μισθοδοτική Κατάσταση» ανοίγει με 3 γραμμές (LENOVO-PC), όχι με 4",
   windows=[W("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση", [False, False, True])],
   drop_windows=[("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση")], vars={"test": "yes"}, expect=DONE("ΔΟΚΙΜΗ"))
sc("payroll_unknown_rows", "Καμία γνωστή γραμμή (4/3/5/6) δεν ανοίγει τη Μισθοδοτική → πρόβλημα, προσθήκη στο PAYROLL_ROWS",
   drop_windows=[("TfMisthodotikiKatastasi", "Μισθοδοτική Κατάσταση")], expect=FAILED("PAYROLL_ROWS"))
sc("no_records", "Δεν βρίσκεται το «Επιλογή Όλων» → 0 ΕΓΓΡΑΦΕΣ (επιτυχία, δεν υποβλήθηκε)",
   store={"select_all_runs": []}, expect=DONE("0 ΕΓΓΡΑΦΕΣ"))
sc("no_employees", "«Πληροφορία» αντί για «Αποθήκευση» στο «Εξαγωγή σε EXCEL» → ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ",
   windows=[W("TMessageForm", "Πληροφορία")], drop_windows=[("#32770", "Αποθήκευση")],
   expect=DONE("ΧΩΡΙΣ ΕΡΓΑΖΟΜΕΝΟΥΣ"))
sc("empty_fields", "Excel με κενά κελιά ΦΜΥ / Ψηφ.Τέλος (= 0) και οθόνη 0 → μηδενική",
   vars={"test": "yes"}, store={"fmy_amount": "", "digital_signature": "",
                                "payroll_rows": excel(((E1, 500, None, None),), total=(500, 0, 0)),
                                "jl10_lines": jl10(emps=())},
   expect=DONE("ΔΟΚΙΜΗ"))
sc("excel_no_totals", "Το Excel χωρίς γραμμή συνόλων (κενό σύνολο ΦΜΥ): ένα μη διαβασμένο σύνολο δεν γίνεται 0 → σταματά",
   store={"payroll_rows": excel(((E1, 500, 10.0, 0),), total=(500, None, ""))},
   expect=FAILED("δεν βρέθηκε γραμμή συνόλων"))

# ── comparisons ──
sc("mismatch_fmy", "ΦΜΥ οθόνης ≠ Excel → ΑΣΥΜΦΩΝΙΑ (πρόβλημα, δεν ξανατρέχει)",
   store={"fmy_amount": "125,00"}, expect=FAILED("ΑΣΥΜΦΩΝΙΑ Φ.Μ.Υ", retry=False))
sc("mismatch_psif", "Ψηφ.Τέλος οθόνης ≠ Excel → ΑΣΥΜΦΩΝΙΑ",
   store={"digital_signature": "3,60"}, expect=FAILED("ΑΣΥΜΦΩΝΙΑ Ψηφ.Τέλους", retry=False))
sc("jl10_mismatch", "ΦΜΥ ενός εργαζόμενου στο JL10 ≠ Excel → ΑΣΥΜΦΩΝΙΑ JL10",
   store={"jl10_lines": jl10(emps=((E1, 100000, 5000), (E2, 120000, 6900)))},
   expect=FAILED("ΑΣΥΜΦΩΝΙΑ JL10", retry=False))
sc("jl10_other_employer", "Το JL10 είναι άλλου εργοδότη → ΑΣΥΜΦΩΝΙΑ JL10",
   store={"jl10_lines": jl10(employer="999999999")}, expect=FAILED("ΑΦΜ εργοδότη 999999999", retry=False))
sc("jl10_other_employee", "Εργαζόμενος μόνο στο JL10 → ΑΣΥΜΦΩΝΙΑ JL10",
   store={"jl10_lines": jl10(emps=((E1, 100000, 5000), ("333333333", 120000, 7000)))},
   expect=FAILED("εργαζόμενοι μόνο στο αρχείο", retry=False))
sc("jl10_zero_tax_intern", "Ασκούμενος μόνο στο Excel με 0 € και 0 ΦΜΥ (11/11): η Epsilon τον αφήνει έξω από το JL10 → ΟΚ",
   vars={"test": "yes"}, store={"payroll_rows": excel(((E1, 1000, 50.0, 0), (E2, 1200, 70.0, 0), ("444444444", 0, 0, 0)))},
   expect=DONE("ΔΟΚΙΜΗ"))
sc("jl10_not_written", "Το JL10 δεν γράφτηκε μετά την «Εξαγωγή» → σταματά",
   fail=[{"action": "wait_for_file", "error": "JL10 not newer"}], expect=FAILED("JL10 not newer"))

# ── calculation / declaration options ──
sc("calc_fallback", "Τα «Αποθηκευμένα στοιχεία» δεν ανοίγουν σε 120″ → «Υπολογισμός» ξανά με UIA, μετά ανοίγουν",
   windows=[W("TfJL10Edit", "Αποθηκευμένα στοιχεία", [False, True])],
   drop_windows=[("TfJL10Edit", "Αποθηκευμένα στοιχεία")], vars={"test": "yes"}, expect=DONE("ΔΟΚΙΜΗ"))
sc("declaration_amended", "fmy.amended=yes → ένα κλικ στο «Τροποποιητική», τα άλλα δύο ανέγγιχτα",
   vars={"test": "yes", "fmy.amended": "yes", "fmy.total_pause": "no", "fmy.temp_absence": "no"},
   store={"jl10_option_fmy.amended": True}, expect=DONE("ΔΟΚΙΜΗ"))
sc("declaration_box_missing", "fmy.total_pause=yes αλλά το κουτί δεν βρίσκεται → σταματά πριν τον υπολογισμό",
   vars={"fmy.total_pause": "yes"}, store={"jl10_option_fmy.total_pause": False},
   expect=FAILED("Ολική παύση απασχόλησης προσωπικού"))
sc("declaration_stop_before_calc", "stop.before_calc=yes → ΣΤΟΠ πριν τον «Υπολογισμό», όλα ανοιχτά (halt)",
   vars={"stop.before_calc": "yes", "fmy.temp_absence": "yes"}, store={"jl10_option_fmy.temp_absence": True},
   expect=HALT("ΣΤΟΠ ΠΡΙΝ ΤΟΝ «Υπολογισμό»"))

# ── submission ──
sc("bad_taxis_login", "Λάθος κωδικοί TAXISnet: ο Firefox γράφει «λανθασμένο όνομα χρήστη» → σταματά αμέσως, δεν ξανατρέχει",
   drop_windows=[("#32770", "Πληροφορία")], store={"login_page": {"found": "true"}},
   expect=FAILED("ΛΑΘΟΣ ΚΩΔΙΚΟΙ TAXISnet", retry=False))
sc("no_uia", "Ο Firefox χωρίς UI Automation → ΣΤΟΠ ΠΡΙΝ ΤΗ ΜΕΤΑΦΟΡΤΩΣΗ",
   fail=[{"action": "browser_focus_uia", "error": "no UIA elements"}], expect=FAILED("ΣΤΟΠ ΠΡΙΝ ΤΗ ΜΕΤΑΦΟΡΤΩΣΗ"))
sc("already_submitted", "«Σφάλμα» μετά το «Υποβολή Αρχείου» (δήλωση) → ΗΔΗ ΥΠΟΒΕΒΛΗΜΕΝΟΣ",
   windows=[W("#32770", "Σφάλμα", [False, True])], expect=DONE("ΗΔΗ ΥΠΟΒΕΒΛΗΜΕΝΟΣ"))
sc("confirm_after_info", "LENOVO-PC: η «Επιβεβαίωση … Υποβλήθηκε;» έρχεται ΜΕΤΑ την «Πληροφορία» → «Όχι»",
   windows=[W("TMessageForm", "Επιβεβαίωση", [False, False, True, True, False])], vars={"test": "yes"},
   expect=DONE("ΔΟΚΙΜΗ"))
sc("pdf_missing", "Υποβλήθηκε αλλά το PDF δεν γράφτηκε → σταματά (ήδη committed: δεν ξανατρέχει)",
   fail=[{"action": "wait_for_file", "error": "no PDF"}], expect=FAILED("no PDF"))

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for name, s in SCEN.items():
        s = {k: v for k, v in s.items() if v not in ([], None)}
        with open(os.path.join(OUT, name + ".json"), "w", encoding="utf-8", newline="\n") as fh:
            json.dump(s, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
    print("written:", len(SCEN))

