# Python flows — γρήγορη αρχή (myrmex.exe)

Το `myrmex.exe` είναι αυτό του εγκατεστημένου agent (`C:\Myrmex\`), σε κάθε PC
γραφείου και στο δικό σου: χωρίς git και χωρίς Python. Από τον φάκελο των flows:

```bat
:: έλεγχος χωρίς Epsilon: lint, import και όλη η ροή με mock
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --dry-run
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --dry-run --scenario scenarios\ΦΜΥ\no_employees.json

:: ένας πελάτης, ΔΟΚΙΜΗ: σταματά πριν τη ΜΕΤΑΦΟΡΤΩΣΗ, τίποτα δεν φτάνει στην ΑΑΔΕ
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --var test=yes --var period.year=2026 --var period.month=9 --var client.afm=123456789 --var client.name="ΕΤΑΙΡΙΑ" --var output.folder=C:\Myrmex\exports\ΕΤΑΙΡΙΑ

:: batch
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --clients clients.csv --output-root C:\Myrmex\exports --var period.year=2026 --var period.month=9

:: συνέχιση από φάση
C:\Myrmex\myrmex.exe flows\ΦΜΥ.py --from export_excel --var ...
```

Τοπικά δεν χρειάζεται υπογραφή. Υπογραφή μόνο για ανέβασμα στο cloud (`CLAUDE.md`).

- [CLAUDE.md](CLAUDE.md): οι κανόνες, τι διαφέρει από PC σε PC, δοκιμές, υπογραφή.
- `FLOW_SDK_REFERENCE.md`: όλες οι εντολές `f.*` (έρχεται με κάθε έκδοση του agent).
