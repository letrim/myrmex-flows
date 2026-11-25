"""
dialog_text_probe.py — Does f.dialog_text read an Epsilon message? (αίτημα Ε)

With a message of the Epsilon on the screen (e.g. close Epsilon → «Έξοδος από
την εφαρμογή;», or any «Επιβεβαίωση» / «Πληροφορία»):

    C:\\Myrmex\\myrmex.exe examples\\dialog_text_probe.py

It answers nothing — the message stays open. It writes the title, the buttons and
the text. The text comes through Ctrl+C (Epsilon paints it): you'll hear a beep
and the message comes to the front. Expected, e.g.:

    dialog_text: TMessageForm/None -> 'Έξοδος από την εφαρμογή;' (Ctrl+C)

'' (κενό) means it couldn't be read: send Lefteris the log (myrmex.log).
"""

def run(f, ctx):
    if not f.window_exists("TMessageForm", timeout=30):
        return f.fail("δεν βρέθηκε μήνυμα (TMessageForm) σε 30″ — ανοίξτε ένα μήνυμα της Epsilon και ξανατρέξτε")
    buttons = f.dialog_buttons("TMessageForm")
    text = f.dialog_text("TMessageForm")
    f.log(f"κουμπιά: {buttons} — κείμενο: {text!r}")
    return f.done(f"ΚΕΙΜΕΝΟ: {text!r}" if text else "ΤΟ ΚΕΙΜΕΝΟ ΔΕΝ ΔΙΑΒΑΣΤΗΚΕ — στείλτε το myrmex.log")
