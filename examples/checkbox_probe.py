"""
checkbox_probe.py — Does MYRMEX read the «Επιλογές Δήλωσης» boxes (TwwCheckBox) right? (SDK 1.5)

With «Μηνιαίες Βεβαιώσεις Αποδοχών» open in Epsilon (the window TfJL10), tick some of the
boxes BY HAND, then:

    C:\\Myrmex\\myrmex.exe examples\\checkbox_probe.py

Nothing is clicked. For every box it writes ticked / not ticked and how dark its box is,
e.g. in myrmex.log:

    checkbox_state: checked (picture, ink 23) — TwwCheckBox 'Τροποποιητική'
    checkbox_state: unchecked (picture, ink 0) — TwwCheckBox 'Ολική παύση απασχόλησης προσωπικού'

Compare with the screen. A wrong answer → send Lefteris the log (myrmex.log) and a screenshot.
"""

MIN_SDK_VERSION = "1.5"

BOXES = ("Τροποποιητική", "Ολική παύση απασχόλησης προσωπικού", "Προσωρινή απουσία προσωπικού", "Αμοιβές Επ.ΙΚΑ")


def run(f, ctx):
    if not f.window_exists("TfJL10", timeout=10):
        return f.fail("δεν βρέθηκε η «Μηνιαίες Βεβαιώσεις Αποδοχών» (TfJL10) — ανοίξτε την και ξανατρέξτε")
    f.connect_window(window_class="TfJL10", timeout=5)
    states = []
    for caption in BOXES:
        if not f.control_exists("TwwCheckBox", caption, exact=False):
            states.append(f"{caption}: δεν βρέθηκε")
            continue
        ticked = f.checkbox_state(caption, class_name="TwwCheckBox")
        states.append(f"{caption}: {'ΝΑΙ' if ticked else 'όχι'}")
    f.screenshot(store_as="checkbox_probe_shot")
    return f.done(" · ".join(states))
