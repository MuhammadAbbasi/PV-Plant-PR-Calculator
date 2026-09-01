"""Self-check for the v13 additions: day-scope resolution, the batch day filter,
the configurable degradation and the settings round-trip. Run: python test_v13_scope.py
"""
import os
import sys
import tempfile

import PR_Calculator_GUI_v13 as v13


def test_settings_roundtrip():
    tmp = os.path.join(tempfile.mkdtemp(), "settings.json")
    original = v13.get_settings_path
    v13.get_settings_path = lambda: tmp
    try:
        assert v13.load_settings() == v13.DEFAULT_SETTINGS, "missing file must give defaults"

        cfg = v13.load_settings()
        cfg["deg_rate"] = 0.006
        cfg["day_scope"] = "intervallo"
        cfg["pvsyst_monthly"]["5"] = 0.777
        cfg["bogus_key"] = "dropped"
        assert v13.save_settings(cfg)

        back = v13.load_settings()
        assert back["deg_rate"] == 0.006
        assert back["day_scope"] == "intervallo"
        assert back["pvsyst_monthly"]["5"] == 0.777
        assert back["pvsyst_monthly"]["1"] == 0.904, "unlisted months keep their defaults"
        assert "bogus_key" not in back, "unknown keys must not leak into the config"

        # A corrupt file must not take the app down.
        open(tmp, "w").write("{ not json")
        assert v13.load_settings()["deg_rate"] == 0.004
    finally:
        v13.get_settings_path = original


def test_day_scope_and_degradation():
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    app = v13.PRCalculatorGUI(root)          # also proves the layout still builds
    try:
        app.date_var.set("2026-05-17")

        app.day_scope_var.set("mese")
        assert app._resolve_day_filter() is None

        app.day_scope_var.set("giorno")
        assert app._resolve_day_filter() == {17}

        app.day_scope_var.set("intervallo")
        app.day_from_var.set("8")
        app.day_to_var.set("11")
        assert app._resolve_day_filter() == {8, 9, 10, 11}

        for bad_from, bad_to in (("11", "8"), ("0", "5"), ("1", "32"), ("x", "5")):
            app.day_from_var.set(bad_from)
            app.day_to_var.set(bad_to)
            try:
                app._resolve_day_filter()
            except ValueError:
                pass
            else:
                raise AssertionError(f"range {bad_from}-{bad_to} should be rejected")

        app.day_scope_var.set("giorno")
        app.date_var.set("non-una-data")
        try:
            app._resolve_day_filter()
        except ValueError:
            pass
        else:
            raise AssertionError("a malformed date must be rejected")

        # The filter narrows exactly the day folders the batch loop iterates over.
        available = ["01", "07", "08", "15", "31"]
        assert [d for d in available if int(d) in {8, 9, 10, 11}] == ["08"]

        # Degradation now reads the config: Feb-2025 start, 0.4%/yr, year 1 undegraded.
        assert app._pr_degradation_factor(2025, 6) == (0, 1.0)
        assert app._pr_degradation_factor(2026, 1)[0] == 0, "contract year runs Feb..Jan"
        assert app._pr_degradation_factor(2026, 2)[0] == 1
        n, factor = app._pr_degradation_factor(2027, 3)
        assert n == 2 and abs(factor - 0.996 ** 2) < 1e-12

        app.cfg["deg_rate"] = 0.01
        app.cfg["deg_start_year"] = 2026
        assert app._pr_degradation_factor(2028, 3) == (2, 0.99 ** 2), "options must reach the engine"

        # The advanced-options dialog must build and expose every knob.
        app.show_options()
        dialog = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]
        assert len(dialog) == 1 and dialog[0].title() == "Opzioni Avanzate"
        entries = _descendants(dialog[0], v13.ttk.Entry)
        assert len(entries) == 15, f"expected 3 degradation + 12 monthly fields, got {len(entries)}"
        assert len(_descendants(dialog[0], v13.ttk.Checkbutton)) == 2
        dialog[0].destroy()

        # _collect_settings must snapshot the live UI, not the stale loaded config.
        app.sync_mother_var.set(False)
        app.day_scope_var.set("giorno")
        snap = app._collect_settings()
        assert snap["sync_mother"] is False and snap["day_scope"] == "giorno"
        assert snap["pvsyst_monthly"]["1"] == app.pvsyst_monthly[1]
    finally:
        root.destroy()


def _descendants(widget, cls):
    found = [widget] if isinstance(widget, cls) else []
    for child in widget.winfo_children():
        found += _descendants(child, cls)
    return found


if __name__ == "__main__":
    test_settings_roundtrip()
    test_day_scope_and_degradation()
    print("v13 self-check OK")
