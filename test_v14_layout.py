"""Self-check for the v14 UI/UX fixes. Each assertion pins a defect that was measured
before the fix, so a layout regression fails here instead of on screen.
Run: python test_v14_layout.py
"""
import tkinter as tk

import PR_Calculator_GUI_v14 as v14


def settle(root, passes=6):
    """Drain the Tk event queue, including the after_idle card re-fit."""
    for _ in range(passes):
        root.update_idletasks()
        root.update()


def main():
    root = tk.Tk()
    app = v14.PRCalculatorGUI(root)
    settle(root)
    try:
        # The log console used to be pushed off-screen entirely (mapped == 0).
        assert app.log_widget.winfo_ismapped(), "live log console is not on screen"
        assert app.log_widget.winfo_height() >= 80, \
            f"log console squeezed to {app.log_widget.winfo_height()}px"

        # The whole layout must fit the window it opens at, or something gets clipped.
        assert root.winfo_reqheight() <= root.winfo_height(), (
            f"layout needs {root.winfo_reqheight()}px but the window is "
            f"{root.winfo_height()}px - something is cut off")

        # All 12 months visible: Dicembre used to be sliced in half by a card 17px short.
        tree = app.pvsyst_tree
        box = tree.bbox("m12")
        assert box, "Dicembre row is not rendered at all"
        assert box[1] + box[3] <= tree.winfo_height(), (
            f"Dicembre row ends at {box[1] + box[3]}px but the tree is "
            f"{tree.winfo_height()}px tall - last month is clipped")

        # RoundedCard must fit its content rather than staying short.
        card = app.metrics_card
        assert card.winfo_height() >= card.winfo_reqheight(), (
            f"card is {card.winfo_height()}px for {card.winfo_reqheight()}px of content")

        # Da/A row appears only for "Intervallo" (it used to render clipped, always).
        for scope, want in (("mese", False), ("giorno", False), ("intervallo", True)):
            app.day_scope_var.set(scope)
            app._on_day_scope_change()
            settle(root)
            assert app.day_range_frame.winfo_ismapped() == want, \
                f"day range row visibility wrong for scope={scope}"
        assert app.spin_day_from.winfo_width() >= 40, \
            f"Dal giorno spinbox clipped to {app.spin_day_from.winfo_width()}px"

        # VCOM days row appears only for "Misto" - startup used to leave it showing.
        for src, want in (("scada", False), ("vcom", False), ("misto", True)):
            app.data_source_var.set(src)
            app._on_data_source_change()
            settle(root)
            assert app.vcom_days_frame.winfo_ismapped() == want, \
                f"VCOM days row visibility wrong for source={src}"

        # Progress bar: hidden while idle, shown and positioned while running.
        assert not app.progress.winfo_ismapped(), "progress bar visible while idle"
        app._set_progress(3, 12)
        settle(root)
        assert app.progress.winfo_ismapped(), "progress bar did not appear"
        assert abs(app.progress.cget("value") - 25.0) < 1e-9, "progress value wrong"
        app._set_progress(0, 0)          # a zero total must not raise
        app._reset_run_buttons()
        settle(root)
        assert not app.progress.winfo_ismapped(), "progress bar not hidden when idle again"
    finally:
        root.destroy()
    print("v14 layout self-check OK")


if __name__ == "__main__":
    main()
