"""Madre column resolution: the layout must survive repeated syncs unchanged.

Regression guard for the bug where the fixed-position column checks re-inserted
PR VCOM / PR Compensated / External Availability on a file that already had them,
shifting the Energy Loss block three columns right -- after which the hardcoded
summary formulas wrote PR Compensated values into the TX3 - Energy Loss column.

Run: python test_mother_columns.py
"""
from PR_Calculator_GUI_v15 import (MOTHER_CANONICAL, MOTHER_DAILY_ADDR, MOTHER_HEADERS,
                                   check_mother_layout, ensure_mother_columns, mother_col_key)


class FakeSheet:
    """Minimum COM surface ensure_mother_columns() touches."""

    def __init__(self, headers):
        self.h = [None] + list(headers)  # 1-indexed

    class _Cell:
        def __init__(self, sheet, col): self.s, self.c = sheet, col
        @property
        def Value(self): return self.s.h[self.c] if self.c < len(self.s.h) else None
        @Value.setter
        def Value(self, v): self.s.h[self.c] = v
        def __getattr__(self, name): raise AttributeError(name)  # styling -> caught

    def Cells(self, row, col):
        assert row == 4
        while len(self.h) <= col:
            self.h.append(None)
        return self._Cell(self, col)

    def Columns(self, at):
        sheet = self
        class _C:
            def Insert(self): sheet.h.insert(at, None)
        return _C()

    def headers(self):
        return [x for x in self.h[1:] if x is not None]


PRISTINE = ["Dati", "Irradiance Conditional MAX\n[KWh/m2]", "Energy (day)", "PR Total",
            "PR SCADA", "TX1 - Energy Loss\nkW/H", "TX2 - Energy Loss\nkW/H",
            "TX3 - Energy Loss\nkW/H"] + [f"PR TX{t}-INV-{i}" for t in (1, 2, 3) for i in range(1, 13)]

PRODUCTION = ["Dati", "Irradiance TX1", "Irradiance TX3", "Irradiance Conditional MAX\n[KWh/m2]",
              "Energy (day)", "PR Total", "PR SCADA", "PR VCOM", "PR Compensated",
              "External Availability\n[%]", "TX1 - Energy Loss\nkW/H", "TX2 - Energy Loss\nkW/H",
              "TX3 - Energy Loss\nkW/H"] + [f"PR TX{t}-INV-{i}" for t in (1, 2, 3) for i in range(1, 13)]


def check(name, headers, expect_stable):
    ws = FakeSheet(headers)
    cols = ensure_mother_columns(ws)
    after = ws.headers()

    assert all(k in cols for k in MOTHER_CANONICAL), f"{name}: unresolved {set(MOTHER_CANONICAL) - set(cols)}"
    if expect_stable:
        assert after == list(headers), f"{name}: layout changed\n  was {headers}\n  now {after}"

    # No duplicated logical column.
    keys = [mother_col_key(h) for h in after]
    real = [k for k in keys if k]
    assert len(real) == len(set(real)), f"{name}: duplicate columns {[k for k in real if real.count(k) > 1]}"

    # The thing that actually broke: each loss column must resolve to its own
    # daily-sheet loss cell, never to a PR cell.
    for key, addr in (("loss_tx1", "$AA$111"), ("loss_tx2", "$AN$111"), ("loss_tx3", "$BA$111")):
        assert MOTHER_DAILY_ADDR[key] == addr
        assert mother_col_key(after[cols[key] - 1]) == key, f"{name}: {key} resolved to the wrong column"
    assert cols["loss_tx1"] < cols["loss_tx2"] < cols["loss_tx3"] < cols["pr_inv_1_1"], f"{name}: block order"

    # Re-running a sync must be a no-op.
    ws2 = FakeSheet(after)
    cols2 = ensure_mother_columns(ws2)
    assert ws2.headers() == after, f"{name}: second sync shifted columns\n  {after}\n  {ws2.headers()}"
    assert cols2 == cols, f"{name}: second sync remapped columns"
    print(f"  ok  {name}: loss cols at {cols['loss_tx1']},{cols['loss_tx2']},{cols['loss_tx3']}; "
          f"inverters start at {cols['pr_inv_1_1']}")
    return cols


def main():
    print("mother column resolution")
    check("production Madre (already complete)", PRODUCTION, expect_stable=True)
    cols = check("pristine template (original_format)", PRISTINE, expect_stable=False)
    assert cols == {k: i for k, i in zip(MOTHER_CANONICAL, range(2, 14))} | \
           {f"pr_inv_{t}_{i}": 13 + (t - 1) * 12 + i for t in (1, 2, 3) for i in range(1, 13)}, \
           f"pristine template did not migrate to the canonical layout: {cols}"

    # Header classification: "TX3 - Energy Loss" must never fall through to "energy".
    assert mother_col_key("TX3 - Energy Loss\nkW/H") == "loss_tx3"
    assert mother_col_key("Perdita Energia TX2") == "loss_tx2"
    assert mother_col_key("Energy (day)") == "energy"
    assert mother_col_key("PR VCOM") == "pr_vcom"
    assert mother_col_key("PR Total") == "pr_total"
    assert mother_col_key("PR TX3-INV-12") == "pr_inv_3_12"
    assert mother_col_key("PR TX1-INV-10") == "pr_inv_1_10"
    assert mother_col_key("") is None and mother_col_key(None) is None
    print("  ok  header classification")

    # A clean file passes the layout check.
    dupes, order_ok = check_mother_layout(FakeSheet(PRODUCTION))
    assert not dupes and order_ok, f"production Madre flagged: {dupes} {order_ok}"

    # What the old positional inserts produced: duplicate PR VCOM / PR Compensated /
    # External Availability, with the Energy Loss block pushed three columns right.
    wrecked = (PRODUCTION[:4] + ["Meter Reading\n[MWh]", "Energy (day)", "PR Total", "PR VCOM",
               "PR Compensated", "External Availability\n[%]", "PR SCADA", "PR VCOM",
               "PR Compensated", "External Availability\n[%]"] + PRODUCTION[10:])
    dupes, _ = check_mother_layout(FakeSheet(wrecked))
    assert {k for k, _, _ in dupes} == {"pr_vcom", "pr_comp", "ext_avail"}, f"missed duplicates: {dupes}"
    assert all(MOTHER_HEADERS[k] for k, _, _ in dupes)  # every dupe is reportable
    print(f"  ok  wrecked layout rejected ({len(dupes)} duplicate columns)")

    # Out of canonical order but not duplicated: formulas still resolve, warn only.
    swapped = PRODUCTION[:5] + ["PR Total", "PR VCOM", "PR SCADA"] + PRODUCTION[8:]
    dupes, order_ok = check_mother_layout(FakeSheet(swapped))
    assert not dupes and not order_ok, f"order check missed the swap: {dupes} {order_ok}"
    print("  ok  non-canonical order detected")
    print("PASS")


if __name__ == "__main__":
    main()
