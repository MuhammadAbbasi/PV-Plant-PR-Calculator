import numpy as np

def repair_meter_series(raw_meter, left_anchor=0.0):
    """
    Detects missing or resetting meter values across 96 intervals of a day
    and applies linear interpolation using nearest valid neighbors.
    """
    n = len(raw_meter)
    series = np.array(raw_meter, dtype=float)
    bad = np.zeros(n, dtype=bool)

    # Step 1: Detect non-increasing / zero / jump drops
    for i in range(n):
        prev_val = left_anchor if i == 0 else series[i-1]
        curr_val = series[i]
        if curr_val <= 0 or curr_val < prev_val - 0.001 or (prev_val > 0 and curr_val < prev_val):
            bad[i] = True

    # Step 2: Interpolate bad intervals
    for i in range(n):
        if bad[i]:
            left_idx = i - 1
            while left_idx >= 0 and bad[left_idx]:
                left_idx -= 1
            left_v = left_anchor if left_idx < 0 else series[left_idx]

            right_idx = i + 1
            while right_idx < n and bad[right_idx]:
                right_idx += 1
            right_v = series[right_idx] if right_idx < n else left_v

            step_count = right_idx - left_idx
            if step_count > 0:
                step_val = (right_v - left_v) / float(step_count)
                series[i] = left_v + step_val * (i - left_idx)
            else:
                series[i] = left_v

    return series, bad
