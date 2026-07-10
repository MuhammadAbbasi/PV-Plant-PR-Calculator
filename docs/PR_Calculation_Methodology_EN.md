# Performance Ratio (PR) Calculation — Methodology

**Plant:** Mazara 01 photovoltaic plant (GET S.R.L.) · **Document generated:** 2026-07-10 · **Last updated:** 2026-07-10

> This document explains, in plain language, how the Performance Ratio of the Mazara 01 plant is calculated. It is written for **both technical and non-technical readers**. Every technical term is defined the first time it appears, every data source is listed, every formula is broken down step by step, and worked examples show how the numbers come together. An identical Italian version is available at [`PR_Calculation_Methodology_IT.md`](PR_Calculation_Methodology_IT.md).

---

## Table of Contents
1. [Introduction to PR Calculation](#1-introduction-to-pr-calculation)
2. [Data Sources Used](#2-data-sources-used)
3. [Core Formulae with Annotations](#3-core-formulae-with-annotations)
4. [Step-by-Step Calculation Walkthrough](#4-step-by-step-calculation-walkthrough)
5. [Frequently Asked Questions](#5-frequently-asked-questions)
6. [Glossary](#6-glossary)

---

## 1. Introduction to PR Calculation

### What is Performance Ratio?

**Performance Ratio (PR)** is a single percentage that answers one question:

> *"Of all the energy the plant **could** have produced given the sunlight it received, how much did it **actually** produce?"*

A PR of **100%** would mean the plant produced exactly the ideal amount for that sunlight (no losses at all). Real plants always sit below 100% because of unavoidable physics (heat, wiring, inverter efficiency) and occasional problems (equipment down, grid limits). A healthy plant like Mazara 01 typically runs around **82–83%**.

PR is calculated as:

```
                Actual energy produced
PR (%)  =  ---------------------------------  × 100
            Energy the sunlight could support
```

- **Actual energy produced** — measured directly from the plant's meters and inverters.
- **Energy the sunlight could support** — the plant's nameplate size multiplied by how much sun (irradiance) actually fell on the panels.

### Why it matters

The plant operator has a **contractual guarantee** (Allegato 9.1) to keep PR above a target value each year. If the real PR falls below the guaranteed PR, a **penalty** is owed. Because money depends on it, the calculation must be transparent, repeatable, and defensible — which is the purpose of this document.

### The four PR figures you will see

The reports show more than one PR number, because there are different ways to look at performance:

| Name | Plain meaning |
|------|---------------|
| **Raw PR** (a.k.a. "PR Total") | Straight actual-vs-ideal, **no allowances**. Drops whenever the plant is down, even if the outage wasn't the operator's fault. |
| **SCADA PR** | The PR reported by the plant's own control system (SCADA), taken from the vendor's daily KPI export. |
| **VCOM PR** | The PR reported by the independent **VCOM** monitoring platform (a third-party check). |
| **Compensated PR** | Raw PR **plus** the energy that was provably lost to outages and grid curtailment **added back**. This isolates the plant's *technical* quality from events outside the equipment's control, and is the figure most relevant to the guarantee. |

> **SCADA** = *Supervisory Control And Data Acquisition*, the plant's on-site control and monitoring system.

---

## 2. Data Sources Used

Every calculation day reads a set of raw files (sampled every **15 minutes**, giving **96 intervals per day**). Below is every dataset used and what it contributes.

### 2.1 Raw measurement files (one folder per day)

| File | What it contains | Role in PR |
|------|------------------|-----------|
| `TS_01_Inverter_15Min.xlsx`<br>`TS_02_Inverter_15Min.xlsx`<br>`TS_03_Inverter_15Min.xlsx` | **Active power (kW)** of each inverter, for the three transformer stations **TX1, TX2, TX3** (12 inverters each = **36 inverters**). | Source of **actual energy** and of **downtime detection** (an inverter reading ≈0 while the sun shines is "down"). |
| `TS_01_Weather_15Min.xlsx`<br>`TS_03_Weather_15Min.xlsx` | **Irradiance (W/m²)** from the two **pyranometers** mounted in the plane of the panels (POA), at stations TX1 and TX3. | Source of **sunlight received** — the denominator of PR. |
| `SATAC_Meter_15Min.xlsx` | The **revenue energy meter** cumulative reading (kWh) at the grid connection point. | Independent measure of **energy actually exported**; also used for availability. |
| `Regolazione_della_potenza_attiva_YYYY_MM_DD.xlsx` | The **active-power regulation limit ratio** (0–1): how much the grid operator allowed the plant to produce. | Detects **curtailment** (grid-ordered reduction). |

> **Pyranometer** = an instrument that measures solar irradiance (sunlight power) in watts per square metre (W/m²).
> **POA** = *Plane Of Array*, i.e. irradiance measured on the same tilt/orientation as the panels — the sunlight the panels actually "see".
> **Curtailment** = the grid operator instructing the plant to produce **less** than it could, for network reasons.

### 2.2 External vendor PR files (optional, one per month)

If present in the month folder, these provide ready-made daily PR figures that are written into the monthly summary ("Mother") file:

| File | Format | Provides |
|------|--------|----------|
| `KPI_Report_Daily.xls` | Excel | **SCADA daily PR (%)** from the plant control system. |
| `Performance_ratio_vcom.csv` | UTF-16, tab-separated | **VCOM daily PR (%)** from the independent monitoring platform. |

### 2.3 Design inputs and fixed constants

| Input | Value | Meaning |
|-------|-------|---------|
| **Total nameplate power** (P_nominal) | **12,625 kWp** | The sum of the rated DC power of all 36 inverters' arrays. "kWp" = kilowatt-peak, the power under standard test sunlight. |
| **Per-inverter DC power** | 328.125 – 359.375 kWp | Each inverter's array size (varies slightly by inverter). |
| **Inverter AC cap** | 320 kW (usable ≈ 280.3 kW) | The most an inverter can push to the grid; the usable factor 0.876 reflects real limits. |
| **Minimum irradiance threshold** | **50 W/m²** | Below this weak sunlight, intervals are ignored (readings are too noisy to be meaningful). |
| **POA deviation tolerance** | **10%** | How far the two pyranometers may disagree before the higher one is trusted (see §3.1). |
| **PVSyst monthly target PR** | 82.0% – 90.4% (per month) | The design-expected PR for each calendar month, from the PVSyst simulation. Used as the reference/target and inside loss estimates. |
| **Annual degradation** | **0.4% per year** | The guaranteed PR target is reduced 0.4% each year (compounding) as panels age — see §3.8. |
| **Sampling interval** | 15 minutes = 0.25 h | Each reading represents a quarter-hour of operation. |

**PVSyst monthly target PR baselines (Year 1):**

| Jan | Feb | Mar | Apr | May | Jun | Jul | Aug | Sep | Oct | Nov | Dec |
|----|----|----|----|----|----|----|----|----|----|----|----|
|90.4|89.6|89.7|86.8|83.2|83.3|82.0|82.8|85.2|87.6|89.4|90.0|

> **PVSyst** = an industry-standard software used to simulate the expected output of a solar plant. Its monthly PR values are the "design target".

---

## 3. Core Formulae with Annotations

This section builds the calculation from the ground up. Each formula is followed by an explanation of every component.

### 3.1 Step 1 — Turn sunlight into a "reference irradiance" (Column I)

Irradiance arrives in **W/m²** (an instantaneous power). To compare it with energy, we convert each 15-minute reading into **kWh/m²** (an energy density):

```
irradiance_kWh/m²  =  irradiance_W/m²  ÷  4000
```

- **÷ 1000** converts watts to kilowatts.
- **÷ 4** converts a full hour to a quarter-hour (15 minutes).
- Combined: **÷ 4000**.

There are **two** pyranometers (TX1 and TX3). For each interval we must pick **one** reference value from the two. The rule (the "**Conditional MAX**" method, the default) is:

```
if both sensors read 0:          reference = 0
else if exactly one reads 0:     reference = the working sensor
else if |POA1 − POA3| / average  > 10% (tolerance):   reference = the HIGHER sensor
else:                            reference = the average of the two
```

- **Why prefer the higher sensor when they disagree?** A pyranometer that is dirty or shaded reads **too low**. Averaging a faulty low reading would understate the sunlight and *inflate* PR. Trusting the higher (clean) sensor is the conservative choice.
- An alternative "**Average**" method (plain mean of the two) is also selectable in the tool; the plant uses one method consistently.

Finally, a **minimum-sunlight gate** is applied:

```
if reference × 4000  <  50 W/m²:   reference = 0   (interval discarded)
```

The surviving per-interval values form **Column I** in the daily file. Their daily sum is the key denominator driver:

```
Σ I  =  sum of Column I over all 96 intervals   (units: kWh/m²)
```

> **Column I / Σ I** is the total useful sunlight energy density for the day — the "fuel" the plant had to work with.

### 3.2 Step 2 — Measure the energy actually produced

Energy is captured two ways:

**(a) From the inverters** — each inverter's power (kW) over a 15-minute interval becomes energy:

```
inverter_energy_kWh  =  inverter_power_kW  ×  0.25       (0.25 h = 15 min)
```

Summed over all 36 inverters and all 96 intervals gives the day's **inverter energy**.

**(b) From the meter** — the difference between consecutive cumulative meter readings:

```
meter_energy_kWh  =  (reading_now − reading_before)  ×  1000
```

> Missing, zero, or "backwards" meter readings (a cumulative meter can only go up) are automatically **repaired** by interpolating between the nearest valid readings, and are flagged in orange in the daily file.

### 3.3 Step 3 — Raw (uncompensated) PR

The theoretical energy the sunlight could support is:

```
Expected energy  =  P_nominal  ×  Σ I  =  12,625 kWp  ×  Σ I (kWh/m²)
```

> Intuition: at standard sunlight (1000 W/m² = 1 kWh/m² per hour), each kWp of panel produces 1 kWh. So nameplate × sunlight-energy-density = the ideal energy.

```
                 Actual energy (inverters)
Raw PR (%)  =  -----------------------------  × 100
                    12,625  ×  Σ I
```

This is honest but **unforgiving**: if the plant was switched off for hours, the actual energy is low and Raw PR plunges — even if the shutdown was ordered by the grid and not a fault of the equipment.

### 3.4 Step 4 — Recoverable losses

To judge the *equipment's* quality, we estimate the energy that was lost to events outside normal operation and **add it back**. Every 15-minute interval, for every inverter, **at most one** of the following applies (they are mutually exclusive):

**(a) Downtime loss** — the inverter is off (< 1 kW) while the sun is up:

```
if other inverters on the same transformer are still working:
      downtime_loss = (average power of the working inverters) × 0.25
else (the whole transformer is down):
      downtime_loss = (POA_avg ÷ 1000) × DC_inverter × PVSyst_target × 0.25
```

- If neighbours are running, they are the best estimate of what the dead inverter *would* have produced.
- If everything is down, we fall back to the physics estimate (sunlight × panel size × design PR).

**(b) Curtailment loss** — the inverter **is** producing (≥ 1 kW) but the grid limited it (limit ratio < 0.875):

```
curtailment_loss = max( 0,  min(expected, AC_cap) − AC_cap × limit_ratio )  × 0.25
```

- This captures the gap between what the inverter could have made and the reduced level the grid allowed.
- **Important:** curtailment is only counted when the inverter is actually producing. During a full outage the grid signal also reads ~0, but that energy is already captured as *downtime* — counting it again would double-count it.

**(c) Ramp / recovery loss** — the inverter is producing (≥ 1 kW), not curtailed, on an interval **immediately before or after a full plant outage**:

```
ramp_loss = max( 0,  expected_power − actual_power )  × 0.25
```

- When a plant trips and comes back, the recovery interval shows the inverter already ramping up but still below what the sunlight supported. This books that shortfall, which the plain downtime test (which needs a near-zero reading) would otherwise miss.

The three add up to the inverter's loss for that interval:

```
inverter_loss = downtime_loss + curtailment_loss + ramp_loss
```

> **expected_power** = (POA_avg ÷ 1000) × DC_inverter × PVSyst_target, capped at the inverter's usable AC limit (≈ 280.3 kW).

### 3.5 Step 5 — Compensated PR

```
                 Actual energy  +  Σ recoverable losses
Compensated PR (%) =  ---------------------------------------  × 100
                            12,625  ×  Σ I
```

By adding the recoverable losses back into the numerator, Compensated PR reflects how the plant performs **when it is allowed to run** — the fairest measure of technical quality, and the one aligned with the contractual guarantee.

> **Sanity rule:** a correct Compensated PR can never meaningfully exceed **100%** — you cannot recover more than the sunlight could support. A value above 100% signals a data or double-counting problem (see FAQ).

### 3.6 Per-inverter PR

The same idea applied to a single inverter, to spot under-performers:

```
                       inverter_energy  +  inverter_loss
Per-inverter PR (%) =  -----------------------------------  × 100
                            DC_inverter  ×  Σ I
```

### 3.7 External availability

How much of the potential energy was actually delivered, expressed as a percentage:

```
                        E
Availability (%) =  -----------------  × 100
                     E + L1 + L2 + L3
```

- **E** = metered energy for the day.
- **L1, L2, L3** = total losses attributed to transformers TX1, TX2, TX3.
- A day with no outages gives 100%; a heavy-outage day (like the worked example in §4.2) drops well below.

### 3.8 The target and its yearly degradation

Solar panels slowly lose efficiency with age. The contract (Allegato 9.1) reflects this by **reducing the guaranteed PR target by 0.4% each year**, compounding:

```
Target(month, year)  =  PVSyst_base(month)  ×  (1 − 0.4%) ^ n
```

- **PVSyst_base(month)** — the design target for that calendar month (table in §2.3).
- **n** — the number of completed contract years since the plant start (**February 2025**). Contract years run **February → January**.
  - **Year 1** (Feb 2025 – Jan 2026): n = 0 → no reduction.
  - **Year 2** (Feb 2026 – Jan 2027): n = 1 → factor 0.996.
  - **Year 3**: n = 2 → factor 0.996² = 0.99202, and so on.

The tool reads the year and month automatically from the report folder path (`…/YYYY MM/DD`) and applies the correct factor. The degraded target feeds both the displayed target and the physics-based loss estimates in §3.4.

---

## 4. Step-by-Step Calculation Walkthrough

### 4.1 A single 15-minute interval

Suppose at **15:15** the two pyranometers read **POA1 = 990 W/m²** and **POA3 = 968 W/m²**, and it is June 2026 (contract Year 2).

**1. Convert to energy density:**
```
POA1 = 990 ÷ 4000 = 0.24750 kWh/m²
POA3 = 968 ÷ 4000 = 0.24200 kWh/m²
```

**2. Pick the reference (Conditional MAX):**
```
average   = (0.24750 + 0.24200) / 2 = 0.24475
deviation = |0.24750 − 0.24200| / 0.24475 = 2.2%   →  below the 10% tolerance
reference (Column I) = average = 0.24475 kWh/m²
```
Check the gate: 0.24475 × 4000 = 979 W/m² ≥ 50 → **kept**.

**3. Degraded target for June 2026:**
```
n = 1  →  factor = 0.996
Target = 83.3% × 0.996 = 82.97%   (base June PVSyst = 0.833)
```

**4. If all 36 inverters were producing normally**, this interval simply contributes its inverter energy to the numerator and `0.24475` to Σ I. No loss is booked.

**5. If the whole plant were OFF this interval** (a full outage) with the same sunlight, the downtime loss for one 343.75 kWp inverter would be:
```
downtime_loss = (979 ÷ 1000) × 343.75 × 0.8297 × 0.25
              = 0.979 × 343.75 × 0.8297 × 0.25
              ≈ 69.8 kWh
```
Across all 36 inverters this recovers ≈ 2,500 kWh for that single interval — the energy the plant *would* have made.

### 4.2 A full day — a normal day vs. an outage day

**Normal day (e.g. 20 June 2026):** the plant runs all day.
- Σ I ≈ 10.13 kWh/m² · Actual energy ≈ 104,400 kWh · Losses ≈ 0.
- Raw PR ≈ Compensated PR ≈ **83.0%** · Availability = 100%.

**Outage day (12 June 2026):** the plant tripped for ~5 hours (21 of the sunlit intervals had all inverters at 0).
- Actual energy ≈ 46,900 kWh (about half a normal day) → **Raw PR ≈ 39.6%** (badly hit by the shutdown).
- Recoverable downtime losses ≈ 49,900 kWh are added back.
- **Compensated PR ≈ 81.7%** — showing the *equipment* was fine; the loss was an outage, not a defect.
- Availability ≈ 36% (the plant delivered only about a third of its potential that day).

> This day is also the textbook example of the double-count guard: during the outage the grid regulation signal read ~0, which — before the fix — wrongly added a *curtailment* loss on top of the *downtime* loss, inflating Compensated PR to an impossible **108.5%**. Counting curtailment only while the inverter is actually producing (§3.4b) brings it back to the correct 81.7%.

---

## 5. Frequently Asked Questions

**Q1. What is the difference between Raw, SCADA, VCOM and Compensated PR?**
Raw PR is actual-vs-ideal with no allowances. SCADA PR and VCOM PR are the figures reported by the plant's control system and by the independent VCOM platform, respectively (used as cross-checks). Compensated PR adds back the energy lost to outages and grid curtailment, isolating the plant's technical quality. The guarantee is assessed against the compensated view.

**Q2. Why is energy ignored below 50 W/m²?**
At dawn/dusk the sunlight is so weak that sensor noise and tiny production values make PR meaningless (and can divide by almost zero). The 50 W/m² gate removes these unreliable intervals.

**Q3. Why trust the higher pyranometer when the two disagree?**
A soiled or shaded sensor reads **too low**. Averaging it in would understate the sunlight and make PR look artificially high. Using the higher (clean) reading is the conservative, defensible choice.

**Q4. Can Compensated PR be above 100%?**
No — not legitimately. You cannot recover more energy than the sunlight could support. A value over 100% always points to a data problem or a double-counted loss (e.g. counting curtailment during a full outage). Such cases are treated as bugs and corrected.

**Q5. Why "compensate" for outages at all — isn't downtime the plant's problem?**
Compensation separates *technical performance* (which the equipment guarantee is about) from *availability events* (grid curtailment, external outages). Both matter, but they are reported separately: Compensated PR for quality, Availability for uptime.

**Q6. Why does the target go down every year?**
Panels degrade with age, so the contract lowers the guaranteed PR by 0.4% per year. Year 1 (Feb 2025–Jan 2026) uses the full design values; each following February the target steps down one notch.

**Q7. Where do the SCADA and VCOM numbers in the Mother file come from?**
When the monthly `KPI_Report_Daily.xls` (SCADA) and `Performance_ratio_vcom.csv` (VCOM) files are present, their daily PR values are written directly into the "PR SCADA" and "PR VCOM" columns of the monthly summary, replacing the formula estimates.

**Q8. Why are there 96 values per day?**
Data is sampled every 15 minutes; 24 hours × 4 = 96 intervals. Each represents a quarter-hour (0.25 h) of operation.

---

## 6. Glossary

| Term | Definition |
|------|-----------|
| **Performance Ratio (PR)** | Actual energy ÷ ideal energy for the sunlight received, as a percentage. |
| **POA (Plane Of Array)** | Irradiance measured in the panels' own plane — the sunlight they actually receive. |
| **Pyranometer** | Sensor measuring solar irradiance in W/m². |
| **Irradiance** | Instantaneous solar power per area (W/m²). |
| **kWp (kilowatt-peak)** | Panel/plant rated power under standard test sunlight (1000 W/m²). |
| **SCADA** | The plant's on-site supervisory control and data acquisition system. |
| **VCOM** | An independent third-party monitoring platform. |
| **Curtailment** | Grid-ordered reduction of plant output. |
| **Downtime** | An inverter or transformer not producing while the sun is up. |
| **Compensated PR** | Raw PR with recoverable outage/curtailment losses added back. |
| **Σ I (reference irradiance sum)** | Daily total of the gated per-interval reference irradiance (kWh/m²). |
| **PVSyst** | Simulation software providing the monthly design PR targets. |
| **Degradation (0.4%/yr)** | Annual reduction of the guaranteed PR target as panels age. |
| **Availability** | Share of potential energy actually delivered (%). |

---

*This document describes the methodology implemented in `PR_Calculator_GUI_v11.py` for the Mazara 01 plant. It is tailored to that plant's configuration but the principles apply to PV plants generally.*
