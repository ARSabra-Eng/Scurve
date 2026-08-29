# Parametric S-Curve Workbench v1.0
## User Guide

---

### 1. Introduction

The **Parametric S-Curve Workbench** is a desktop application that converts budget, duration, and shape parameters into time-phased expenditure profiles with live quality assurance (QA). This tool replicates and enhances legacy Excel workbook functionality with automated governance.

**Technology Stack:** Python 3.12 + PySide6 + matplotlib + openpyxl  
**Provenance:** NASA-derived polynomial lineage (via spreadsheet UDF "shg 2012/2013")

---

### 2. Quick Start

#### 2.1 Launch the Application
```bash
python s_curve_workbench.py
```

Or run the packaged executable (Phase 1 deliverable):
```bash
s_curve_workbench.exe
```

#### 2.2 Default Configuration
On startup, the application loads with:
- **Budget (C):** $120,000,000 (illustrative only; any value accepted)
- **Duration (N):** 24 months
- **Preset:** JSC (A=0.32, B=0.68)
- **Start Date:** Current date

---

### 3. User Interface Overview

The main window is divided into three sections:

| Panel | Description |
|-------|-------------|
| **Left Input Panel** | Budget, duration, shape parameters (sliders + direct entry), preset buttons |
| **Center Chart** | Combo chart: bars = Monthly %, line = Cumulative % |
| **Right Panels** | Milestone summary table, QA status panel, data grid tab |

---

### 4. Input Controls

#### 4.1 Budget & Duration
- **Total Budget (C):** Enter any positive value (default: 120,000,000)
- **Duration (N months):** 1–120 months (default: 24)
- **Start Date:** Calendar picker for project start month-year

#### 4.2 Shape Parameters (A, B)
Two methods to set parameters:

**Method 1: Sliders (0–100 scale)**
- Slider position maps to parameter range [−1, +1]
- Formula: `Param = (Scale − 50) / 50`
- Example: Slider at 66 → A = 0.32

**Method 2: Direct Numeric Entry**
- Enter values directly in spin boxes
- Range: −1.0 to +1.0
- Live synchronization with sliders

**Live Display:** A+B sum shown below sliders
- Green text: A+B ≤ 1.0 (admissible)
- Red text: A+B > 1.0 (requires override for export)

#### 4.3 Quick Presets
Five preset buttons replicate legacy workbook parity:

| Preset | Scale A/B | Parameter A | Parameter B | Profile Characteristic |
|--------|-----------|-------------|-------------|------------------------|
| **JSC** | 66 / 84 | 0.32 | 0.68 | Standard NASA-style curve |
| **Rear Loaded** | 50 / 50 | 0.00 | 0.00 | Standard S-curve |
| **Front Loaded** | 100 / 50 | 1.00 | 0.00 | Steep initial spend |
| **Mid Loaded** | 50 / 100 | 0.00 | 1.00 | Delayed expenditure |
| **Balanced** | 75 / 50 | 0.50 | 0.00 | Moderate front-load |

---

### 5. Core Mathematics

The cumulative percentage function F(T) uses the normative polynomial formula:

**F(T) = 10·T²·(1−T)²·(A + B·T) + T⁴·(5 − 4·T)** for 0 < T < 1

Where:
- T = normalized time (period / N)
- A = front-loading parameter
- B = back-loading parameter

**Boundary conditions:**
- F(0) = 0 (no spend before start)
- F(1) = 1 (100% spend by end)

**Monthly rate calculation:**
- p(i) = F(i/N) − F((i−1)/N)
- Monthly(i) = C × p(i)
- Cumulative(i) = Σ Monthly(1..i)

---

### 6. Quality Assurance (QA) Engine

Six automatic checks run on every recalculation:

| Test | Criterion | Failure Consequence |
|------|-----------|---------------------|
| **Admissibility** | A ≥ 0 AND A+B ≤ 1 | Red banner; export requires logged override |
| **Boundary** | F(0) ≈ 0, F(1) ≈ 1 (±1e-9) | Indicates formula error |
| **Range** | 0 ≤ F(T) ≤ 1 on 1,000-point grid | Curve exceeds valid bounds |
| **Monotonicity** | No negative increments | Non-physical spend profile |
| **Period Test** | MIN(Monthly) ≥ 0 | Negative monthly allocation |
| **Reconciliation** | \|Σ Monthly − C\| ≤ 0.01 | Budget mismatch |

**QA Status Display:**
- **PASS (green):** All checks passed; ready for export
- **FAIL (red):** One or more failures; details listed; export requires override reason

---

### 7. Outputs & Exports

#### 7.1 Excel Export (.xlsx)
Click **Export → Excel** to generate workbook with sheets:

| Sheet | Contents |
|-------|----------|
| **Inputs** | Budget, duration, A, B, scales, A+B sum |
| **PackageRegister** | Single package placeholder (Phase 1) |
| **Overlays** | Milestone overlay placeholder (Phase 2) |
| **Monthly** | Period, monthly amount, monthly % |
| **Cumulative** | Period, cumulative amount, cumulative % |
| **Milestones** | F(T) at T = 0.10, 0.25, 0.50, 0.65, 0.75, 0.90 |
| **QA_Log** | Six QA test results with pass/fail status |
| **AnnexB_Record** | Coefficient selection audit record |

**Note:** E2/E4 scale values round-trip exactly via §4 conversion formula.

#### 7.2 PDF Report (Phase 2)
One-page report including:
- S-curve chart
- Milestone summary table
- QA status panel
- Parameter/version stamp

#### 7.3 Project Save/Load (.json) (Phase 2)
Save and reload complete project state including parameters and audit trail.

---

### 8. Audit Trail

All parameter changes are automatically logged with:
- Timestamp
- Event type (parameter change, coefficient selection, export)
- Old/new values
- Rationale (auto-populated or user-entered)
- Approver (default: "System")
- Version stamp

**Access:** View audit log in AnnexB_Record sheet after export.

---

### 9. Acceptance Tests (TC1–TC5)

Run automated tests from command line:
```bash
python s_curve_workbench.py --test
```

| Test | Description | Expected Result |
|------|-------------|-----------------|
| **TC1** | JSC preset (A=0.32, B=0.68) | P12=60.0000%, P18=92.4609%, P24=100% |
| **TC1b** | Milestone % table | 3.19/18.79/60.00/82.28/92.46/99.40 (±0.01) |
| **TC2** | Negative A (A=-0.06) | QA FAIL; P1 monthly ≈ -$58,840 |
| **TC3** | Boundary (A=1.0, B=0.0) | QA PASS; monotonic |
| **TC4** | All presets | Σ Monthly = C within 0.01 |
| **TC5** | Excel export | Opens in Excel; values match grid to 2 dp |

---

### 10. Troubleshooting

| Issue | Resolution |
|-------|------------|
| Red QA banner appears | Check A ≥ 0 and A+B ≤ 1; adjust sliders or enter valid values |
| Export fails | Ensure QA passes or provide override reason in audit log |
| Chart not updating | Click Recalculate button or modify any input to trigger refresh |
| Negative monthly values | Reduce A if negative; check admissibility status |

---

### 11. Technical Support

**Log File:** `s_curve_workbench.log` (located in application directory)  
**Audit Trail:** `audit_trail.json` (auto-saved with each session)

For issues, provide:
1. Log file contents
2. Screenshot of QA panel
3. Steps to reproduce

---

### 12. Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-08-29 | Phase 1 release: single-package S-curve generation, QA engine, Excel export |

---

**Document Control:** This user guide corresponds to Parametric S-Curve Workbench v1.0, implementing requirements R-01 through R-05 per Developer Brief v1.1.
