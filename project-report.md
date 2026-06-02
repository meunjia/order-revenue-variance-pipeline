# Heavy Electric Equipment — Monthly Variance Analysis Project Report

Date: 2026-06-02  
Company: Iljin Electric Co., Ltd. (Internship)  
Division: Heavy Electric Equipment — Power Transformers & Circuit Breakers

---

## 1. Project Overview

### Purpose

Build an automated pipeline that tracks and explains, at the individual project level, how much
Q1 2026 (January–March) order intake and revenue deviated from the annual business plan — and why.

Previously, the business planning team had to manually compare pivot tables in Excel or
visually cross-reference rows across files, which was time-consuming and prone to omissions.
The core goal was to automate this repetitive work with Python scripts, eliminating manual steps
and improving accuracy.

---

## 2. Input Data

| File | Contents | Rows |
|------|----------|-----:|
| `사업계획_수주_List_123월.csv` | Q1 order plan (Jan–Mar filter from annual plan) | 187 |
| `실적_수주예상.csv` | Order actuals (rows with `수주IN = IN` only) | 336 |
| `사업계획_매출List.csv` | Q1 revenue plan | 468 |
| `실적_매출예상.csv` | Revenue actuals (pre-tax) | 866 |

The original analysis Excel file (`중전기_월별차이분析_20260422.xlsx`) is DRM-protected and
cannot be modified programmatically → converted to CSV for analysis.

---

## 3. Implementation

### Tech Stack

| Library | Purpose |
|---------|---------|
| pandas / numpy | Data loading, aggregation, comparison |
| chardet | Automatic encoding detection (EUC-KR / UTF-8-sig mixed) |
| rapidfuzz | Fuzzy project name matching (≥ 80% similarity) |
| matplotlib | Chart visualization (Korean font support) |
| openpyxl | Formatted Excel report generation |

### Script Structure

```
분析.py                  ← Core: load → summary comparison → project matching → charts + xlsx
차이원인분析.py           ← Structured cause analysis xlsx (extinct / new / changed)
서식생성.py              ← Monthly reporting template xlsx (plan vs actual vs prior year)
보고기준_서식및매칭.py    ← Management report vs CSV reconciliation file
수치흐름.py / 차이원인.py / 양식생성.py  ← Validation and supplementary scripts
```

### Processing Pipeline (분析.py)

```
Step 1. Load & Preprocess
  ├─ Auto-detect encoding per file (chardet)
  ├─ Order plan CSV: skiprows=9  (top 9 rows = exchange rate / composition metadata)
  ├─ Revenue actuals CSV: skiprows=4  (top 4 rows = exchange rate header)
  ├─ Team name normalization:  "01.국내1" → "국내1"
  └─ Product classification:  TR/RT/변압기 → Transformer / HH/HG/GIS/차단기 → Circuit Breaker

Step 2. Summary Comparison
  └─ groupby(['month', 'team', 'product'])
     → plan vs actual outer merge
     → variance / achievement rate calculation

Step 3. Project-Level Deep Comparison
  └─ Matching priority cascade:
       Priority 1: Exact project name match
       Priority 2: PJT number match
       Priority 3: Client (발주처) match
       Priority 4: General contractor (원청) match
       Priority 5: Fuzzy match (token_sort_ratio ≥ 80%)
     → Matched rows:       Amount variance analysis
     → Plan-only rows:     Extinct (cancelled / deferred)
     → Actual-only rows:   New (unplanned orders)

Step 4. Output Generation
  ├─ 7 PNG charts (charts\ folder)
  └─ 3 xlsx reports
```

---

## 4. Results

### 4-1. Orders — Q1 Total

| Category | Plan (₩100M) | Actual (₩100M) | Variance | Achievement |
|----------|------------:|---------------:|---------:|------------:|
| Heavy Electric Total | 2,607.5 | 4,290.1 | **+1,682.6** | **164.5%** |
| └ Power Transformers | 2,231.4 | 3,875.0 | +1,643.6 | 173.7% |
| └ Circuit Breakers | 376.2 | 415.0 | +38.8 | 110.3% |

Actual results significantly exceeded plan.  
The primary driver was a large number of unplanned new orders for overseas renewable energy projects
(wind and solar farms in North America).

### 4-2. Revenue — Q1 Total

| Category | Plan (₩100M) | Actual (₩100M) | Variance | Achievement |
|----------|------------:|---------------:|---------:|------------:|
| Heavy Electric Total | 1,140.5 | 1,214.3 | **+73.8** | **106.5%** |
| └ Power Transformers | 906.7 | 974.2 | +67.5 | 107.4% |
| └ Circuit Breakers | 233.8 | 240.1 | +6.3 | 102.7% |

### 4-3. Project Matching Results

| Category | Matched | Extinct | New |
|----------|--------:|--------:|----:|
| Orders | 38 | 11 | **68** |
| Revenue | 41 | 7 | 52 |

The 68 new order projects (unplanned wins) are the primary reason for the 164.5% achievement rate.

### 4-4. Management Report vs CSV Discrepancy

| Item | Report (₩100M) | CSV Calc (₩100M) | Gap | Root Cause |
|------|---------------:|-----------------:|----:|------------|
| Order Actuals | 4,290.1 | 4,320.9 | +30.8 | Project Apex (Intl): reclassified from Q1 to May |
| Revenue Actuals | 1,214.3 | 1,186.8 | −27.5 | USD exchange rate difference (1,400 vs 1,300) |

Order discrepancy fully traced item by item:

| Team | Project | Variance (₩100M) | Note |
|------|---------|----------------:|------|
| Overseas 2 | Project Apex (Intl) | −30.0 | Confirmation month moved Q1 → May |
| Overseas 1 | Project Nova Phase 1 | −5.8 | Amount revised downward |
| Overseas 1 | Service contract / storage fees (new items) | +2.4 | Project Ridge, Project Gem, etc. |
| Domestic 1 | Small-scale additions | +2.6 | Spare parts, etc. |
| Domestic 2 | — | 0.0 | No change |
| **Total** | | **−30.8** | Reconciles with management report |

---

## 5. Deliverables

### Excel Reports (3 files)

| File | Sheets | Contents |
|------|-------:|---------|
| `차이分析_결과.xlsx` | 7 | Summary comparison (orders/revenue), project deep-dive, new/extinct list, text reasons, visualization dashboard |
| `차이원인分析.xlsx` | 3 | Flow summary, order detail (extinct/new/changed), revenue detail |
| `월별차이分析_1분기.xlsx` | 1 | Monthly reporting template — Q1 filled, Apr–Jul blank |

### Charts — 7 PNG files (`charts\`)

| File | Contents |
|------|---------|
| `수주_월별.png` | Order plan vs actual by month, transformer / CB split |
| `매출_월별.png` | Revenue plan vs actual by month |
| `수주_워터폴.png` | Waterfall: Plan → New → Extinct → Changed → Actual |
| `수주_top10.png` | Top 10 projects by absolute order variance |
| `매출_top10.png` | Top 10 projects by absolute revenue variance |
| `수주_팀달성률.png` | Team-level order achievement donut charts (4 teams) |
| `매출_팀달성률.png` | Team-level revenue achievement donut charts |

---

## 6. Key Technical Issues & Solutions

| Issue | Solution |
|-------|---------|
| Order plan CSV: top 9 rows are metadata (exchange rates, product mix) | `skiprows=list(range(9))` — discovered by inspecting raw bytes |
| Revenue actuals CSV: top 4 rows are exchange rate headers | `skiprows=list(range(4))` applied selectively |
| Mixed encoding: EUC-KR and UTF-8 in different files | chardet auto-detection per file at load time |
| Team name format mismatch (`01.국내1` vs `국내1`) | `re.sub(r'^\d+\.', '')` regex normalization |
| Project name drift between plan and actual systems | rapidfuzz `token_sort_ratio` fuzzy matching at 80% threshold |
| Source Excel DRM-protected — cannot write programmatically | Route through CSV exports; print manual-entry values when write-back fails |

---

## 7. Pending Items

| # | Item | Status |
|---|------|--------|
| 1 | Order actuals +30.8 residual | Root cause identified (Project Apex month reclassification); source filter criteria to be confirmed |
| 2 | Revenue actuals −27.5 residual | Attributed to exchange rate application difference; unconfirmed |
| 3 | Apr–Jul monthly data | Column structure prepared in template; data not yet entered |
| 4 | Auto write-back to source report | DRM prevents direct modification; manual entry required |

---

## 8. Repository Structure

```
월별차이\
├─ README.md                         # English portfolio README
├─ README.ko.md                      # Korean README
├─ project-report.md                 # This document (English)
├─ 월별차이 프로젝트보고서.md          # Korean project report
│
├─ 분析.py                            # Core analysis pipeline
├─ 차이원인分析.py                     # Cause classification Excel
├─ 서식생성.py                        # Reporting template (legacy)
├─ 보고기준_서식및매칭.py              # Management report reconciliation
├─ 수치흐름.py / 차이원인.py / 양식생성.py
│
├─ csv\
│   ├─ sample_사업계획_수주.csv        # Sample: order plan
│   ├─ sample_실적_수주.csv            # Sample: order actuals
│   ├─ sample_사업계획_매출.csv        # Sample: revenue plan
│   └─ sample_실적_매출.csv            # Sample: revenue actuals
│   (Real data files are gitignored)
│
└─ charts\                            # 7 PNG visualization charts
```
