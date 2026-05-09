# Day 34 — Data Type Conversion & Missing Values Basics

**Roadmap Phase:** Phase 2 — Advanced Excel + pandas Intro
**Date:** May 2026
**Stack:** Python · pandas · openpyxl

---

## What This Covers

- Detecting column data types with `dtypes`
- Safe numeric conversion using `pd.to_numeric(errors='coerce')`
- Multi-format date parsing using `pd.to_datetime(format='mixed')`
- Force-converting clean columns with `astype()`
- Detecting nulls with `isnull()` / `notnull()`
- Filling missing values with `fillna()` — fixed value, median, forward-fill, calculated
- Removing incomplete rows with `dropna()` — by subset, threshold, strategy
- Stripping string-encoded values (e.g. `"9.82%"`) into usable floats

---

## Files

| File | Description |
|---|---|
| `day34_input.xlsx` | Raw messy transaction + portfolio data (intentional errors) |
| `day34_practice.py` | Full annotated practice script with 10 sections + 5 tasks |
| `day34_output.xlsx` | Cleaned output — two sheets exported by pandas |

---

## Key Functions Reference

```python
df.dtypes                                   # Check all column types
pd.to_numeric(df["col"], errors="coerce")   # Safe → NaN on bad values
pd.to_datetime(df["col"], format="mixed")   # Parse any date format
df["col"].astype(str)                       # Force convert type
df.isnull().sum()                           # Count missing per column
df.fillna(value)                            # Fill with fixed value
df.fillna(method="ffill")                   # Forward fill
df.dropna(subset=["col"])                   # Drop rows missing in col
```

---

## Portfolio Connection

Handles the most common real-world data quality problem: dirty input files.
Feeds directly into **Day 36 (Missing Data Handling)** and the
**Phase 7 Financial Dashboard** where live API data will need identical cleaning pipelines.

---

*84-Day Python + Excel Roadmap | Bhanu Pratap Singh*
