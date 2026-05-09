# ============================================================
# DAY 34 — Data Type Conversion & Missing Values Basics
# Bhanu Pratap Singh | 84-Day Python + Excel Roadmap
# ============================================================

import pandas as pd
import numpy as np

# ──────────────────────────────────────────────
# SECTION 1: Load the data
# ──────────────────────────────────────────────

df = pd.read_excel("day34_input.xlsx", sheet_name="Transactions")
pf = pd.read_excel("day34_input.xlsx", sheet_name="Portfolio")

print("=" * 55)
print("SECTION 1 — Raw Data Overview")
print("=" * 55)
print(df.head())
print()

# ──────────────────────────────────────────────
# SECTION 2: Check Data Types (dtypes)
# ──────────────────────────────────────────────
# dtypes tells you what pandas THINKS each column is.
# 'object' means string/mixed. int64/float64 are numeric.

print("=" * 55)
print("SECTION 2 — Data Types (as loaded)")
print("=" * 55)
print(df.dtypes)
print()

# Notice: Quantity and Price are 'object' instead of numeric
# because rows with 'abc' or '' confused pandas.

# ──────────────────────────────────────────────
# SECTION 3: pd.to_numeric() — Safe Number Conversion
# ──────────────────────────────────────────────
# errors='coerce' → converts bad values to NaN (not crash)
# errors='raise'  → crashes on bad values (default)
# errors='ignore' → leaves bad values unchanged (risky)

print("=" * 55)
print("SECTION 3 — pd.to_numeric() with errors='coerce'")
print("=" * 55)

df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
df["Price"]    = pd.to_numeric(df["Price"],    errors="coerce")
df["Amount"]   = pd.to_numeric(df["Amount"],   errors="coerce")

print("After conversion:")
print(df[["TxnID", "Quantity", "Price", "Amount"]].to_string())
# Rows with 'abc' or 'N/A' now show NaN — safe for calculations.
print()

# ──────────────────────────────────────────────
# SECTION 4: pd.to_datetime() — Date Standardization
# ──────────────────────────────────────────────
# format='mixed' handles multiple date formats at once.
# dayfirst=True helps for formats like 22-01-2024.

print("=" * 55)
print("SECTION 4 — pd.to_datetime() — multiple date formats")
print("=" * 55)

print("Raw Date column (mixed formats):")
print(df["Date"].tolist())
print()

df["Date"] = pd.to_datetime(df["Date"], format="mixed", dayfirst=True, errors="coerce")
print("After standardization:")
print(df["Date"].tolist())
print()

# ──────────────────────────────────────────────
# SECTION 5: astype() — Force-Convert a Column
# ──────────────────────────────────────────────
# Use astype() when you are SURE there are no bad values.
# astype(str) → convert to string
# astype(int) → convert to integer (crashes on NaN!)
# astype(float) → convert to float

print("=" * 55)
print("SECTION 5 — astype() for clean columns")
print("=" * 55)

# TxnID is fine as string — make it explicit
df["TxnID"] = df["TxnID"].astype(str)

# Broker_Fee: has some NaNs — fill first, THEN convert
df["Broker_Fee"] = pd.to_numeric(df["Broker_Fee"], errors="coerce")  # safe convert
print("TxnID dtype:", df["TxnID"].dtype)
print("Broker_Fee dtype:", df["Broker_Fee"].dtype)
print()

# ──────────────────────────────────────────────
# SECTION 6: Detect Missing Values
# ──────────────────────────────────────────────
# isnull() / isna()   → True where value is NaN/None
# notnull() / notna() → True where value EXISTS

print("=" * 55)
print("SECTION 6 — Detecting Missing Values")
print("=" * 55)

# Total missing per column
print("Missing values per column:")
print(df.isnull().sum())
print()

# Total missing in entire DataFrame
total_missing = df.isnull().sum().sum()
print(f"Total missing cells: {total_missing}")
print()

# Show ONLY rows that have at least one missing value
print("Rows with missing data:")
print(df[df.isnull().any(axis=1)][["TxnID", "Quantity", "Price", "Amount", "Broker_Fee", "Notes"]])
print()

# ──────────────────────────────────────────────
# SECTION 7: fillna() — Fill Missing Values
# ──────────────────────────────────────────────
# fillna(value)         → fill with a fixed value
# fillna(method='ffill') → forward fill (use previous row's value)
# fillna(method='bfill') → backward fill (use next row's value)
# fillna(df.mean())     → fill with column average

print("=" * 55)
print("SECTION 7 — fillna() — Different Strategies")
print("=" * 55)

df_filled = df.copy()  # work on a copy to preserve original

# Strategy 1: Fill Broker_Fee missing with column median
median_fee = df_filled["Broker_Fee"].median()
df_filled["Broker_Fee"] = df_filled["Broker_Fee"].fillna(median_fee)
print(f"Broker_Fee filled with median: {median_fee:.2f}")

# Strategy 2: Fill Notes with a placeholder string
df_filled["Notes"] = df_filled["Notes"].fillna("No notes")

# Strategy 3: Forward fill Quantity (carry forward previous row)
# Note: In newer pandas (2.0+), use .ffill() directly instead of fillna(method='ffill')
df_filled["Quantity"] = df_filled["Quantity"].ffill()

# Strategy 4: Fill Amount with calculated value where possible
mask = df_filled["Amount"].isnull() & df_filled["Quantity"].notnull() & df_filled["Price"].notnull()
df_filled.loc[mask, "Amount"] = df_filled.loc[mask, "Quantity"] * df_filled.loc[mask, "Price"]

print("\nAfter fillna strategies:")
print(df_filled[["TxnID", "Quantity", "Price", "Amount", "Broker_Fee", "Notes"]].to_string())
print()

# ──────────────────────────────────────────────
# SECTION 8: dropna() — Remove Rows with Missing Data
# ──────────────────────────────────────────────
# dropna()                   → drop any row with at least 1 NaN
# dropna(subset=['col'])     → drop only if specific column is NaN
# dropna(thresh=N)           → keep rows that have at least N non-NaN values
# dropna(how='all')          → drop only if ALL columns are NaN

print("=" * 55)
print("SECTION 8 — dropna() — Selective Row Removal")
print("=" * 55)

print(f"Original rows: {len(df)}")

# Drop rows where Price is missing (can't do any calc without price)
df_clean = df.dropna(subset=["Price"])
print(f"After dropping rows with missing Price: {len(df_clean)}")

# Drop rows where BOTH Quantity AND Amount are missing
df_clean2 = df.dropna(subset=["Quantity", "Amount"], how="all")
print(f"After dropping rows where both Qty & Amount missing: {len(df_clean2)}")
print()

# ──────────────────────────────────────────────
# SECTION 9: Portfolio Sheet — Return_Pct Cleanup
# ──────────────────────────────────────────────
# Return_Pct loaded as string ('9.82%') — strip % and convert

print("=" * 55)
print("SECTION 9 — Portfolio: Clean Return_Pct column")
print("=" * 55)

print("Return_Pct as loaded:", pf["Return_Pct"].tolist())
print("dtype:", pf["Return_Pct"].dtype)
print()

# Step 1: Replace 'N/A' string with actual NaN
pf["Return_Pct"] = pf["Return_Pct"].replace("N/A", np.nan)

# Step 2: Strip the '%' sign and convert to float
pf["Return_Pct"] = pf["Return_Pct"].str.replace("%", "", regex=False)
pf["Return_Pct"] = pd.to_numeric(pf["Return_Pct"], errors="coerce")

# Step 3: Convert Current_Price similarly
pf["Current_Price"] = pf["Current_Price"].replace("N/A", np.nan)
pf["Current_Price"] = pd.to_numeric(pf["Current_Price"], errors="coerce")

# Step 4: Rating — convert to numeric (some are None)
pf["Rating"] = pd.to_numeric(pf["Rating"], errors="coerce")

print("After cleanup:")
print(pf[["Stock", "Return_Pct", "Current_Price", "Rating"]])
print()

# ──────────────────────────────────────────────
# SECTION 10: Export Clean Data to Excel
# ──────────────────────────────────────────────

print("=" * 55)
print("SECTION 10 — Export cleaned data to Excel")
print("=" * 55)

with pd.ExcelWriter("day34_output.xlsx", engine="openpyxl") as writer:
    df_filled.to_excel(writer, sheet_name="Transactions_Clean", index=False)
    pf.to_excel(writer, sheet_name="Portfolio_Clean", index=False)

print("Exported: day34_output.xlsx")
print("  Sheet 1 → Transactions_Clean")
print("  Sheet 2 → Portfolio_Clean")

# ──────────────────────────────────────────────
# DAILY TASKS
# ──────────────────────────────────────────────
# TASK 1: In df_filled, calculate a new column 'Net_Amount'
#         = Amount + Broker_Fee. Print the result.
#
# TASK 2: In df_filled, filter and print only rows where
#         Status == 'Completed' AND Amount > 100000.
#
# TASK 3: On the Portfolio sheet, fill missing Current_Price
#         with the column median. Then calculate a new column
#         'Current_Value' = Shares * Current_Price.
#         (Hint: Shares column is currently a string — fix it first!)
#
# TASK 4: Count how many transactions happened per month.
#         (Hint: use df['Date'].dt.month and groupby/value_counts)
#
# TASK 5 [CHALLENGE]: Build a summary dict with these stats:
#         - Total completed transactions
#         - Total amount (completed only)
#         - Average broker fee (after filling NaN)
#         - Number of rows that had any missing data originally
#         Print it cleanly with labels.
