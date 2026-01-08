import pandas as pd

# ============================================================
# FILE PATHS
# ============================================================
BEFORE_FILE = "before.xlsx"
AFTER_FILE = "after.xlsx"

OUT_ONLY_BEFORE = "present_in_before_not_in_after.xlsx"
OUT_ONLY_AFTER = "present_in_after_not_in_before.xlsx"
OUT_COMMON = "present_in_both.xlsx"

# ============================================================
# LOAD FILES
# ============================================================
before = pd.read_excel(BEFORE_FILE)
after = pd.read_excel(AFTER_FILE)

# ============================================================
# REQUIRED COLUMNS CHECK
# ============================================================
KEYS = ["task_id", "rsrc_id"]

for col in KEYS:
    if col not in before.columns:
        raise ValueError(f"Missing column '{col}' in BEFORE file")
    if col not in after.columns:
        raise ValueError(f"Missing column '{col}' in AFTER file")

# ============================================================
# CLEAN KEY COLUMNS
# ============================================================
for col in KEYS:
    before[col] = before[col].astype(str).str.strip()
    after[col] = after[col].astype(str).str.strip()

# ============================================================
# 1. PRESENT IN BEFORE BUT NOT IN AFTER
# ============================================================
only_before = before.merge(
    after[KEYS],
    on=KEYS,
    how="left",
    indicator=True
).query("_merge == 'left_only'").drop(columns="_merge")

only_before.to_excel(OUT_ONLY_BEFORE, index=False)

# ============================================================
# 2. PRESENT IN AFTER BUT NOT IN BEFORE
# ============================================================
only_after = after.merge(
    before[KEYS],
    on=KEYS,
    how="left",
    indicator=True
).query("_merge == 'left_only'").drop(columns="_merge")

only_after.to_excel(OUT_ONLY_AFTER, index=False)

# ============================================================
# 3. PRESENT IN BOTH FILES
# ============================================================
common = before.merge(
    after,
    on=KEYS,
    how="inner",
    suffixes=("_before", "_after")
)

common.to_excel(OUT_COMMON, index=False)

# ============================================================
# SUMMARY
# ============================================================
print("✔ COMPARISON COMPLETE")
print(f"✔ Only in BEFORE : {len(only_before)} rows")
print(f"✔ Only in AFTER  : {len(only_after)} rows")
print(f"✔ Present in BOTH: {len(common)} rows")
