import pandas as pd

# ============================================================
# FILE PATH
# ============================================================
FILE_PATH = "OCB-1_06 Jan- 2026-Activity Resource Assignments.xlsx"
BOQ_FILE = "boq-resource-id-hrs.xlsx"

# ============================================================
# LOAD FILES
# ============================================================
act = pd.read_excel(FILE_PATH)
boq = pd.read_excel(BOQ_FILE)

# ============================================================
# RENAME COLUMNS (Primavera → Logical)
# ============================================================
act = act.rename(columns={
    "task_id": "Activity ID",
    "rsrc_id": "Resource ID",
    "rsrc_type": "Resource Type",
    "target_qty": "Budgeted Units",
    "act_qty": "Actual Units",
    "remain_qty": "Remaining Units",
    "user_field_130": "BOQ Item No."
})

# ============================================================
# CLEAN TEXT
# ============================================================
for c in ["Activity ID", "Resource ID", "Resource Type", "BOQ Item No."]:
    act[c] = act[c].astype(str).str.strip()

boq["BOQ Item No."] = boq["BOQ Item No."].astype(str).str.strip()
boq["Resource ID"] = boq["Resource ID"].astype(str).str.strip()

# ============================================================
# FORCE NUMERIC (CRITICAL)
# ============================================================
qty_cols = ["Budgeted Units", "Actual Units", "Remaining Units"]
for c in qty_cols:
    act[c] = pd.to_numeric(act[c], errors="coerce")

boq["Qty/Unit (Norms)"] = pd.to_numeric(
    boq["Qty/Unit (Norms)"], errors="coerce"
)

# ============================================================
# SPLIT MATERIAL & NONLABOR
# ============================================================
material = act[act["Resource Type"].str.lower() == "material"].copy()
equipment_idx = act["Resource Type"].str.lower() == "nonlabor"

# ============================================================
# AGGREGATE MATERIAL QUANTITIES
# (Activity ID + BOQ Item No.)
# ============================================================
material_qty = (
    material
    .groupby(["Activity ID", "BOQ Item No."], as_index=False)
    .agg({
        "Budgeted Units": "sum",
        "Actual Units": "sum",
        "Remaining Units": "sum"
    })
)

# ============================================================
# PREPARE EQUIPMENT DATAFRAME
# ============================================================
equipment = act.loc[equipment_idx].copy()

equipment = equipment.merge(
    material_qty,
    on=["Activity ID", "BOQ Item No."],
    how="left",
    suffixes=("", "_MAT")
)

equipment = equipment.merge(
    boq,
    on=["BOQ Item No.", "Resource ID"],
    how="left"
)

# ============================================================
# CALCULATE EQUIPMENT UNITS
# ============================================================
equipment["Budgeted Units"] = (
    equipment["Budgeted Units_MAT"] * equipment["Qty/Unit (Norms)"]
)

equipment["Actual Units"] = (
    equipment["Actual Units_MAT"] * equipment["Qty/Unit (Norms)"]
)

equipment["Remaining Units"] = (
    equipment["Remaining Units_MAT"] * equipment["Qty/Unit (Norms)"]
)

# ============================================================
# WRITE BACK INTO ORIGINAL DATAFRAME
# ============================================================
act.loc[equipment_idx, "Budgeted Units"] = equipment["Budgeted Units"].values
act.loc[equipment_idx, "Actual Units"] = equipment["Actual Units"].values
act.loc[equipment_idx, "Remaining Units"] = equipment["Remaining Units"].values

# ============================================================
# RESTORE ORIGINAL COLUMN NAMES (OPTIONAL, SAFE)
# ============================================================
act = act.rename(columns={
    "Activity ID": "task_id",
    "Resource ID": "rsrc_id",
    "Resource Type": "rsrc_type",
    "Budgeted Units": "target_qty",
    "Actual Units": "act_qty",
    "Remaining Units": "remain_qty",
    "BOQ Item No.": "user_field_130"
})

# ============================================================
# OVERWRITE INPUT FILE
# ============================================================
act.to_excel(FILE_PATH, index=False)

print("✔ SUCCESS")
print("✔ Nonlabor quantities overwritten with calculated values")
print("✔ Material and Labor rows untouched")
