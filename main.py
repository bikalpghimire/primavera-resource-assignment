import pandas as pd

# ============================================================
# FILE PATHS
# ============================================================
ACTIVITY_FILE = "OCB-1_06 Jan- 2026-Activity Resource Assignments.xlsx"
BOQ_FILE = "boq-resource-id-hrs.xlsx"
OUTPUT_FILE = "activity_equipment_units_material_driven.xlsx"

# ============================================================
# LOAD FILES
# ============================================================
act = pd.read_excel(ACTIVITY_FILE)
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
# CLEAN & NORMALIZE TEXT FIELDS
# ============================================================
for c in ["Activity ID", "Resource ID", "Resource Type", "BOQ Item No."]:
    act[c] = act[c].astype(str).str.strip()

boq["BOQ Item No."] = boq["BOQ Item No."].astype(str).str.strip()
boq["Resource ID"] = boq["Resource ID"].astype(str).str.strip()

# ============================================================
# FORCE NUMERIC (CRITICAL FIX)
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
equipment = act[act["Resource Type"].str.lower() == "nonlabor"].copy()

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
# ATTACH MATERIAL QUANTITIES TO EQUIPMENT
# ============================================================
equipment = equipment.merge(
    material_qty,
    on=["Activity ID", "BOQ Item No."],
    how="left",
    suffixes=("", "_MAT")
)

# Ensure numeric after merge
for c in ["Budgeted Units_MAT", "Actual Units_MAT", "Remaining Units_MAT"]:
    equipment[c] = pd.to_numeric(equipment[c], errors="coerce")

# ============================================================
# ATTACH BOQ NORMS (BOQ + Resource ID)
# ============================================================
equipment = equipment.merge(
    boq,
    on=["BOQ Item No.", "Resource ID"],
    how="left"
)

# ============================================================
# CALCULATE EQUIPMENT UNITS
# ============================================================
equipment["Calc Budgeted Units"] = (
    equipment["Budgeted Units_MAT"] * equipment["Qty/Unit (Norms)"]
)

equipment["Calc Actual Units"] = (
    equipment["Actual Units_MAT"] * equipment["Qty/Unit (Norms)"]
)

equipment["Calc Remaining Units"] = (
    equipment["Remaining Units_MAT"] * equipment["Qty/Unit (Norms)"]
)

equipment["CALC_OK"] = (
    equipment["Qty/Unit (Norms)"].notna() &
    equipment["Budgeted Units_MAT"].notna()
)

# ============================================================
# COMBINE BACK & EXPORT
# ============================================================
final = pd.concat([material, equipment], ignore_index=True)

final.to_excel(OUTPUT_FILE, index=False)

print("✔ SUCCESS")
print("✔ Output written to:", OUTPUT_FILE)
