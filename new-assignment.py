import pandas as pd

# ============================================================
# FILE PATHS
# ============================================================
ACTIVITY_FILE = "temp-2.xlsx"
BOQ_FILE = "boq-resource-id-hrs.xlsx"
OUTPUT_FILE = "output-final-3.xlsx"

# ============================================================
# LOAD FILES
# ============================================================
act = pd.read_excel(ACTIVITY_FILE)
boq = pd.read_excel(BOQ_FILE)

# ============================================================
# NORMALIZE BOQ COLUMNS (CRITICAL)
# ============================================================
boq = boq.rename(columns={
    "Resource ID": "boq_rsrc_id",
    "BOQ Item No.": "BOQ Item No."
})

required_boq_cols = {"boq_rsrc_id", "BOQ Item No.", "Qty/Unit (Norms)"}
missing = required_boq_cols - set(boq.columns)
if missing:
    raise ValueError(f"BOQ file missing required columns: {missing}")

boq["boq_rsrc_id"] = boq["boq_rsrc_id"].astype(str).str.strip()
boq["BOQ Item No."] = boq["BOQ Item No."].astype(str).str.strip()
boq["Qty/Unit (Norms)"] = pd.to_numeric(boq["Qty/Unit (Norms)"], errors="coerce")

# ============================================================
# CLEAN ACTIVITY FILE
# ============================================================
act["task_id"] = act["task_id"].astype(str).str.strip()
act["user_field_130"] = act["user_field_130"].astype(str).str.strip()

for c in ["target_qty", "act_qty", "remain_qty"]:
    act[c] = pd.to_numeric(act[c], errors="coerce")

# ============================================================
# STEP 1: ASSIGN NONLABOR RESOURCES FROM BOQ
# ============================================================
generated = act.merge(
    boq,
    left_on="user_field_130",
    right_on="BOQ Item No.",
    how="inner"
)

new_nonlabor = pd.DataFrame({
    "rsrc_id": generated["boq_rsrc_id"],
    "task_id": generated["task_id"],
    "TASK__status_code": "",
    "role_id": "",
    "acct_id": "",
    "rsrc_type": "Nonlabor",
    "rsrc__rsrc_name": "",
    "target_cost": 0,
    "act_cost": 0,
    "remain_cost": 0,
    "target_qty": 0,
    "act_qty": 0,
    "remain_qty": 0,
    "user_field_130": generated["user_field_130"],
    "user_field_131": "",
    "delete_record_flag": ""
})

# Align columns and append
for col in act.columns:
    if col not in new_nonlabor.columns:
        new_nonlabor[col] = None

new_nonlabor = new_nonlabor[act.columns]
act = pd.concat([act, new_nonlabor], ignore_index=True)

# ============================================================
# STEP 2: CALCULATE NONLABOR QUANTITIES FROM MATERIAL
# ============================================================
# Rename for processing
act = act.rename(columns={
    "task_id": "Activity ID",
    "rsrc_id": "Resource ID",
    "rsrc_type": "Resource Type",
    "target_qty": "Budgeted Units",
    "act_qty": "Actual Units",
    "remain_qty": "Remaining Units",
    "user_field_130": "BOQ Item No."
})

for c in ["Activity ID", "Resource ID", "Resource Type", "BOQ Item No."]:
    act[c] = act[c].astype(str).str.strip()

material = act[act["Resource Type"].str.lower() == "material"]

# Identify nonlabor rows and PRESERVE ORIGINAL INDEX
nonlabor_mask = act["Resource Type"].str.lower() == "nonlabor"
equipment = act.loc[nonlabor_mask].copy()
equipment["_orig_index"] = equipment.index

# Aggregate material quantities per Activity + BOQ
material_qty = (
    material
    .groupby(["Activity ID", "BOQ Item No."], as_index=False)
    .agg({
        "Budgeted Units": "sum",
        "Actual Units": "sum",
        "Remaining Units": "sum"
    })
)

# Merge material quantities
equipment = equipment.merge(
    material_qty,
    on=["Activity ID", "BOQ Item No."],
    how="left",
    suffixes=("", "_MAT")
)

# Merge BOQ norms
equipment = equipment.merge(
    boq[["BOQ Item No.", "boq_rsrc_id", "Qty/Unit (Norms)"]],
    left_on=["BOQ Item No.", "Resource ID"],
    right_on=["BOQ Item No.", "boq_rsrc_id"],
    how="left"
)

# Calculate quantities
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
# WRITE BACK BY ORIGINAL INDEX (CRITICAL FIX)
# ============================================================
act.loc[equipment["_orig_index"], "Budgeted Units"] = equipment["Budgeted Units"].values
act.loc[equipment["_orig_index"], "Actual Units"] = equipment["Actual Units"].values
act.loc[equipment["_orig_index"], "Remaining Units"] = equipment["Remaining Units"].values

# ============================================================
# RESTORE ORIGINAL COLUMN NAMES & EXPORT
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

act.to_excel(OUTPUT_FILE, index=False)

print("✔ SUCCESS")
print("✔ Nonlabor resources assigned from BOQ")
print("✔ Quantities calculated from material")
print("✔ Output written to:", OUTPUT_FILE)
