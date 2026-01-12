import pandas as pd

# ============================================================
# FILE PATHS
# ============================================================
ACTIVITY_FILE = "primavera-output.xlsx"
BOQ_FILE = "boq-resource-id-hrs.xlsx"
OUTPUT_FILE = "resource-assignment-code-output.xlsx"

# ============================================================
# LOAD FILES
# ============================================================
act = pd.read_excel(ACTIVITY_FILE)
boq = pd.read_excel(BOQ_FILE)

# ============================================================
# NORMALIZE BOQ FILE (SOURCE OF TRUTH)
# ============================================================
boq = boq.rename(columns={
    "Resource ID": "boq_rsrc_id",
    "BOQ Item No.": "BOQ Item No.",
    "Resource Type": "boq_rsrc_type"
})

required_boq_cols = {
    "boq_rsrc_id",
    "BOQ Item No.",
    "Qty/Unit (Norms)",
    "boq_rsrc_type"
}
missing = required_boq_cols - set(boq.columns)
if missing:
    raise ValueError(f"BOQ file missing required columns: {missing}")

boq["boq_rsrc_id"] = boq["boq_rsrc_id"].astype(str).str.strip()
boq["BOQ Item No."] = boq["BOQ Item No."].astype(str).str.strip()
boq["boq_rsrc_type"] = boq["boq_rsrc_type"].astype(str).str.strip().str.title()
boq["Qty/Unit (Norms)"] = pd.to_numeric(boq["Qty/Unit (Norms)"], errors="coerce").fillna(0)

# Validate BOQ resource types
valid_types = {"Labor", "Nonlabor"}
invalid = boq.loc[~boq["boq_rsrc_type"].isin(valid_types)]
if not invalid.empty:
    raise ValueError(
        "Invalid resource types in BOQ:\n"
        + invalid[["boq_rsrc_id", "boq_rsrc_type"]].to_string(index=False)
    )

# ============================================================
# CLEAN ACTIVITY FILE
# ============================================================
act["task_id"] = act["task_id"].astype(str).str.strip()
act["rsrc_id"] = act["rsrc_id"].astype(str).str.strip()
act["rsrc_type"] = act["rsrc_type"].astype(str).str.strip()
act["user_field_130"] = act["user_field_130"].astype(str).str.strip()

for c in ["target_qty", "act_qty", "remain_qty"]:
    act[c] = pd.to_numeric(act[c], errors="coerce").fillna(0)

# ============================================================
# LOOKUP: ACTIVITY STATUS
# ============================================================
activity_status = (
    act.groupby("task_id")["TASK__status_code"]
    .first()
    .to_dict()
)

# ============================================================
# STEP 1: GENERATE LABOR + NONLABOR RESOURCES FROM BOQ
# ============================================================
generated = act.merge(
    boq,
    left_on="user_field_130",
    right_on="BOQ Item No.",
    how="inner"
)

new_resources = pd.DataFrame({
    "rsrc_id": generated["boq_rsrc_id"],
    "task_id": generated["task_id"],
    "TASK__status_code": generated["task_id"].map(activity_status),
    "role_id": "",
    "acct_id": "",
    "rsrc_type": generated["boq_rsrc_type"],   # ✅ FROM BOQ
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
    if col not in new_resources.columns:
        new_resources[col] = None

new_resources = new_resources[act.columns]
act = pd.concat([act, new_resources], ignore_index=True)

# ============================================================
# STEP 2: CALCULATE LABOR + NONLABOR FROM MATERIAL
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

for c in ["Activity ID", "Resource ID", "Resource Type", "BOQ Item No."]:
    act[c] = act[c].astype(str).str.strip()

# Aggregate material quantities
material = act[act["Resource Type"].str.lower() == "material"]

material_qty = (
    material
    .groupby(["Activity ID", "BOQ Item No."], as_index=False)
    .agg({
        "Budgeted Units": "sum",
        "Actual Units": "sum",
        "Remaining Units": "sum"
    })
)

# Labor + Nonlabor rows
calc_mask = act["Resource Type"].str.lower().isin(["labor", "nonlabor"])
calc_df = act.loc[calc_mask].copy()
calc_df["_orig_index"] = calc_df.index

# Merge material totals
calc_df = calc_df.merge(
    material_qty,
    on=["Activity ID", "BOQ Item No."],
    how="left",
    suffixes=("", "_MAT")
)

# Merge BOQ norms
calc_df = calc_df.merge(
    boq[["BOQ Item No.", "boq_rsrc_id", "Qty/Unit (Norms)"]],
    left_on=["BOQ Item No.", "Resource ID"],
    right_on=["BOQ Item No.", "boq_rsrc_id"],
    how="left"
)

calc_df[[
    "Budgeted Units_MAT",
    "Actual Units_MAT",
    "Remaining Units_MAT"
]] = calc_df[[
    "Budgeted Units_MAT",
    "Actual Units_MAT",
    "Remaining Units_MAT"
]].fillna(0)

calc_df["Qty/Unit (Norms)"] = calc_df["Qty/Unit (Norms)"].fillna(0)

# Calculate quantities
calc_df["Budgeted Units"] = calc_df["Budgeted Units_MAT"] * calc_df["Qty/Unit (Norms)"]
calc_df["Actual Units"] = calc_df["Actual Units_MAT"] * calc_df["Qty/Unit (Norms)"]
calc_df["Remaining Units"] = calc_df["Remaining Units_MAT"] * calc_df["Qty/Unit (Norms)"]

# ============================================================
# WRITE BACK USING ORIGINAL INDEX
# ============================================================
act.loc[calc_df["_orig_index"], "Budgeted Units"] = calc_df["Budgeted Units"].values
act.loc[calc_df["_orig_index"], "Actual Units"] = calc_df["Actual Units"].values
act.loc[calc_df["_orig_index"], "Remaining Units"] = calc_df["Remaining Units"].values

# ============================================================
# RESTORE COLUMN NAMES & EXPORT
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
print("✔ Resource type taken strictly from BOQ")
print("✔ Labor & Nonlabor quantities derived from material")
print("✔ TASK status inherited correctly")
print("✔ Output written to:", OUTPUT_FILE)
