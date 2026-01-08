import pandas as pd

old = pd.read_excel("before.xlsx")
new = pd.read_excel("after.xlsx")

KEYS = ["task_id", "rsrc_id"]
COMPARE_COLS = ["target_qty", "act_qty", "remain_qty"]

# Merge old vs new
diff = old.merge(
    new,
    on=KEYS,
    how="outer",
    suffixes=("_old", "_new"),
    indicator=True
)

# Find changed quantities
changes = diff[
    (diff["_merge"] == "both") &
    (
        (diff["target_qty_old"] != diff["target_qty_new"]) |
        (diff["act_qty_old"] != diff["act_qty_new"]) |
        (diff["remain_qty_old"] != diff["remain_qty_new"])
    )
]

changes.to_excel("changes-only.xlsx", index=False)
