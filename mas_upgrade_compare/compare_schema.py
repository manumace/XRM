"""Compare Maximo (MAS) column schema: Prod vs SysTest.

Scope: only tables that exist in Prod.
  - columns added / dropped on tables present in both
  - data type, length, precision, scale, nullability changes
  - Prod tables missing from SysTest
New SysTest-only tables are ignored.

Usage: python compare_schema.py <MAS_Upgrade.xlsx> [out_dir]
"""
import sys
from pathlib import Path

import pandas as pd

src = sys.argv[1]
out = Path(sys.argv[2] if len(sys.argv) > 2 else ".")
out.mkdir(parents=True, exist_ok=True)

KEY = ["schema_name", "table_name", "column_name"]
ATTRS = ["data_type", "length", "max_length_bytes", "precision", "scale", "is_nullable"]


def load(sheet):
    df = pd.read_excel(src, sheet_name=sheet)
    for c in KEY:
        df[c] = df[c].astype(str).str.strip().str.lower()  # SQL Server is case-insensitive
    df["base_type"] = df["data_type"].str.extract(r"^(\w+)", expand=False)
    df["length"] = df["length"].fillna(-1)
    return df


prod, sys_ = load("Prod"), load("SysTest")
tbl = ["schema_name", "table_name"]
prod_tables = prod[tbl].drop_duplicates()
sys_tables = sys_[tbl].drop_duplicates()
common = prod_tables.merge(sys_tables, on=tbl)

# --- tables in Prod missing from SysTest
dropped_tables = (
    prod.groupby(tbl).agg(prod_column_count=("column_name", "size")).reset_index()
    .merge(sys_tables, on=tbl, how="left", indicator=True)
    .query("_merge == 'left_only'").drop(columns="_merge")
)

# --- column-level diff on common tables
p = prod.merge(common, on=tbl)
s = sys_.merge(common, on=tbl)
m = p.merge(s, on=KEY, how="outer", suffixes=("_prod", "_systest"), indicator=True)



rows = []
for d in m.to_dict("records"):
    base = {k: d[k] for k in KEY}
    pv = {a: d.get(f"{a}_prod") for a in ATTRS}
    sv = {a: d.get(f"{a}_systest") for a in ATTRS}
    if d["_merge"] == "left_only":
        rows.append({**base, "change_type": "COLUMN_DROPPED", "detail": f"was {pv['data_type']}",
                     **{f"prod_{a}": pv[a] for a in ATTRS}, **{f"systest_{a}": None for a in ATTRS}})
        continue
    if d["_merge"] == "right_only":
        rows.append({**base, "change_type": "COLUMN_ADDED", "detail": f"new {sv['data_type']}",
                     **{f"prod_{a}": None for a in ATTRS}, **{f"systest_{a}": sv[a] for a in ATTRS}})
        continue
    changes = []
    if d["base_type_prod"] != d["base_type_systest"]:
        changes.append("DATA_TYPE_CHANGED")
    if pv["length"] != sv["length"]:
        changes.append("LENGTH_INCREASED" if sv["length"] > pv["length"] else "LENGTH_DECREASED")
    if pv["precision"] != sv["precision"]:
        changes.append("PRECISION_INCREASED" if sv["precision"] > pv["precision"] else "PRECISION_DECREASED")
    if pv["scale"] != sv["scale"]:
        changes.append("SCALE_INCREASED" if sv["scale"] > pv["scale"] else "SCALE_DECREASED")
    if pv["is_nullable"] != sv["is_nullable"]:
        changes.append("NOW_NULLABLE" if sv["is_nullable"] == 1 else "NOW_NOT_NULL")
    if changes:
        rows.append({**base, "change_type": "|".join(changes),
                     "detail": f"{pv['data_type']} -> {sv['data_type']}",
                     **{f"prod_{a}": pv[a] for a in ATTRS}, **{f"systest_{a}": sv[a] for a in ATTRS}})

detail = pd.DataFrame(rows)
for c in ("prod_length", "systest_length"):
    detail[c] = detail[c].replace(-1, "NULL")
detail["primary_change"] = detail["change_type"].str.split("|").str[0]
detail = detail.sort_values(["table_name", "primary_change", "column_name"])
cols = ["schema_name", "table_name", "column_name", "primary_change", "change_type", "detail"] + \
       [f"prod_{a}" for a in ATTRS] + [f"systest_{a}" for a in ATTRS]
detail = detail[cols]

# --- per-table summary
cnt = lambda pat: detail.assign(x=detail["change_type"].str.contains(pat)).groupby(tbl)["x"].sum()
lst = lambda pat: detail[detail["change_type"].str.contains(pat)].groupby(tbl)["column_name"].agg(", ".join)
summary = common.set_index(tbl)
summary["prod_column_count"] = p.groupby(tbl).size()
summary["systest_column_count"] = s.groupby(tbl).size()
for name, pat in [("added", "COLUMN_ADDED"), ("dropped", "COLUMN_DROPPED"),
                  ("type_changed", "DATA_TYPE_CHANGED"), ("length_changed", "LENGTH_"),
                  ("precision_scale_changed", "PRECISION_|SCALE_"), ("nullability_changed", "NULL")]:
    summary[f"cols_{name}"] = cnt(pat).reindex(summary.index).fillna(0).astype(int)
summary["added_columns"] = lst("COLUMN_ADDED").reindex(summary.index).fillna("")
summary["dropped_columns"] = lst("COLUMN_DROPPED").reindex(summary.index).fillna("")
summary = summary.reset_index()
summary = summary[summary.filter(like="cols_").sum(axis=1) > 0].sort_values("table_name")

# --- type-transition rollup (e.g. datetime -> datetime2(7))
mod = detail[~detail["primary_change"].isin(["COLUMN_ADDED", "COLUMN_DROPPED"])]
transitions = (mod.groupby(["prod_data_type", "systest_data_type", "change_type"]).size()
               .rename("column_count").reset_index().sort_values("column_count", ascending=False))

# --- overall stats
stats = pd.DataFrame([
    ("Prod tables", len(prod_tables)),
    ("SysTest tables", len(sys_tables)),
    ("Prod tables present in SysTest", len(common)),
    ("Prod tables MISSING in SysTest", len(dropped_tables)),
    ("New SysTest-only tables (ignored)", len(sys_tables) - len(common)),
    ("Common tables with any column change", len(summary)),
    ("Common tables with columns added", int((summary.cols_added > 0).sum())),
    ("Common tables with columns dropped", int((summary.cols_dropped > 0).sum())),
    ("Columns added (to existing Prod tables)", int((detail.primary_change == "COLUMN_ADDED").sum())),
    ("Columns dropped (from existing Prod tables)", int((detail.primary_change == "COLUMN_DROPPED").sum())),
    ("Columns with data type change", int(detail.change_type.str.contains("DATA_TYPE_CHANGED").sum())),
    ("Columns with length change", int(detail.change_type.str.contains("LENGTH_").sum())),
    ("  of which length decreased", int(detail.change_type.str.contains("LENGTH_DECREASED").sum())),
    ("Columns with precision/scale change", int(detail.change_type.str.contains("PRECISION_|SCALE_").sum())),
    ("Columns with nullability change", int(detail.change_type.str.contains("NOW_NULLABLE|NOW_NOT_NULL").sum())),
], columns=["metric", "value"])

stats.to_csv(out / "00_summary_stats.csv", index=False)
summary.to_csv(out / "01_table_summary.csv", index=False)
detail.to_csv(out / "02_column_changes_detail.csv", index=False)
transitions.to_csv(out / "03_datatype_transitions.csv", index=False)
dropped_tables.to_csv(out / "04_prod_tables_missing_in_systest.csv", index=False)

with pd.ExcelWriter(out / "MAS_Prod_vs_SysTest_Compare.xlsx") as xw:
    for n, df in [("Stats", stats), ("TableSummary", summary), ("ColumnChanges", detail),
                  ("TypeTransitions", transitions), ("ProdTablesMissing", dropped_tables)]:
        df.to_excel(xw, sheet_name=n, index=False)

print(stats.to_string(index=False))
