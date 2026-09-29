"""Compare Maximo (MAS) column schema: Prod vs SysTest.

Scope: only tables that exist in Prod.
  - columns added / dropped on tables present in both
  - data type, length, precision, scale, nullability changes
  - Prod tables missing from SysTest
New SysTest-only tables are ignored.

Output: one workbook, MAS_Prod_vs_SysTest_Compare.xlsx (Stats, TableSummary,
ColumnChanges, TypeTransitions, ProdTablesMissing tabs).

Usage: python compare_schema.py <MAS_Upgrade.xlsx> [out_dir] [--tables list.txt]
  --tables  restrict the comparison to the table names in list.txt (one per line);
            output is then MAS_Prod_vs_SysTest_Compare_<list name>.xlsx
"""
import argparse
from pathlib import Path

import pandas as pd

ap = argparse.ArgumentParser()
ap.add_argument("src")
ap.add_argument("out_dir", nargs="?", default=".")
ap.add_argument("--tables")
args = ap.parse_args()
src, out = args.src, Path(args.out_dir)
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
wanted = None
if args.tables:
    wanted = sorted({l.strip().lower() for l in open(args.tables) if l.strip()})
    prod = prod[prod.table_name.isin(wanted)]
    sys_ = sys_[sys_.table_name.isin(wanted)]
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
# --- overall stats: counts derivable from the other tabs are live COUNTIF formulas
CC, TS = "ColumnChanges", "TableSummary"
stats = [
    ("Prod tables", len(prod_tables), "Value from Prod sheet of source workbook"),
    ("SysTest tables", len(sys_tables), "Value from SysTest sheet of source workbook"),
    ("Prod tables present in SysTest", len(common), "Value from source workbook"),
    ("Prod tables MISSING in SysTest", "=COUNTA(ProdTablesMissing!B:B)-1", "Rows on ProdTablesMissing tab"),
    ("New SysTest-only tables (out of scope)", len(sys_tables) - len(common), "Value from source workbook"),
    ("Existing tables with any column change", f"=COUNTA({TS}!B:B)-1", "Rows on TableSummary tab"),
    ("Existing tables with columns added", f"=COUNTIF({TS}!E:E,\">0\")", "TableSummary cols_added > 0"),
    ("Existing tables with columns dropped", f"=COUNTIF({TS}!F:F,\">0\")", "TableSummary cols_dropped > 0"),
    ("Columns added", f"=COUNTIF({CC}!D:D,\"COLUMN_ADDED\")", "ColumnChanges primary_change"),
    ("Columns dropped", f"=COUNTIF({CC}!D:D,\"COLUMN_DROPPED\")", "ColumnChanges primary_change"),
    ("Columns with data type change", f"=COUNTIF({CC}!E:E,\"*DATA_TYPE_CHANGED*\")", "ColumnChanges change_type"),
    ("Columns with length change", f"=COUNTIF({CC}!E:E,\"*LENGTH_*\")", "ColumnChanges change_type"),
    ("  of which length decreased", f"=COUNTIF({CC}!E:E,\"*LENGTH_DECREASED*\")", "ColumnChanges change_type"),
    ("Columns with precision/scale change",
     f"=SUMPRODUCT(--((ISNUMBER(SEARCH(\"PRECISION_\",{CC}!E2:E{len(detail)+1})))+(ISNUMBER(SEARCH(\"SCALE_\",{CC}!E2:E{len(detail)+1})))>0))",
     "ColumnChanges change_type"),
    ("Columns with nullability change", f"=COUNTIF({CC}!E:E,\"*NOW_N*\")", "ColumnChanges change_type"),
]
notes = [
    "Scope: only tables that exist in Prod. New SysTest-only tables are excluded.",
    "Table/column names matched case-insensitively (SQL Server collation).",
    "primary_change = first item of change_type; change_type lists every change on the column, pipe-separated.",
    "Length 'NULL' = non-character type; -1 in source = (max).",
    "Tables missing from SysTest are listed on ProdTablesMissing, not repeated as dropped columns.",
]
if wanted:
    in_prod = set(prod_tables.table_name)
    changed = set(summary.table_name) | set(dropped_tables.table_name)
    notes.insert(0, f"Subset: limited to {len(wanted)} requested tables ({Path(args.tables).name}).")
    notes.append("Requested tables with no schema change: "
                 + (", ".join(t for t in wanted if t in in_prod and t not in changed) or "none"))
    notes.append("Requested tables not found in Prod: "
                 + (", ".join(t for t in wanted if t not in in_prod) or "none"))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONT, BOLD = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True, color="FFFFFF")
HDR = PatternFill("solid", fgColor="1F4E78")

wb = Workbook()
ws = wb.active
ws.title = "Stats"
ws.append(["Metric", "Value", "Source"])
for r in stats:
    ws.append(list(r))
ws.append([])
ws.append(["Notes"])
for n in notes:
    ws.append([n])
ws.cell(ws.max_row - len(notes), 1).font = Font(name="Arial", size=10, bold=True)

for name, df in [(TS, summary), (CC, detail), ("TypeTransitions", transitions),
                 ("ProdTablesMissing", dropped_tables)]:
    sh = wb.create_sheet(name)
    sh.append(list(df.columns))
    for row in df.astype(object).where(df.notna(), None).itertuples(index=False):
        sh.append(list(row))

for sh in wb.worksheets:
    last_col = get_column_letter(sh.max_column)
    for row in sh.iter_rows():
        for c in row:
            c.font = FONT
    for c in sh[1]:
        c.font, c.fill = BOLD, HDR
        c.alignment = Alignment(wrap_text=True, vertical="center")
    sh.freeze_panes = "A2"
    if sh.title != "Stats":
        sh.auto_filter.ref = f"A1:{last_col}{sh.max_row}"
    for i, col in enumerate(sh.iter_cols(min_row=1, max_row=min(sh.max_row, 500)), 1):
        w = max((len(str(c.value)) for c in col if c.value is not None and not str(c.value).startswith("=")), default=8)
        sh.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 10), 60)
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 12

suffix = f"_{Path(args.tables).stem}" if wanted else ""
xlsx = out / f"MAS_Prod_vs_SysTest_Compare{suffix}.xlsx"
wb.calculation.fullCalcOnLoad = True  # Excel computes the Stats formulas on open
wb.save(xlsx)
print(f"Wrote {xlsx}")
