"""Generate a monthly Salary Reallocation Report.

Reads payroll transactions and project hour transactions from the local
spreadsheets, prompts for Month and Year, and writes a styled HTML report
that reallocates each employee's monthly salary cost across projects in
proportion to the hours logged.
"""

from __future__ import annotations

import calendar
import datetime as _dt
import html
import os
import sys

try:
    import pandas as pd
except ImportError:
    print("Missing dependency: pandas. Install with:\n    pip install pandas openpyxl")
    sys.exit(1)

try:
    import openpyxl  # noqa: F401
except ImportError:
    print("Missing dependency: openpyxl. Install with:\n    pip install pandas openpyxl")
    sys.exit(1)


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PAYROLL_PATH = os.path.join(SCRIPT_DIR, "INS_PayrollEmplTrans_20260519.xlsx")
HOURS_PATH = os.path.join(SCRIPT_DIR, "Project hour transactions_JantoApr2026.xlsx")


def prompt_int(label: str, low: int, high: int) -> int:
    while True:
        raw = input(label).strip()
        try:
            value = int(raw)
        except ValueError:
            print(f"  Please enter a whole number between {low} and {high}.")
            continue
        if value < low or value > high:
            print(f"  Value must be between {low} and {high}.")
            continue
        return value


def normalize_id(series: pd.Series) -> pd.Series:
    """Coerce IDs to a consistent string form so payroll and hours can join."""
    def _one(v):
        if pd.isna(v):
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v).strip()
    return series.map(_one)


def fmt_money(x: float) -> str:
    return f"{x:,.2f}"


def fmt_hours(x: float) -> str:
    return f"{x:,.2f}"


def fmt_pct(x: float) -> str:
    return f"{x * 100:,.2f}%"


def build_allocation(month: int, year: int):
    payroll = pd.read_excel(PAYROLL_PATH, sheet_name="Sheet1")
    payroll = payroll[(payroll["Month"] == month) & (payroll["Year"] == year)].copy()
    payroll["PersonnelNumber"] = normalize_id(payroll["PersonnelNumber"])

    if payroll.empty:
        return None, None, None

    cost_center = (
        payroll.dropna(subset=["EmplCostCenter"])
        .groupby("PersonnelNumber")["EmplCostCenter"]
        .first()
    )
    salary = (
        payroll.groupby("PersonnelNumber")["CostAmount"]
        .sum()
        .abs()
        .rename("TotalSalary")
    )
    salary_df = pd.concat([cost_center, salary], axis=1).reset_index()
    salary_df["EmplCostCenter"] = salary_df["EmplCostCenter"].fillna("")

    hours = pd.read_excel(HOURS_PATH, sheet_name="Sheet1")
    hours["Project date"] = pd.to_datetime(hours["Project date"], errors="coerce")
    mask = (hours["Project date"].dt.month == month) & (hours["Project date"].dt.year == year)
    hours = hours[mask].copy()
    hours["ResourceID"] = normalize_id(hours["Resource.Resource ID"])

    name_lookup = (
        hours.dropna(subset=["Resource.Resource name"])
        .drop_duplicates(subset=["ResourceID"])
        .set_index("ResourceID")["Resource.Resource name"]
        .to_dict()
    )

    hours_grp = (
        hours.groupby(["ResourceID", "Project ID"], as_index=False)["Hours"]
        .sum()
    )
    totals = hours_grp.groupby("ResourceID")["Hours"].sum().rename("TotalHours")
    hours_grp = hours_grp.join(totals, on="ResourceID")
    hours_grp = hours_grp[hours_grp["TotalHours"] > 0].copy()
    hours_grp["HoursPct"] = hours_grp["Hours"] / hours_grp["TotalHours"]

    merged = hours_grp.merge(
        salary_df, left_on="ResourceID", right_on="PersonnelNumber", how="inner"
    )
    merged["AllocatedCost"] = merged["TotalSalary"] * merged["HoursPct"]
    merged["ResourceName"] = merged["ResourceID"].map(name_lookup).fillna("")

    allocated = merged[
        [
            "PersonnelNumber",
            "ResourceName",
            "EmplCostCenter",
            "Project ID",
            "Hours",
            "HoursPct",
            "TotalSalary",
            "AllocatedCost",
        ]
    ].sort_values(
        ["PersonnelNumber", "Hours"], ascending=[True, False]
    ).reset_index(drop=True)

    allocated_ids = set(allocated["PersonnelNumber"].unique())
    unallocated_mask = ~salary_df["PersonnelNumber"].isin(allocated_ids)
    unallocated = salary_df[unallocated_mask].copy()
    unallocated["ResourceName"] = unallocated["PersonnelNumber"].map(name_lookup).fillna("")
    unallocated = unallocated[
        ["PersonnelNumber", "ResourceName", "EmplCostCenter", "TotalSalary"]
    ].sort_values("PersonnelNumber").reset_index(drop=True)

    return allocated, unallocated, salary_df


def render_html(allocated: pd.DataFrame, unallocated: pd.DataFrame, month: int, year: int) -> str:
    month_name = calendar.month_name[month]
    generated = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")

    grand_hours = allocated["Hours"].sum() if not allocated.empty else 0.0
    grand_alloc = allocated["AllocatedCost"].sum() if not allocated.empty else 0.0
    unalloc_total = unallocated["TotalSalary"].sum() if not unallocated.empty else 0.0

    css = """
    <style>
      body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
             margin: 24px; color: #222; }
      h1 { margin-bottom: 4px; }
      .meta { color: #666; margin-bottom: 24px; font-size: 0.9em; }
      h2 { border-bottom: 2px solid #2d6cdf; padding-bottom: 4px; margin-top: 32px; }
      table { border-collapse: collapse; width: 100%; margin-top: 8px;
              font-size: 0.92em; }
      thead th { background: #2d6cdf; color: #fff; padding: 8px 10px;
                 text-align: left; position: sticky; top: 0; }
      tbody td { padding: 6px 10px; border-bottom: 1px solid #eee; }
      tbody tr:nth-child(even) td { background: #f7f9fc; }
      td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
      tr.subtotal td { background: #eef3fb !important; font-weight: 600;
                       border-top: 1px solid #cfd9ea; }
      tr.grand td { background: #2d6cdf !important; color: #fff;
                    font-weight: 700; font-size: 1.02em; }
      .name { direction: rtl; unicode-bidi: plaintext; }
      .summary { margin-top: 6px; color: #444; }
    </style>
    """

    parts: list[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="en">')
    parts.append('<head><meta charset="utf-8">')
    parts.append(f"<title>Salary Reallocation {year}-{month:02d}</title>")
    parts.append(css)
    parts.append("</head><body>")
    parts.append(f"<h1>Salary Reallocation Report — {html.escape(month_name)} {year}</h1>")
    parts.append(f'<div class="meta">Generated {generated}</div>')

    distinct_emps = allocated["PersonnelNumber"].nunique() if not allocated.empty else 0
    distinct_projs = allocated["Project ID"].nunique() if not allocated.empty else 0
    parts.append(
        f'<div class="summary">Allocated: <b>{distinct_emps}</b> employees across '
        f"<b>{distinct_projs}</b> projects · Allocated total <b>{fmt_money(grand_alloc)}</b> · "
        f"Unallocated employees: <b>{len(unallocated)}</b> "
        f"(total <b>{fmt_money(unalloc_total)}</b>)</div>"
    )

    parts.append("<h2>Allocation by Employee &amp; Project</h2>")
    parts.append("<table>")
    parts.append(
        "<thead><tr>"
        "<th>PersonnelNumber</th>"
        "<th>Resource name</th>"
        "<th>EmplCostCenter</th>"
        "<th>Project ID</th>"
        '<th class="num">Hours</th>'
        '<th class="num">Hours %</th>'
        '<th class="num">Total Salary</th>'
        '<th class="num">Allocated Cost</th>'
        "</tr></thead><tbody>"
    )

    if allocated.empty:
        parts.append('<tr><td colspan="8" style="text-align:center;color:#888;">'
                     "No allocations for this period.</td></tr>")
    else:
        for emp_id, group in allocated.groupby("PersonnelNumber", sort=False):
            first = group.iloc[0]
            name = html.escape(str(first["ResourceName"]))
            cc = html.escape(str(first["EmplCostCenter"]))
            total_salary = float(first["TotalSalary"])
            for _, row in group.iterrows():
                parts.append("<tr>")
                parts.append(f"<td>{html.escape(str(row['PersonnelNumber']))}</td>")
                parts.append(f'<td class="name">{name}</td>')
                parts.append(f"<td>{cc}</td>")
                parts.append(f"<td>{html.escape(str(row['Project ID']))}</td>")
                parts.append(f'<td class="num">{fmt_hours(row["Hours"])}</td>')
                parts.append(f'<td class="num">{fmt_pct(row["HoursPct"])}</td>')
                parts.append(f'<td class="num">{fmt_money(total_salary)}</td>')
                parts.append(f'<td class="num">{fmt_money(row["AllocatedCost"])}</td>')
                parts.append("</tr>")
            sub_hours = group["Hours"].sum()
            sub_alloc = group["AllocatedCost"].sum()
            parts.append(
                f'<tr class="subtotal">'
                f"<td>{html.escape(str(emp_id))}</td>"
                f'<td class="name">{name}</td>'
                f"<td>{cc}</td>"
                f"<td>Subtotal</td>"
                f'<td class="num">{fmt_hours(sub_hours)}</td>'
                f'<td class="num">100.00%</td>'
                f'<td class="num">{fmt_money(total_salary)}</td>'
                f'<td class="num">{fmt_money(sub_alloc)}</td>'
                "</tr>"
            )
        parts.append(
            f'<tr class="grand">'
            '<td colspan="4">Grand Total</td>'
            f'<td class="num">{fmt_hours(grand_hours)}</td>'
            '<td class="num"></td>'
            '<td class="num"></td>'
            f'<td class="num">{fmt_money(grand_alloc)}</td>'
            "</tr>"
        )

    parts.append("</tbody></table>")

    parts.append("<h2>Unallocated Employees (no project hours this month)</h2>")
    parts.append("<table>")
    parts.append(
        "<thead><tr>"
        "<th>PersonnelNumber</th>"
        "<th>Resource name</th>"
        "<th>EmplCostCenter</th>"
        '<th class="num">Total Salary</th>'
        "</tr></thead><tbody>"
    )
    if unallocated.empty:
        parts.append('<tr><td colspan="4" style="text-align:center;color:#888;">'
                     "None — every payroll employee has hours this month.</td></tr>")
    else:
        for _, row in unallocated.iterrows():
            parts.append("<tr>")
            parts.append(f"<td>{html.escape(str(row['PersonnelNumber']))}</td>")
            parts.append(f'<td class="name">{html.escape(str(row["ResourceName"]))}</td>')
            parts.append(f"<td>{html.escape(str(row['EmplCostCenter']))}</td>")
            parts.append(f'<td class="num">{fmt_money(row["TotalSalary"])}</td>')
            parts.append("</tr>")
        parts.append(
            '<tr class="grand">'
            '<td colspan="3">Grand Total</td>'
            f'<td class="num">{fmt_money(unalloc_total)}</td>'
            "</tr>"
        )
    parts.append("</tbody></table>")

    parts.append("</body></html>")
    return "".join(parts)


def main() -> int:
    month = prompt_int("Enter Month (1-12): ", 1, 12)
    year = prompt_int("Enter Year (e.g. 2026): ", 2000, 2100)

    allocated, unallocated, salary_df = build_allocation(month, year)

    if salary_df is None:
        print(f"No payroll rows found for {calendar.month_name[month]} {year}. Nothing written.")
        return 0

    html_text = render_html(allocated, unallocated, month, year)
    out_path = os.path.join(SCRIPT_DIR, f"Salary_Reallocation_{year}_{month:02d}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_text)

    grand_alloc = allocated["AllocatedCost"].sum() if not allocated.empty else 0.0
    distinct_projs = allocated["Project ID"].nunique() if not allocated.empty else 0
    print(
        f"Wrote {os.path.basename(out_path)} — "
        f"{len(allocated)} allocated rows across {distinct_projs} projects, "
        f"{len(unallocated)} unallocated employees, "
        f"grand total {fmt_money(grand_alloc)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
