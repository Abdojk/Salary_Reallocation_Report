"""Salary Reallocation Report Generator for Engicon (WNGJ / C01)."""

import os
import sys
import glob
import calendar
import html
from datetime import datetime

import pandas as pd


# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------

TRANSCODE_ORDER = [95, 1, 2, 6, 55, 54, 53, 7]

TRANSCODE_LABELS = {
    95: ("اجازة امومه شهرية", "Maternity Leave"),
    1:  ("الرواتب و الاجور", "Salaries & Wages"),
    2:  ("بدل تنقلات", "Transportation Allowance"),
    6:  ("بدلات اخرى", "Other Allowances"),
    55: ("تامين صحي", "Health Insurance"),
    54: ("صندوق ادخار", "Saving Fund"),
    53: ("ضمان اجتماعي", "Social Security"),
    7:  ("علاوة موقع", "Site Allowance"),
}

ARABIC_MONTHS = {
    1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل",
    5: "مايو", 6: "يونيو", 7: "يوليو", 8: "أغسطس",
    9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر",
}

ENGLISH_MONTHS = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# -----------------------------------------------------------------------------
# File resolution
# -----------------------------------------------------------------------------

def resolve_file(patterns, label):
    matches = []
    for pat in patterns:
        matches.extend(glob.glob(os.path.join(BASE_DIR, pat)))
    # Dedupe while preserving order
    seen = set()
    unique = []
    for m in matches:
        if m not in seen:
            seen.add(m)
            unique.append(m)

    if len(unique) == 0:
        shown = " or ".join(patterns)
        print(f"ERROR: No {label} file matching {shown} found in script directory.")
        sys.exit(1)
    if len(unique) == 1:
        print(f"{label} file resolved: {os.path.basename(unique[0])}")
        return unique[0]

    print(f"Multiple {label} files found — select one:")
    for i, path in enumerate(unique, 1):
        print(f"  {i}. {os.path.basename(path)}")
    while True:
        choice = input("Enter number: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(unique):
            selected = unique[int(choice) - 1]
            print(f"{label} file resolved: {os.path.basename(selected)}")
            return selected
        print("Invalid selection. Try again.")


# -----------------------------------------------------------------------------
# Runtime prompts
# -----------------------------------------------------------------------------

def prompt_month_year():
    raw_month = input("Enter Month (1-12): ").strip()
    try:
        month = int(raw_month)
    except ValueError:
        print("ERROR: Month must be an integer between 1 and 12.")
        sys.exit(1)
    if not (1 <= month <= 12):
        print("ERROR: Month must be an integer between 1 and 12.")
        sys.exit(1)

    raw_year = input("Enter Year (e.g. 2026): ").strip()
    try:
        year = int(raw_year)
    except ValueError:
        print("ERROR: Year must be a four-digit integer between 1900 and 2100.")
        sys.exit(1)
    if not (1900 <= year <= 2100):
        print("ERROR: Year must be a four-digit integer between 1900 and 2100.")
        sys.exit(1)
    return month, year


# -----------------------------------------------------------------------------
# Loaders & filters
# -----------------------------------------------------------------------------

def load_and_filter_ins(path, month, year):
    df = pd.read_excel(path, engine="openpyxl")

    df["TransCode"] = pd.to_numeric(df["TransCode"], errors="coerce")
    df["Month"] = pd.to_numeric(df["Month"], errors="coerce")
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")

    keep = (
        (df["CompanyCode"] == "C01")
        & (df["Month"] == month)
        & (df["Year"] == year)
        & (df["ProjectCost"] == "Yes")
        & (df["TransCode"].isin([1, 2, 6, 7, 53, 54, 55, 95]))
    )
    df = df.loc[keep].copy()

    before = len(df)
    df = df.dropna(subset=["PersonnelNumber"])
    dropped = before - len(df)
    if dropped > 0:
        print(f"WARNING: Dropped {dropped} INS rows with null PersonnelNumber.")

    df["PersonnelNumber"] = df["PersonnelNumber"].astype(str).str.strip()
    df["TransCode"] = df["TransCode"].astype(int)
    print(f"INS filtered rows: {len(df)}")
    return df


def load_and_filter_hours(path, month, year):
    df = pd.read_excel(path, engine="openpyxl")
    df["Project date"] = pd.to_datetime(df["Project date"], errors="coerce")

    keep = (
        (df["Project date"].dt.month == month)
        & (df["Project date"].dt.year == year)
        & (df["Category"].isin(["Hour", "Hours"]))
    )
    df = df.loc[keep].copy()
    df["Resource.Resource ID"] = df["Resource.Resource ID"].astype(str).str.strip()
    df["Project ID"] = df["Project ID"].astype(str).str.strip()
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0.0)
    df["Cost price"] = pd.to_numeric(df["Cost price"], errors="coerce").fillna(0.0)
    print(f"Hours filtered rows: {len(df)}")
    return df


# -----------------------------------------------------------------------------
# Allocation engine
# -----------------------------------------------------------------------------

def build_worker_project_detail(ins_df, hours_df):
    # Step 1 — Worker actual cost per TransCode (absolute values)
    ins_df = ins_df.copy()
    ins_df["AbsCost"] = ins_df["CostAmount"].abs()
    worker_tc = (
        ins_df.groupby(["PersonnelNumber", "TransCode"])["AbsCost"]
        .sum()
        .unstack(fill_value=0.0)
    )
    # Ensure every required TC column exists
    for tc in TRANSCODE_ORDER:
        if tc not in worker_tc.columns:
            worker_tc[tc] = 0.0
    worker_tc = worker_tc[TRANSCODE_ORDER]
    worker_actual_total = worker_tc.sum(axis=1)

    # Step 2 — Worker hours per project
    project_hours = (
        hours_df.groupby(["Resource.Resource ID", "Project ID"])["Hours"]
        .sum()
        .reset_index()
    )
    worker_total_hours = (
        hours_df.groupby("Resource.Resource ID")["Hours"].sum()
    )

    # Standard cost per worker-project (Σ Hours × Cost price)
    hours_df = hours_df.copy()
    hours_df["LineStandard"] = hours_df["Hours"] * hours_df["Cost price"]
    standard_per_wp = (
        hours_df.groupby(["Resource.Resource ID", "Project ID"])["LineStandard"]
        .sum()
    )

    # Worker name lookup
    worker_names = (
        hours_df.groupby("Resource.Resource ID")["Resource.Resource name"]
        .first()
    )

    # Step 3 — Identify unallocated workers
    ins_workers = set(ins_df["PersonnelNumber"].unique())
    hours_workers_with_time = set(
        worker_total_hours[worker_total_hours > 0].index
    )
    allocated_workers = sorted(ins_workers & hours_workers_with_time)
    unallocated_workers = sorted(ins_workers - hours_workers_with_time)

    if unallocated_workers:
        print(f"Unallocated workers: {', '.join(unallocated_workers)}")
    else:
        print("Unallocated workers: none")

    # Step 4 + 5 — Build per worker-project detail rows
    detail_rows = []
    for worker in allocated_workers:
        total_hours = worker_total_hours.get(worker, 0.0)
        if total_hours <= 0:
            print(f"WARNING: worker {worker} has zero total hours after filter; skipping.")
            continue
        actual_total_w = worker_actual_total.get(worker, 0.0)
        worker_projects = project_hours[
            project_hours["Resource.Resource ID"] == worker
        ]
        for _, prow in worker_projects.iterrows():
            project_id = prow["Project ID"]
            hrs = float(prow["Hours"])
            ratio = hrs / total_hours
            std_total_wp = float(
                standard_per_wp.get((worker, project_id), 0.0)
            )
            row = {
                "PersonnelNumber": worker,
                "WorkerName": worker_names.get(worker, ""),
                "ProjectID": project_id,
                "total_hours": hrs,
                "actual_total": float(actual_total_w) * ratio,
                "standard_total": std_total_wp,
            }
            row["variance_total"] = row["standard_total"] - row["actual_total"]
            for tc in TRANSCODE_ORDER:
                actual_tc_w = float(worker_tc.at[worker, tc]) if worker in worker_tc.index else 0.0
                actual_tc_wp = actual_tc_w * ratio
                if actual_total_w > 0:
                    standard_tc_wp = std_total_wp * (actual_tc_w / actual_total_w)
                else:
                    standard_tc_wp = 0.0
                row[f"actual_{tc}"] = actual_tc_wp
                row[f"standard_{tc}"] = standard_tc_wp
                row[f"variance_{tc}"] = standard_tc_wp - actual_tc_wp
            detail_rows.append(row)

    detail_df = pd.DataFrame(detail_rows)

    # Step 6 — Unallocated synthetic row
    unallocated_row = None
    if unallocated_workers:
        unalloc_row = {
            "PersonnelNumber": "UNALLOCATED",
            "WorkerName": "غير مخصص / Unallocated",
            "ProjectID": "—",
            "total_hours": 0.0,
            "actual_total": 0.0,
            "standard_total": 0.0,
            "variance_total": 0.0,
        }
        sum_actual_total = 0.0
        for tc in TRANSCODE_ORDER:
            actual_sum = 0.0
            for w in unallocated_workers:
                if w in worker_tc.index:
                    actual_sum += float(worker_tc.at[w, tc])
            unalloc_row[f"actual_{tc}"] = actual_sum
            unalloc_row[f"standard_{tc}"] = 0.0
            unalloc_row[f"variance_{tc}"] = -actual_sum
            sum_actual_total += actual_sum
        unalloc_row["actual_total"] = sum_actual_total
        unalloc_row["variance_total"] = -sum_actual_total
        unallocated_row = unalloc_row

    return detail_df, unallocated_row, len(allocated_workers), len(unallocated_workers)


# -----------------------------------------------------------------------------
# Aggregations
# -----------------------------------------------------------------------------

def numeric_columns():
    cols = []
    for tc in TRANSCODE_ORDER:
        cols.extend([f"standard_{tc}", f"actual_{tc}", f"variance_{tc}"])
    cols.extend(["total_hours", "standard_total", "actual_total", "variance_total"])
    return cols


def aggregate_by_project(detail_df, unallocated_row):
    cols = numeric_columns()
    if detail_df.empty:
        project_df = pd.DataFrame(columns=["ProjectID"] + cols)
    else:
        project_df = (
            detail_df.groupby("ProjectID")[cols].sum().reset_index()
        )
        project_df = project_df.sort_values("ProjectID").reset_index(drop=True)

    rows = []
    for _, r in project_df.iterrows():
        row = {"ProjectID": r["ProjectID"]}
        for c in cols:
            row[c] = float(r[c])
        row["_type"] = "data"
        rows.append(row)

    if unallocated_row is not None:
        urow = {"ProjectID": unallocated_row["ProjectID"]}
        for c in cols:
            urow[c] = float(unallocated_row[c])
        urow["_type"] = "unallocated"
        rows.append(urow)

    # Grand total
    grand = {"ProjectID": "Grand Total / المجموع الكلي"}
    for c in cols:
        grand[c] = sum(r[c] for r in rows)
    grand["_type"] = "grand"
    rows.append(grand)
    return rows


def build_worker_rows(detail_df, unallocated_row):
    cols = numeric_columns()
    rows = []
    if not detail_df.empty:
        sorted_df = detail_df.sort_values(["PersonnelNumber", "ProjectID"])
        for worker, grp in sorted_df.groupby("PersonnelNumber", sort=False):
            grp = grp.sort_values("ProjectID")
            worker_name = grp["WorkerName"].iloc[0]
            project_rows = []
            for _, r in grp.iterrows():
                pr = {"ProjectID": r["ProjectID"]}
                for c in cols:
                    pr[c] = float(r[c])
                project_rows.append(pr)
            subtotal = {"ProjectID": f"Subtotal / مجموع {worker}"}
            for c in cols:
                subtotal[c] = sum(pr[c] for pr in project_rows)
            rows.append({
                "_type": "worker_group",
                "worker_id": worker,
                "worker_name": worker_name,
                "projects": project_rows,
                "subtotal": subtotal,
            })

    if unallocated_row is not None:
        rows.append({"_type": "unallocated_worker", "row": unallocated_row})

    # Grand total computed over allocated subtotals + unallocated
    grand = {"ProjectID": "Grand Total / المجموع الكلي"}
    for c in cols:
        total = 0.0
        for r in rows:
            if r["_type"] == "worker_group":
                total += r["subtotal"][c]
            elif r["_type"] == "unallocated_worker":
                total += float(r["row"][c])
        grand[c] = total
    rows.append({"_type": "grand", "row": grand})
    return rows


# -----------------------------------------------------------------------------
# HTML rendering
# -----------------------------------------------------------------------------

def format_num(value, decimals):
    try:
        v = float(value)
    except (TypeError, ValueError):
        return ""
    v = round(v, decimals)
    if decimals <= 0:
        return f"{v:,.0f}"
    return f"{v:,.{decimals}f}"


def variance_cell(value, extra_class=""):
    try:
        v = float(value)
    except (TypeError, ValueError):
        v = 0.0
    style = ' style="color:#cc0000"' if round(v, 3) < 0 else ""
    cls = f' class="{extra_class}"' if extra_class else ""
    return f'<td{cls}{style}>{format_num(v, 3)}</td>'


def money_cell(value, extra_class=""):
    cls = f' class="{extra_class}"' if extra_class else ""
    return f'<td{cls}>{format_num(value, 3)}</td>'


def hours_cell(value, extra_class=""):
    cls = f' class="{extra_class}"' if extra_class else ""
    return f'<td{cls}>{format_num(value, 2)}</td>'


def text_cell(value, extra_class="", colspan=None, rowspan=None):
    attrs = ""
    if extra_class:
        attrs += f' class="{extra_class}"'
    if colspan:
        attrs += f' colspan="{colspan}"'
    if rowspan:
        attrs += f' rowspan="{rowspan}"'
    return f'<td{attrs}>{html.escape(str(value))}</td>'


STYLE_BLOCK = """
<style>
    body { font-family: Arial, sans-serif; margin: 20px; color: #222; }
    .title-block { margin-bottom: 12px; }
    .title-main { font-size: 20px; font-weight: bold; margin-bottom: 4px; }
    .title-sub { font-size: 14px; margin-bottom: 2px; }
    .title-small { font-size: 12px; color: #555; margin-bottom: 12px; }
    .filter-line { font-size: 12px; color: #333; margin-bottom: 10px; }
    table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
    th, td { border: 1px solid #cccccc; padding: 4px 8px; }
    th { background-color: #2e4a7a; color: #ffffff; font-weight: bold;
         text-align: center; font-size: 12px; }
    td { font-size: 11px; text-align: right; }
    td.text { text-align: left; }
    tr.even td { background-color: #f9f9f9; }
    tr.odd td { background-color: #ffffff; }
    tr.subtotal td { background-color: #dce6f1; font-weight: bold; }
    tr.unallocated td { background-color: #fff3cd; font-weight: bold; }
    tr.grand td { background-color: #2e4a7a; color: #ffffff; font-weight: bold; }
    tr.grand td[style*="color:#cc0000"] { color: #ffb3b3 !important; }
    .footer { margin-top: 12px; font-size: 11px; color: #888; }
</style>
"""


def render_header(extra_left_columns):
    """Header HTML for both reports.

    extra_left_columns: list of column header labels rendered with rowspan=2
    before the TransCode group columns.
    """
    row1_cells = []
    for label in extra_left_columns:
        row1_cells.append(f'<th rowspan="2">{label}</th>')
    for tc in TRANSCODE_ORDER:
        ar, en = TRANSCODE_LABELS[tc]
        row1_cells.append(
            f'<th colspan="3">{html.escape(ar)}<br>{html.escape(en)}</th>'
        )
    row1_cells.append('<th rowspan="2">إجمالي الساعات<br>Total Hours</th>')
    row1_cells.append('<th rowspan="2">إجمالي معياري<br>Total Standard</th>')
    row1_cells.append('<th rowspan="2">إجمالي فعلي<br>Total Actual</th>')
    row1_cells.append('<th rowspan="2">إجمالي الفرق<br>Total Variance</th>')

    row2_cells = []
    for _ in TRANSCODE_ORDER:
        row2_cells.append('<th>معيار / Standard</th>')
        row2_cells.append('<th>فعلي / Actual</th>')
        row2_cells.append('<th>الفرق / Variance</th>')

    return (
        "<thead>"
        f"<tr>{''.join(row1_cells)}</tr>"
        f"<tr>{''.join(row2_cells)}</tr>"
        "</thead>"
    )


def render_data_cells(row):
    cells = []
    for tc in TRANSCODE_ORDER:
        cells.append(money_cell(row[f"standard_{tc}"]))
        cells.append(money_cell(row[f"actual_{tc}"]))
        cells.append(variance_cell(row[f"variance_{tc}"]))
    cells.append(hours_cell(row["total_hours"]))
    cells.append(money_cell(row["standard_total"]))
    cells.append(money_cell(row["actual_total"]))
    cells.append(variance_cell(row["variance_total"]))
    return "".join(cells)


def render_project_report(rows, month, year):
    header = render_header(["معرف المشروع<br>Project ID"])
    body_rows = []
    parity = 0
    for r in rows:
        if r["_type"] == "data":
            cls = "even" if parity % 2 == 0 else "odd"
            parity += 1
            body_rows.append(
                f'<tr class="{cls}">'
                f'{text_cell(r["ProjectID"], extra_class="text")}'
                f'{render_data_cells(r)}'
                f'</tr>'
            )
        elif r["_type"] == "unallocated":
            body_rows.append(
                f'<tr class="unallocated">'
                f'{text_cell(r["ProjectID"], extra_class="text")}'
                f'{render_data_cells(r)}'
                f'</tr>'
            )
        elif r["_type"] == "grand":
            body_rows.append(
                f'<tr class="grand">'
                f'{text_cell(r["ProjectID"], extra_class="text")}'
                f'{render_data_cells(r)}'
                f'</tr>'
            )
    body = "<tbody>" + "".join(body_rows) + "</tbody>"
    return build_document(
        title_en="Salary Reallocation Report — By Project",
        title_ar="تقرير إعادة توزيع الرواتب — حسب المشروع",
        table_html=f"<table>{header}{body}</table>",
        month=month,
        year=year,
    )


def render_worker_report(rows, month, year):
    header = render_header([
        "اسم الموظف<br>Worker Name",
        "معرف المشروع<br>Project ID",
    ])
    body_rows = []
    parity = 0
    for r in rows:
        if r["_type"] == "worker_group":
            project_rows = r["projects"]
            subtotal = r["subtotal"]
            span = len(project_rows) + 1
            worker_label = f'{html.escape(str(r["worker_id"]))} — {html.escape(str(r["worker_name"]))}'
            for i, pr in enumerate(project_rows):
                cls = "even" if parity % 2 == 0 else "odd"
                parity += 1
                cells = ""
                if i == 0:
                    cells += f'<td class="text" rowspan="{span}">{worker_label}</td>'
                cells += text_cell(pr["ProjectID"], extra_class="text")
                cells += render_data_cells(pr)
                body_rows.append(f'<tr class="{cls}">{cells}</tr>')
            body_rows.append(
                f'<tr class="subtotal">'
                f'{text_cell(subtotal["ProjectID"], extra_class="text")}'
                f'{render_data_cells(subtotal)}'
                f'</tr>'
            )
        elif r["_type"] == "unallocated_worker":
            urow = r["row"]
            body_rows.append(
                f'<tr class="unallocated">'
                f'{text_cell(urow["WorkerName"], extra_class="text")}'
                f'{text_cell(urow["ProjectID"], extra_class="text")}'
                f'{render_data_cells(urow)}'
                f'</tr>'
            )
        elif r["_type"] == "grand":
            grand = r["row"]
            body_rows.append(
                f'<tr class="grand">'
                f'{text_cell("Grand Total / المجموع الكلي", extra_class="text", colspan=2)}'
                f'{render_data_cells(grand)}'
                f'</tr>'
            )
    body = "<tbody>" + "".join(body_rows) + "</tbody>"
    return build_document(
        title_en="Salary Reallocation Report — By Worker by Project",
        title_ar="تقرير إعادة توزيع الرواتب — حسب الموظف والمشروع",
        table_html=f"<table>{header}{body}</table>",
        month=month,
        year=year,
    )


def build_document(title_en, title_ar, table_html, month, year):
    ar_month = ARABIC_MONTHS[month]
    en_month = ENGLISH_MONTHS[month]
    generated = datetime.now().strftime("%d %b %Y %H:%M")
    return (
        '<html dir="auto">\n'
        '<head>\n'
        '<meta charset="UTF-8">\n'
        f'<title>{html.escape(title_ar)} | {html.escape(title_en)} — {en_month} {year}</title>\n'
        f'{STYLE_BLOCK}\n'
        '</head>\n'
        '<body>\n'
        '<div class="title-block">\n'
        '<div class="title-main">تقرير إعادة توزيع الرواتب | Salary Reallocation Report</div>\n'
        f'<div class="title-sub">{ar_month} {year} | {en_month} {year}</div>\n'
        '<div class="title-small">الشركة / Company: Engicon (WNGJ) — CompanyCode: C01</div>\n'
        '</div>\n'
        f'<div class="filter-line">Filter: Month={month} · Year={year} · CompanyCode=C01 · ProjectCost=Yes</div>\n'
        f'{table_html}\n'
        f'<div class="footer">Generated: {generated}</div>\n'
        '</body>\n'
        '</html>\n'
    )


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main():
    ins_path = resolve_file(["INS_PayrollEmplTrans_*.xlsx"], "INS")
    hours_path = resolve_file(
        ["Project_hour_transactions_*.xlsx", "Project hour transactions_*.xlsx"],
        "Hours",
    )

    month, year = prompt_month_year()

    ins_df = load_and_filter_ins(ins_path, month, year)
    hours_df = load_and_filter_hours(hours_path, month, year)

    detail_df, unallocated_row, n_alloc, n_unalloc = build_worker_project_detail(
        ins_df, hours_df
    )

    project_rows = aggregate_by_project(detail_df, unallocated_row)
    worker_rows = build_worker_rows(detail_df, unallocated_row)

    project_html = render_project_report(project_rows, month, year)
    worker_html = render_worker_report(worker_rows, month, year)

    mmm = calendar.month_abbr[month]
    yyyy = str(year)
    out1 = os.path.join(BASE_DIR, f"salary_reallocation_by_project_{mmm}_{yyyy}.html")
    out2 = os.path.join(BASE_DIR, f"salary_reallocation_by_worker_{mmm}_{yyyy}.html")
    with open(out1, "w", encoding="utf-8") as f:
        f.write(project_html)
    with open(out2, "w", encoding="utf-8") as f:
        f.write(worker_html)

    project_count = len({r["ProjectID"] for r in project_rows if r["_type"] == "data"})

    print(f"Report 1 written : {os.path.basename(out1)}")
    print(f"Report 2 written : {os.path.basename(out2)}")
    print(f"Allocated workers  : {n_alloc}")
    print(f"Unallocated workers: {n_unalloc}")
    print(f"Total projects     : {project_count}")


if __name__ == "__main__":
    main()
