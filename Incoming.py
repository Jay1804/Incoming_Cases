import streamlit as st
import pandas as pd
import os
import tempfile
import shutil
import datetime
from concurrent.futures import ThreadPoolExecutor

import win32com.client as win32
import pythoncom

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from bm_mapping import load_bm_mapping, attach_bm_lead
from db import get_engine
from queries import fetch_incoming_case_count, fetch_sent_case

st.set_page_config(page_title="Excel Split & Email Sender", layout="centered")
st.title("📧 Excel Splitter and Email Sender")


def format_workbook(file_path):
    """Apply the standard header/border/width styling to every sheet in the workbook."""
    wb = load_workbook(file_path)

    header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = center_align
                cell.border = thin_border

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 18

    wb.save(file_path)


def bold_last_row(file_path, sheet_name):
    """Bold + shade the last row of a sheet (used for the Monthly tab's Grand Total)."""
    wb = load_workbook(file_path)
    ws = wb[sheet_name]

    total_fill = PatternFill(start_color="E2E2E2", end_color="E2E2E2", fill_type="solid")
    bold_font = Font(bold=True)

    for cell in ws[ws.max_row]:
        cell.font = bold_font
        cell.fill = total_fill

    wb.save(file_path)


ROLLING_DAYS = 30


def _slice_dates(df, date_col, start, end):
    dates = pd.to_datetime(df[date_col]).dt.date
    return df[(dates >= start) & (dates <= end)].reset_index(drop=True)


@st.cache_data(ttl=900, show_spinner=False)
def fetch_report_data(range_start, range_end, roll_start, roll_end):
    """Incoming/Sent data for the user's range and the rolling Summary window.
    The queries are the slow part (large tables), so: run them in parallel; if
    the two windows overlap, fetch their combined span once and slice it instead
    of querying twice; and cache the result for 15 minutes so a re-export is instant."""
    get_engine()  # create the shared engine before threads race to do it
    if range_start <= roll_end and roll_start <= range_end:
        span = (min(range_start, roll_start), max(range_end, roll_end))
        jobs = {"inc": (fetch_incoming_case_count, span), "sent": (fetch_sent_case, span)}
    else:
        jobs = {
            "inc": (fetch_incoming_case_count, (range_start, range_end)),
            "sent": (fetch_sent_case, (range_start, range_end)),
            "inc_roll": (fetch_incoming_case_count, (roll_start, roll_end)),
            "sent_roll": (fetch_sent_case, (roll_start, roll_end)),
        }
    with ThreadPoolExecutor(max_workers=len(jobs)) as pool:
        futures = {k: pool.submit(fn, *args) for k, (fn, args) in jobs.items()}
        res = {k: f.result() for k, f in futures.items()}

    if "inc_roll" not in res:
        res["inc_roll"], res["sent_roll"] = res["inc"], res["sent"]
    return (
        _slice_dates(res["inc"], "received_date", range_start, range_end),
        _slice_dates(res["sent"], "sent_date", range_start, range_end),
        _slice_dates(res["inc_roll"], "received_date", roll_start, roll_end),
        _slice_dates(res["sent_roll"], "sent_date", roll_start, roll_end),
    )


@st.cache_data(ttl=900, show_spinner=False)
def cached_bm_mapping():
    return load_bm_mapping()


def add_summary_sheet(file_path, incoming_df, sent_df, end_date, days=ROLLING_DAYS):
    """Insert a 'Summary' sheet (first tab) laid out like the 'Client-wise Daily
    Received vs Sent Cases' reference: one Received/Sent row pair per client with
    one column per day for the `days` days ending at `end_date` (D2..AG2 for 30),
    then Grand Total / Daily Run Rate / Net Gap. Values are written directly (not
    SUMIFS) so previews and mail clients show numbers without a recalculation.
    Returns False (sheet not added) when there is no client data in the window."""
    dates = [end_date - datetime.timedelta(days=days - 1 - i) for i in range(days)]

    inc = incoming_df.copy()
    sent = sent_df.copy()
    inc["received_date"] = pd.to_datetime(inc["received_date"]).dt.date
    sent["sent_date"] = pd.to_datetime(sent["sent_date"]).dt.date
    inc = inc[inc["received_date"].isin(dates)]
    sent = sent[sent["sent_date"].isin(dates)]

    recv = inc.groupby(["Client_Id", "received_date"])["Case Count"].sum().to_dict()
    sent_map = sent.groupby(["Client Code", "sent_date"])["case_count"].sum().to_dict()

    names = dict(zip(sent["Client Code"], sent["Client_name"]))
    names.update(dict(zip(inc["Client_Id"], inc["Client"])))
    clients = sorted(set(inc["Client_Id"]) | set(sent["Client Code"]))
    if not clients:
        return False

    wb = load_workbook(file_path)
    ws = wb.create_sheet("Summary", 0)
    ws.sheet_view.showGridLines = False

    navy, grey_text = "1F4E78", "595959"
    side = Side(style="thin", color="BFBFBF")
    border = Border(left=side, right=side, top=side, bottom=side)
    head_fill = PatternFill(start_color=navy, end_color=navy, fill_type="solid")
    sent_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    num_fmt = '#,##0;\\-#,##0;"-"'

    first_day_col = 4
    last_day_col = first_day_col + days - 1
    total_col, drr_col, gap_col = last_day_col + 1, last_day_col + 2, last_day_col + 3
    fd, ld = get_column_letter(first_day_col), get_column_letter(last_day_col)
    tl, drr_l = get_column_letter(total_col), get_column_letter(drr_col)

    ws["A1"] = "Client-wise Daily Received vs Sent Cases"
    ws["A1"].font = Font(size=13, bold=True, color=navy)
    info_font = Font(size=9, bold=True, color=grey_text)
    val_font = Font(bold=True, color=navy)
    asof_lbl, asof_val = get_column_letter(last_day_col - 2), get_column_letter(last_day_col - 1)
    days_lbl = get_column_letter(last_day_col)
    for addr, value, font, align, fmt in (
        (f"{asof_lbl}1", "Data as of:", info_font, Alignment(horizontal="right"), None),
        (f"{asof_val}1", end_date, val_font, center, "d-mmm-yy"),
        (f"{days_lbl}1", "Days:", info_font, Alignment(horizontal="right"), None),
        (f"{tl}1", days, val_font, center, None),
    ):
        ws[addr] = value
        ws[addr].font = font
        ws[addr].alignment = align
        if fmt:
            ws[addr].number_format = fmt

    headers = {1: "Client Code", 2: "Client Name", total_col: "Grand Total",
               drr_col: "Daily Run Rate", gap_col: "Net Gap (Received − Sent)"}
    for col in range(1, gap_col + 1):
        c = ws.cell(row=2, column=col)
        if col in headers:
            c.value = headers[col]
        elif first_day_col <= col <= last_day_col:
            c.value = dates[col - first_day_col]
            c.number_format = "d-mmm"
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = center
        c.border = border
    ws.row_dimensions[1].height = 17
    ws.row_dimensions[2].height = 30

    def write_pair(r, code, name, rec_vals, snt_vals, fill=None, bold=False):
        for offset, label, vals in ((0, "Received", rec_vals), (1, "Sent", snt_vals)):
            row = r + offset
            row_fill = fill or (sent_fill if offset else None)
            ws.cell(row=row, column=3, value=label)
            for i, v in enumerate(vals):
                ws.cell(row=row, column=first_day_col + i, value=v).number_format = num_fmt
            total = sum(vals)
            ws.cell(row=row, column=total_col, value=total).number_format = num_fmt
            drr = ws.cell(row=row, column=drr_col, value=round(total / days, 1))
            drr.number_format = "0.0"
            drr.font = Font(bold=True, color=navy)
            for col in range(1, gap_col + 1):
                c = ws.cell(row=row, column=col)
                c.border = border
                if col >= first_day_col:
                    c.alignment = center
                if row_fill:
                    c.fill = row_fill
                if bold and col != drr_col:
                    c.font = Font(bold=True)
            ws.cell(row=row, column=total_col).font = Font(bold=True)
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=name)
        gap = ws.cell(row=r, column=gap_col, value=sum(rec_vals) - sum(snt_vals))
        gap.number_format = num_fmt
        gap.font = Font(bold=True)
        gap.alignment = center
        for col in (1, 2, gap_col):
            ws.merge_cells(start_row=r, start_column=col, end_row=r + 1, end_column=col)
        ws.cell(row=r, column=1).font = Font(bold=True)
        ws.cell(row=r, column=1).alignment = center
        ws.cell(row=r, column=2).alignment = Alignment(vertical="center")
        ws.cell(row=r, column=gap_col).alignment = center

    day_totals_r = [0] * days
    day_totals_s = [0] * days
    r = 3
    for code in clients:
        rec_vals = [int(recv.get((code, d), 0)) for d in dates]
        snt_vals = [int(sent_map.get((code, d), 0)) for d in dates]
        day_totals_r = [a + b for a, b in zip(day_totals_r, rec_vals)]
        day_totals_s = [a + b for a, b in zip(day_totals_s, snt_vals)]
        write_pair(r, int(code), names.get(code, ""), rec_vals, snt_vals)
        r += 2
    last_client_row = r - 1

    write_pair(r, "Grand Total", "", day_totals_r, day_totals_s, fill=total_fill, bold=True)
    ws.cell(row=r, column=1).alignment = center
    grand_row = r

    # Conditional highlights (same rules as the reference sheet).
    legend_r = grand_row + 3
    mult_cell, min_cell = f"$C${legend_r + 5}", f"$C${legend_r + 6}"
    rng = f"{fd}3:{ld}{last_client_row}"
    asof_abs = f"${asof_val}$1"
    rules = [
        (f'AND({fd}$2<>"",ISNUMBER({fd}3),{fd}3>={min_cell},{fd}3>={mult_cell}*${drr_l}3)', "C6EFCE", "006100"),
        (f'AND({fd}$2<>"",{fd}$2<={asof_abs},ISODD(ROW()),{fd}3=0)', "FFEB9C", "9C5700"),
        (f'AND({fd}$2<>"",{fd}$2<={asof_abs},ISEVEN(ROW()),{fd}3=0)', "FFC7CE", "9C0006"),
    ]
    for formula, bg, fg in rules:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[formula], font=Font(color=fg),
            fill=PatternFill(start_color=bg, end_color=bg, fill_type="solid"),
        ))

    ws.cell(row=legend_r, column=1, value="Highlight legend & thresholds").font = Font(bold=True, color=navy)
    legend = [
        ("C6EFCE", "006100", "High volume day: value ≥ multiplier × that row's Daily Run Rate AND ≥ minimum count"),
        ("FFEB9C", "9C5700", '0 Received cases on a date up to "Data as of"'),
        ("FFC7CE", "9C0006", '0 Sent cases on a date up to "Data as of"'),
    ]
    for i, (bg, fg, text) in enumerate(legend, start=1):
        s = ws.cell(row=legend_r + i, column=1, value="Sample")
        s.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        s.font = Font(bold=True, color=fg)
        s.alignment = center
        ws.cell(row=legend_r + i, column=2, value=text)
    for row, label, value in ((legend_r + 5, "High-volume multiplier (× DRR):", 3),
                              (legend_r + 6, "High-volume minimum count:", 10)):
        ws.cell(row=row, column=1, value=label)
        c = ws.cell(row=row, column=3, value=value)
        c.font = Font(bold=True, color="0000FF")
        c.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        c.alignment = center
    ws.cell(row=legend_r + 8, column=1, value=(
        f"Note: Rolling {days}-day window ending today ({dates[0]:%d-%b-%Y} to {end_date:%d-%b-%Y}); "
        "regenerated on every export. Received uses received_date, Sent uses Sent Date. "
        "Daily Run Rate = Grand Total ÷ days in window. Change the yellow cells to tune the high-volume flag."
    )).font = Font(size=9, color=grey_text)

    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 10
    for col in range(first_day_col, last_day_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 7.5
    ws.column_dimensions[asof_val].width = 10
    ws.column_dimensions[tl].width = 11
    ws.column_dimensions[drr_l].width = 14
    ws.column_dimensions[get_column_letter(gap_col)].width = 13
    ws.freeze_panes = f"{fd}3"

    wb.save(file_path)
    return True


def build_combined_report(incoming_subset, sent_subset):
    """Outer-join Incoming Cases and Sent Cases by Client Code + date, so every
    date in the selected range is kept on both sides (a client with cases sent
    on a day with no new incoming cases still gets a row, and vice versa)."""
    base = incoming_subset.rename(columns={
        "Client_Id": "Client Code",
        "Client": "Client_name",
        "Case Count": "Received_count",
    })

    sent_join = sent_subset.rename(columns={
        "sent_date": "Sent Date",
        "case_count": "Sent Case",
    })

    combined = base.merge(
        sent_join,
        how="outer",
        left_on=["Client Code", "received_date"],
        right_on=["Client Code", "Sent Date"],
        suffixes=("", "_sent"),
    )

    combined["Client_name"] = combined["Client_name"].fillna(combined["Client_name_sent"])
    combined["Received_count"] = combined["Received_count"].fillna(0).astype(int)
    combined["Sent Case"] = combined["Sent Case"].fillna(0).astype(int)

    combined = combined.sort_values(
        ["Client Code", "received_date", "Sent Date"], na_position="last"
    ).reset_index(drop=True)

    return combined[["Client Code", "Client_name", "received_date", "Sent Date", "Received_count", "Sent Case"]]


def build_monthly_pivot(incoming_df, sent_df, bm_lead=None):
    """One row per client, with a Received-/Sent- column pair per calendar month
    covered by incoming_df/sent_df, ordered chronologically (e.g. Received- Sep'26,
    Sent- Sep'26, Received- Oct'26, Sent- Oct'26, ...). bm_lead=None means no
    BM Team Lead filter (used for the consolidated master report)."""
    if bm_lead is None:
        inc = incoming_df.copy()
        sent = sent_df.copy()
    else:
        inc = incoming_df[incoming_df["BM Team Lead"] == bm_lead].copy()
        sent = sent_df[sent_df["BM Team Lead"] == bm_lead].copy()

    if inc.empty and sent.empty:
        return pd.DataFrame(columns=["Client Code", "Client_name"])

    inc["month_key"] = inc["received_date"].apply(lambda d: d.replace(day=1))
    sent["month_key"] = sent["sent_date"].apply(lambda d: d.replace(day=1))

    inc_agg = inc.groupby(["Client_Id", "Client", "month_key"], as_index=False)["Case Count"].sum()
    inc_agg = inc_agg.rename(columns={
        "Client_Id": "Client Code", "Client": "Client_name", "Case Count": "Received_count",
    })

    sent_agg = sent.groupby(["Client Code", "Client_name", "month_key"], as_index=False)["case_count"].sum()
    sent_agg = sent_agg.rename(columns={"case_count": "Sent_Case"})

    merged = inc_agg.merge(
        sent_agg, how="outer", on=["Client Code", "month_key"], suffixes=("", "_sent")
    )
    merged["Client_name"] = merged["Client_name"].fillna(merged.get("Client_name_sent"))
    merged["Received_count"] = merged["Received_count"].fillna(0).astype(int)
    merged["Sent_Case"] = merged["Sent_Case"].fillna(0).astype(int)

    months = sorted(merged["month_key"].dropna().unique())
    client_info = merged[["Client Code", "Client_name"]].drop_duplicates(subset=["Client Code"]).reset_index(drop=True)

    wide = client_info
    for month in months:
        label = pd.Timestamp(month).strftime("%b'%y")
        month_rows = merged.loc[merged["month_key"] == month, ["Client Code", "Received_count", "Sent_Case"]]
        month_rows = month_rows.rename(columns={
            "Received_count": f"Received- {label}",
            "Sent_Case": f"Sent- {label}",
        })
        wide = wide.merge(month_rows, on="Client Code", how="left")

    value_cols = [c for c in wide.columns if c.startswith("Received-") or c.startswith("Sent-")]
    wide[value_cols] = wide[value_cols].fillna(0).astype(int)

    if value_cols:
        grand_total = {col: int(wide[col].sum()) for col in value_cols}
        grand_total["Client Code"] = ""
        grand_total["Client_name"] = "Grand Total"
        wide = pd.concat([wide, pd.DataFrame([grand_total])], ignore_index=True)

    return wide


def init_outlook():
    try:
        pythoncom.CoInitialize()
        return win32.Dispatch('outlook.application')
    except Exception as e:
        st.error(f"❌ Unable to start Outlook COM. Error: {e}")
        st.stop()


def send_mail(outlook, to_email, subject, body, attachment_path):
    mail = outlook.CreateItem(0)
    mail.To = to_email
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Send()


tab1, tab2 = st.tabs(["📁 Manual Upload & Split", "📊 Daily / Monthly BM Lead Report"])

# ---------------------------------------------------------------------------
# Tab 1: original manual upload -> split by column -> email flow (unchanged)
# ---------------------------------------------------------------------------
with tab1:
    st.markdown("""
    Upload your input Excel file and the distribution list. This app will:
    1. Split data by selected columns.
    2. Save files in respective folders.
    3. Match names and designations from distribution list.
    4. Send emails using Outlook.
    5. Let you download output files and updated distribution sheet.
    """)

    input_file = st.file_uploader("📄 Upload Input Excel File", type=["xlsx"], key="manual_input")
    distribution_file = st.file_uploader("📋 Upload Distribution List File", type=["xlsx"], key="manual_dist")

    columns_to_split = st.multiselect(
        "🧩 Select columns to split by",
        options=["BM Team Member", "BM Team Lead"]
    )

    if st.button("🚀 Process and Send Emails", key="manual_send"):
        if input_file and distribution_file and columns_to_split:

            with tempfile.TemporaryDirectory() as tmpdir:

                output_folder = os.path.join(tmpdir, "output_files")
                os.makedirs(output_folder, exist_ok=True)

                df = pd.read_excel(input_file)
                distribution_df = pd.read_excel(distribution_file)
                distribution_df['Sent_Flag'] = 'Not Sent'

                for column in columns_to_split:

                    if column in df.columns:

                        column_folder = os.path.join(output_folder, column)
                        os.makedirs(column_folder, exist_ok=True)

                        for value in df[column].dropna().unique():

                            filtered_df = df[df[column] == value]

                            safe_value = "".join(c for c in str(value) if c.isalnum() or c in (" ", ".", "_")).strip()

                            if safe_value == "":
                                safe_value = "Unknown"

                            filename = f"{safe_value}.xlsx"
                            file_path = os.path.join(column_folder, filename)

                            try:
                                filtered_df.to_excel(file_path, index=False, engine='openpyxl')
                                format_workbook(file_path)
                            except Exception as e:
                                st.error(f"Error saving file {filename}: {e}")

                outlook = init_outlook()

                for subfolder in os.listdir(output_folder):

                    subfolder_path = os.path.join(output_folder, subfolder)

                    if os.path.isdir(subfolder_path):

                        matched_designations = distribution_df[
                            distribution_df['Designation'].str.strip().str.lower() == subfolder.strip().lower()
                        ]

                        for file_name in os.listdir(subfolder_path):

                            file_path = os.path.join(subfolder_path, file_name)
                            name_only = os.path.splitext(file_name)[0]

                            matched_row = matched_designations[
                                matched_designations['Name'].str.strip().str.lower() == name_only.strip().lower()
                            ]

                            if not matched_row.empty:

                                email_id = matched_row['Email_ID'].values[0]

                                try:
                                    send_mail(
                                        outlook, email_id,
                                        f"Attached: {name_only} Data ({subfolder})",
                                        f"Dear {name_only},\n\nPlease find the attached file.\n\nBest regards,\nYour Name",
                                        file_path,
                                    )
                                    distribution_df.loc[matched_row.index, 'Sent_Flag'] = 'Sent'
                                    st.success(f"✅ Email sent to {email_id} with {file_name}")

                                except Exception as e:
                                    st.error(f"❌ Failed to send to {email_id}: {e}")
                                    distribution_df.loc[matched_row.index, 'Sent_Flag'] = 'Failed'

                dist_with_flags = os.path.join(tmpdir, "Distribution_list_with_flags.xlsx")
                distribution_df.to_excel(dist_with_flags, index=False, engine='openpyxl')

                with open(dist_with_flags, 'rb') as f:
                    st.download_button(
                        "⬇️ Download Updated Distribution List",
                        f.read(),
                        file_name="Distribution_list_with_flags.xlsx",
                        key="manual_dl_dist",
                    )

                zip_path = shutil.make_archive(os.path.join(tmpdir, "output_files"), 'zip', output_folder)

                with open(zip_path, 'rb') as f:
                    st.download_button(
                        "⬇️ Download All Output Files as ZIP",
                        f.read(),
                        file_name="output_files.zip",
                        key="manual_dl_zip",
                    )

        else:
            st.warning("⚠️ Please upload both Excel files and select at least one column to split by.")

# ---------------------------------------------------------------------------
# Tab 2: DB-driven report, bifurcated by BM Team Lead. One workbook per BM
# Team Lead with two sheets: a dynamic Daily (month-to-date) breakdown and a
# user-ranged Monthly pivot. Export and Send Mail are separate steps so the
# generated files can be reviewed before anything goes out.
# ---------------------------------------------------------------------------
with tab2:
    st.markdown("""
    Pulls the **Incoming Case Count** and **Sent Case** figures straight from the DB,
    bifurcates them by **BM Team Lead** (using the Client → BM Team Lead mapping
    published in Google Sheets), and builds one workbook per BM Team Lead (plus a
    consolidated Master for all BM Team Leads) with three tabs:
    - **Summary** — client-wise Received vs Sent, always the rolling last 30 days ending today
    - **Daily Incoming vs sent** — row-level daily data for the date range you choose below
    - **Monthly Incoming vs sent** — a month-wise pivot for the same date range
    """)

    today = datetime.date.today()
    roll_start, roll_end = today - datetime.timedelta(days=ROLLING_DAYS - 1), today
    st.info(f"📅 Summary tab: **{roll_start:%d-%b-%Y}** to **{roll_end:%d-%b-%Y}** (rolling {ROLLING_DAYS} days, recalculated on every run)")

    st.markdown("**Date range** (applies to both the Daily and Monthly tabs; defaults to month-to-date):")
    col1, col2 = st.columns(2)
    with col1:
        range_start = st.date_input("Start Date", value=today.replace(day=1), key="range_start")
    with col2:
        range_end = st.date_input("End Date", value=today, key="range_end")

    if range_start > range_end:
        st.error("Start Date must be on or before End Date.")
        st.stop()

    distribution_file_db = st.file_uploader(
        "📋 Upload Distribution List (Designation must be 'BM Team Lead')",
        type=["xlsx"], key="db_dist"
    )

    if st.button("📤 Export Reports", key="db_export"):
        if distribution_file_db is None:
            st.warning("⚠️ Please upload the distribution list.")
            st.stop()

        with st.spinner("Querying database..."):
            try:
                incoming_range, sent_range, incoming_roll, sent_roll = fetch_report_data(
                    range_start, range_end, roll_start, roll_end
                )
            except Exception as e:
                st.error(f"❌ Database query failed: {e}")
                st.stop()

        try:
            mapping = cached_bm_mapping()
        except Exception as e:
            st.error(f"❌ Could not load BM Team Lead mapping from the Google Sheet: {e}")
            st.stop()

        incoming_roll = attach_bm_lead(incoming_roll, "Client_Id", mapping)
        sent_roll = attach_bm_lead(sent_roll, "Client Code", mapping)
        incoming_range = attach_bm_lead(incoming_range, "Client_Id", mapping)
        sent_range = attach_bm_lead(sent_range, "Client Code", mapping)

        unmapped_count = sum(
            (df["BM Team Lead"] == "Unmapped").sum()
            for df in (incoming_roll, sent_roll, incoming_range, sent_range)
        )

        distribution_df = pd.read_excel(distribution_file_db)
        distribution_df['Sent_Flag'] = 'Not Sent'

        bm_leads = sorted(
            set().union(*(set(df["BM Team Lead"]) for df in (
                incoming_roll, sent_roll, incoming_range, sent_range)))
        )
        bm_leads = [b for b in bm_leads if b != "Unmapped"]

        old_dir = st.session_state.get("tab2_output_dir")
        if old_dir and os.path.isdir(old_dir):
            shutil.rmtree(old_dir, ignore_errors=True)

        output_dir = tempfile.mkdtemp(prefix="bm_lead_reports_")
        files = {}
        unmatched_leads = []

        for bm_lead in bm_leads:
            daily_incoming_subset = incoming_range[incoming_range["BM Team Lead"] == bm_lead][
                ["Client_Id", "Client", "received_date", "Case Count"]
            ]
            daily_sent_subset = sent_range[sent_range["BM Team Lead"] == bm_lead][
                ["Client Code", "Client_name", "sent_date", "case_count"]
            ]
            daily_combined = build_combined_report(daily_incoming_subset, daily_sent_subset)
            monthly_pivot = build_monthly_pivot(incoming_range, sent_range, bm_lead)
            roll_inc_lead = incoming_roll[incoming_roll["BM Team Lead"] == bm_lead]
            roll_sent_lead = sent_roll[sent_roll["BM Team Lead"] == bm_lead]

            if daily_combined.empty and monthly_pivot.empty and roll_inc_lead.empty and roll_sent_lead.empty:
                continue

            safe_name = "".join(c for c in bm_lead if c.isalnum() or c in (" ", ".", "_")).strip() or "Unknown"
            file_path = os.path.join(output_dir, f"{safe_name}.xlsx")

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                daily_combined.to_excel(writer, sheet_name="Daily Incoming vs sent", index=False)
                monthly_pivot.to_excel(writer, sheet_name="Monthly Incoming vs sent", index=False)

            format_workbook(file_path)
            if not monthly_pivot.empty:
                bold_last_row(file_path, "Monthly Incoming vs sent")
            add_summary_sheet(file_path, roll_inc_lead, roll_sent_lead, roll_end)
            files[bm_lead] = file_path

            matched_row = distribution_df[
                (distribution_df['Designation'].str.strip().str.lower() == "bm team lead")
                & (distribution_df['Name'].str.strip().str.lower() == bm_lead.strip().lower())
            ]
            if matched_row.empty:
                unmatched_leads.append(bm_lead)

        master_daily = build_combined_report(
            incoming_range[["Client_Id", "Client", "received_date", "Case Count"]],
            sent_range[["Client Code", "Client_name", "sent_date", "case_count"]],
        )
        master_monthly = build_monthly_pivot(incoming_range, sent_range, bm_lead=None)

        master_file_path = os.path.join(output_dir, "Master_Consolidated.xlsx")
        with pd.ExcelWriter(master_file_path, engine='openpyxl') as writer:
            master_daily.to_excel(writer, sheet_name="Daily Incoming vs sent", index=False)
            master_monthly.to_excel(writer, sheet_name="Monthly Incoming vs sent", index=False)
        format_workbook(master_file_path)
        if not master_monthly.empty:
            bold_last_row(master_file_path, "Monthly Incoming vs sent")
        add_summary_sheet(master_file_path, incoming_roll, sent_roll, roll_end)

        zip_path = shutil.make_archive(os.path.join(output_dir, "bm_lead_reports"), 'zip', output_dir)

        st.session_state["tab2_output_dir"] = output_dir
        st.session_state["tab2_files"] = files
        st.session_state["tab2_zip_path"] = zip_path
        st.session_state["tab2_master_file"] = master_file_path
        st.session_state["tab2_distribution_df"] = distribution_df
        st.session_state["tab2_unmatched_leads"] = unmatched_leads
        st.session_state["tab2_unmapped_count"] = unmapped_count
        st.session_state["tab2_meta"] = {
            "roll_start": roll_start, "roll_end": roll_end,
            "range_start": range_start, "range_end": range_end,
        }
        st.session_state["tab2_dist_out_path"] = None

        st.success(f"✅ Exported {len(files)} BM Team Lead workbook(s). Review below, then click 'Send Mail' when ready.")

    # ---- Persistent review + send section (survives reruns via session_state) ----
    if st.session_state.get("tab2_files"):
        if st.session_state.get("tab2_unmapped_count"):
            st.warning(
                f"⚠️ {st.session_state['tab2_unmapped_count']} row(s) across both queries could not be "
                "matched to a BM Team Lead and were excluded (grouped under 'Unmapped')."
            )
        for bm_lead in st.session_state.get("tab2_unmatched_leads", []):
            st.warning(f"⚠️ No distribution list entry found for BM Team Lead '{bm_lead}' — file generated but will not be emailed.")

        if st.session_state.get("tab2_master_file"):
            master_path = st.session_state["tab2_master_file"]
            st.markdown("### 🗂️ Consolidated Master Report (all clients, no BM Team Lead split)")
            with st.expander("View Master_Consolidated.xlsx", expanded=False):
                st.write(f"**Summary** (rolling {ROLLING_DAYS} days)")
                st.dataframe(pd.read_excel(master_path, sheet_name="Summary", header=1), use_container_width=True)
                st.write("**Daily Incoming vs sent**")
                st.dataframe(pd.read_excel(master_path, sheet_name="Daily Incoming vs sent"), use_container_width=True)
                st.write("**Monthly Incoming vs sent**")
                st.dataframe(pd.read_excel(master_path, sheet_name="Monthly Incoming vs sent"), use_container_width=True)
            with open(master_path, 'rb') as f:
                st.download_button(
                    "⬇️ Download Master_Consolidated.xlsx",
                    f.read(), file_name="Master_Consolidated.xlsx", key="tab2_dl_master",
                )

        st.markdown("### 📂 Exported Reports (review before sending)")
        for bm_lead, file_path in st.session_state["tab2_files"].items():
            with st.expander(f"{bm_lead} — {os.path.basename(file_path)}"):
                if "Summary" in pd.ExcelFile(file_path).sheet_names:
                    st.write(f"**Summary** (rolling {ROLLING_DAYS} days)")
                    st.dataframe(pd.read_excel(file_path, sheet_name="Summary", header=1), use_container_width=True)
                st.write("**Daily Incoming vs sent**")
                st.dataframe(pd.read_excel(file_path, sheet_name="Daily Incoming vs sent"), use_container_width=True)
                st.write("**Monthly Incoming vs sent**")
                st.dataframe(pd.read_excel(file_path, sheet_name="Monthly Incoming vs sent"), use_container_width=True)
                with open(file_path, 'rb') as f:
                    st.download_button(
                        f"⬇️ Download {os.path.basename(file_path)}",
                        f.read(), file_name=os.path.basename(file_path), key=f"tab2_dl_{bm_lead}",
                    )

        with open(st.session_state["tab2_zip_path"], 'rb') as f:
            st.download_button(
                "⬇️ Download All Reports as ZIP (per-lead + master)",
                f.read(), file_name="bm_lead_reports.zip", key="db_dl_zip",
            )

        st.markdown("---")
        if st.button("📧 Send Mail", key="db_send_mail"):
            distribution_df = st.session_state["tab2_distribution_df"]
            meta = st.session_state["tab2_meta"]

            outlook = init_outlook()

            for bm_lead, file_path in st.session_state["tab2_files"].items():
                matched_row = distribution_df[
                    (distribution_df['Designation'].str.strip().str.lower() == "bm team lead")
                    & (distribution_df['Name'].str.strip().str.lower() == bm_lead.strip().lower())
                ]
                if matched_row.empty:
                    continue

                email_id = matched_row['Email_ID'].values[0]
                subject = f"Incoming vs Sent Report - {bm_lead}"
                body = (
                    f"Dear {bm_lead},\n\n"
                    f"Please find attached your Incoming vs Sent case report:\n"
                    f"- Summary: {meta['roll_start']:%d-%b-%Y} to {meta['roll_end']:%d-%b-%Y} (rolling {ROLLING_DAYS} days)\n"
                    f"- Daily and Monthly Incoming vs sent: {meta['range_start']:%d-%b-%Y} to {meta['range_end']:%d-%b-%Y}\n\n"
                    f"Best regards,\nCase Reporting"
                )

                try:
                    send_mail(outlook, email_id, subject, body, file_path)
                    distribution_df.loc[matched_row.index, 'Sent_Flag'] = 'Sent'
                    st.success(f"✅ Report emailed to {bm_lead} ({email_id})")
                except Exception as e:
                    st.error(f"❌ Failed to send to {email_id}: {e}")
                    distribution_df.loc[matched_row.index, 'Sent_Flag'] = 'Failed'

            st.session_state["tab2_distribution_df"] = distribution_df

            dist_out_path = os.path.join(st.session_state["tab2_output_dir"], "Distribution_list_with_flags.xlsx")
            distribution_df.to_excel(dist_out_path, index=False, engine='openpyxl')
            st.session_state["tab2_dist_out_path"] = dist_out_path

        if st.session_state.get("tab2_dist_out_path"):
            with open(st.session_state["tab2_dist_out_path"], 'rb') as f:
                st.download_button(
                    "⬇️ Download Updated Distribution List",
                    f.read(), file_name="Distribution_list_with_flags.xlsx", key="db_dl_dist",
                )
