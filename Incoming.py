import streamlit as st
import pandas as pd
import os
import tempfile
import shutil
import datetime
import win32com.client as win32
import pythoncom

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from bm_mapping import load_bm_mapping, attach_bm_lead
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


def build_monthly_pivot(incoming_df, sent_df, bm_lead):
    """One row per client, with a Received-/Sent- column pair per calendar month
    covered by incoming_df/sent_df, ordered chronologically (e.g. Received- Sep'26,
    Sent- Sep'26, Received- Oct'26, Sent- Oct'26, ...)."""
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
    published in Google Sheets), and builds one workbook per BM Team Lead with two tabs:
    - **Daily Incoming vs sent** — always 1st of the current month through today (dynamic, no filter)
    - **Monthly Incoming vs sent** — a month-wise pivot for the date range you choose below
    """)

    today = datetime.date.today()
    mtd_start, mtd_end = today.replace(day=1), today
    st.info(f"📅 Daily tab: **{mtd_start:%d-%b-%Y}** to **{mtd_end:%d-%b-%Y}** (month-to-date, recalculated on every run)")

    def _shift_months(d, delta):
        m = d.month - 1 + delta
        y = d.year + m // 12
        return datetime.date(y, m % 12 + 1, 1)

    st.markdown("**Monthly tab date range** (can span multiple months):")
    col1, col2 = st.columns(2)
    with col1:
        monthly_start = st.date_input("Start Date", value=_shift_months(today, -2), key="monthly_start")
    with col2:
        monthly_end = st.date_input("End Date", value=today, key="monthly_end")

    if monthly_start > monthly_end:
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
                incoming_mtd = fetch_incoming_case_count(mtd_start, mtd_end)
                sent_mtd = fetch_sent_case(mtd_start, mtd_end)
                incoming_range = fetch_incoming_case_count(monthly_start, monthly_end)
                sent_range = fetch_sent_case(monthly_start, monthly_end)
            except Exception as e:
                st.error(f"❌ Database query failed: {e}")
                st.stop()

        try:
            mapping = load_bm_mapping()
        except Exception as e:
            st.error(f"❌ Could not load BM Team Lead mapping from the Google Sheet: {e}")
            st.stop()

        incoming_mtd = attach_bm_lead(incoming_mtd, "Client_Id", mapping)
        sent_mtd = attach_bm_lead(sent_mtd, "Client Code", mapping)
        incoming_range = attach_bm_lead(incoming_range, "Client_Id", mapping)
        sent_range = attach_bm_lead(sent_range, "Client Code", mapping)

        unmapped_count = sum(
            (df["BM Team Lead"] == "Unmapped").sum()
            for df in (incoming_mtd, sent_mtd, incoming_range, sent_range)
        )

        distribution_df = pd.read_excel(distribution_file_db)
        distribution_df['Sent_Flag'] = 'Not Sent'

        bm_leads = sorted(
            set(incoming_mtd["BM Team Lead"]) | set(sent_mtd["BM Team Lead"])
            | set(incoming_range["BM Team Lead"]) | set(sent_range["BM Team Lead"])
        )
        bm_leads = [b for b in bm_leads if b != "Unmapped"]

        old_dir = st.session_state.get("tab2_output_dir")
        if old_dir and os.path.isdir(old_dir):
            shutil.rmtree(old_dir, ignore_errors=True)

        output_dir = tempfile.mkdtemp(prefix="bm_lead_reports_")
        files = {}
        unmatched_leads = []

        for bm_lead in bm_leads:
            daily_incoming_subset = incoming_mtd[incoming_mtd["BM Team Lead"] == bm_lead][
                ["Client_Id", "Client", "received_date", "Case Count"]
            ]
            daily_sent_subset = sent_mtd[sent_mtd["BM Team Lead"] == bm_lead][
                ["Client Code", "Client_name", "sent_date", "case_count"]
            ]
            daily_combined = build_combined_report(daily_incoming_subset, daily_sent_subset)
            monthly_pivot = build_monthly_pivot(incoming_range, sent_range, bm_lead)

            if daily_combined.empty and monthly_pivot.empty:
                continue

            safe_name = "".join(c for c in bm_lead if c.isalnum() or c in (" ", ".", "_")).strip() or "Unknown"
            file_path = os.path.join(output_dir, f"{safe_name}.xlsx")

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                daily_combined.to_excel(writer, sheet_name="Daily Incoming vs sent", index=False)
                monthly_pivot.to_excel(writer, sheet_name="Monthly Incoming vs sent", index=False)

            format_workbook(file_path)
            if not monthly_pivot.empty:
                bold_last_row(file_path, "Monthly Incoming vs sent")
            files[bm_lead] = file_path

            matched_row = distribution_df[
                (distribution_df['Designation'].str.strip().str.lower() == "bm team lead")
                & (distribution_df['Name'].str.strip().str.lower() == bm_lead.strip().lower())
            ]
            if matched_row.empty:
                unmatched_leads.append(bm_lead)

        zip_path = shutil.make_archive(os.path.join(output_dir, "bm_lead_reports"), 'zip', output_dir)

        st.session_state["tab2_output_dir"] = output_dir
        st.session_state["tab2_files"] = files
        st.session_state["tab2_zip_path"] = zip_path
        st.session_state["tab2_distribution_df"] = distribution_df
        st.session_state["tab2_unmatched_leads"] = unmatched_leads
        st.session_state["tab2_unmapped_count"] = unmapped_count
        st.session_state["tab2_meta"] = {
            "mtd_start": mtd_start, "mtd_end": mtd_end,
            "monthly_start": monthly_start, "monthly_end": monthly_end,
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

        st.markdown("### 📂 Exported Reports (review before sending)")
        for bm_lead, file_path in st.session_state["tab2_files"].items():
            with st.expander(f"{bm_lead} — {os.path.basename(file_path)}"):
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
                "⬇️ Download All BM Lead Reports as ZIP",
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
                    f"- Daily Incoming vs sent: {meta['mtd_start']:%d-%b-%Y} to {meta['mtd_end']:%d-%b-%Y} (month-to-date)\n"
                    f"- Monthly Incoming vs sent: {meta['monthly_start']:%d-%b-%Y} to {meta['monthly_end']:%d-%b-%Y}\n\n"
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
