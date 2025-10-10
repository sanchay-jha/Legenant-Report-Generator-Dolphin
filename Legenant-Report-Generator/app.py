import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import io

# ----------------------------
# 🔐 Login Configuration
# ----------------------------
USERNAME = "dolphin"
PASSWORD = "Outsourcinghubindia@2025"

# Initialize login session state
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# Login function
def login():
    st.set_page_config(page_title="Login", layout="centered")
    st.title("🔐 Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == USERNAME and password == PASSWORD:
            st.session_state.logged_in = True
            st.success("Login successful! ✅")
            st.experimental_rerun()
        else:
            st.error("Invalid username or password ❌")

# Show login page if not logged in
if not st.session_state.logged_in:
    login()
    st.stop()

# ----------------------------
# ✅ Main App Starts Here (Only after login)
# ----------------------------
st.set_page_config(page_title="Custom Report Processor", layout="wide")
st.title("📑 Custom Report Processor")

uploaded_file = st.file_uploader("Upload your Excel file (Custom Report.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()

        # Two workbook objects:
        wb = load_workbook(io.BytesIO(file_bytes))
        wb_data = load_workbook(io.BytesIO(file_bytes), data_only=True)

        AR_Aging_sheet = wb["AR Aging (excluding HUD)"]
        Sample_sheet = wb["Sample Report"]
        Rent_sheet = wb["Rent Roll w. Lease Charges"]
        Legal_sheet = wb["Legal Report"]
        Tenant_memo_sheet = wb["Tenant Memo's"]

        AR_Aging_sheet_data = wb_data["AR Aging (excluding HUD)"]
        Rent_sheet_data = wb_data["Rent Roll w. Lease Charges"]
        Legal_sheet_data = wb_data["Legal Report"]
        Tenant_memo_sheet_data = wb_data["Tenant Memo's"]
        Sample_sheet_data = wb_data["Sample Report"]

        ws = Sample_sheet
        try:
            ws.font = Font(bold=True)
            ws.font = Font(size=8)
        except Exception:
            pass

        ws['J2'] = ("Unit")
        ws['H2'] = ("Name")
        ws['S2'] = ("Unit")

        sheet = Legal_sheet
        unit_legal = []
        blank_count = 0

        for row in sheet.iter_rows(min_row=7, min_col=2, max_col=2):
            cell = row[0]
            value = cell.value

            if value is None or str(value).strip() == "":
                blank_count += 1
                if blank_count > 1:
                    break
                continue

            blank_count = 0

            if cell.font and not cell.font.bold:
                unit_legal.append(value)

        ws = Sample_sheet
        for col in range(0, len(unit_legal)):
            char = chr(65)
            ws[f"{char}{col + 3}"] = AR_Aging_sheet["A7"].value

        unit_rent = []

        def get_unit_rent_list(wb_obj, sheet_name):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"A{row}"].value
                if cell_value is None:
                    continue
                cell_str = str(cell_value).strip().lower()
                if "summary" in cell_str:
                    break
                unit_rent.append(cell_value)
            return unit_rent

        unit_rent = get_unit_rent_list(wb, "Rent Roll w. Lease Charges")

        ws = Sample_sheet
        for col in range(0, len(unit_legal)):
            char = chr(83)
            ws[f"{char}{col + 3}"] = unit_legal[col]

        unit_ar = []
        unit_ar_cell = []

        def process_until_total_A(wb_obj, sheet_name, column='A'):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"{column}{row}"].value
                cell_number = ws_local[f"{column}{row}"]
                if cell_value is not None and str(cell_value).strip().lower() == 'total':
                    break
                unit_ar.append(cell_value)
                unit_ar_cell.append(cell_number.coordinate)

        process_until_total_A(wb, "AR Aging (excluding HUD)", column='A')

        resident_ar = []

        def process_until_none_B(wb_obj, sheet_name, column='B'):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"{column}{row}"].value
                if cell_value is None:
                    break
                resident_ar.append(cell_value)

        process_until_none_B(wb, "AR Aging (excluding HUD)", column='B')

        status_ar = []

        def process_until_none_C(wb_obj, sheet_name, column='C'):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"{column}{row}"].value
                if cell_value is None:
                    break
                status_ar.append(cell_value)

        process_until_none_C(wb, "AR Aging (excluding HUD)", column='C')

        tenant_name_ar = []

        def process_until_none_D(wb_obj, sheet_name, column='D'):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"{column}{row}"].value
                if cell_value is None:
                    break
                tenant_name_ar.append(cell_value)

        process_until_none_D(wb, "AR Aging (excluding HUD)", column='D')

        _0_30_ar = []

        def process_until_none_F(wb_obj, sheet_name, column='F'):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"{column}{row}"].value
                if cell_value is None:
                    break
                _0_30_ar.append(cell_value)

        process_until_none_F(wb, "AR Aging (excluding HUD)", column='F')

        total_charges_ar = []

        def process_until_none_E(wb_obj, sheet_name, column='E'):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"{column}{row}"].value
                if cell_value is None:
                    break
                total_charges_ar.append(cell_value)

        process_until_none_E(wb, "AR Aging (excluding HUD)", column='E')

        ws = Sample_sheet
        row_start = 3
        current_row = row_start

        ar_data = {}
        for u, resident, status, tenant_name, total_charges, ar_0_30 in zip(
                unit_ar, resident_ar, status_ar, tenant_name_ar, total_charges_ar, _0_30_ar):
            key = str(u).strip().upper() if u is not None else ""
            ar_data[key] = {
                "resident": resident,
                "status": status,
                "tenant_name": tenant_name,
                "total_charges": total_charges,
                "ar_0_30": ar_0_30
            }

        seen_units = set()
        ar_amount_written_units = set()

        for unit in unit_legal:
            key = str(unit).strip().upper()
            if key not in seen_units:
                ws[f"J{current_row}"] = key
                seen_units.add(key)
            else:
                ws[f"J{current_row}"] = ""

            if key in ar_data:
                ws[f"B{current_row}"] = key
                ws[f"C{current_row}"] = ar_data[key]["resident"]
                ws[f"D{current_row}"] = ar_data[key]["status"]
                ws[f"E{current_row}"] = ar_data[key]["tenant_name"]

                if key not in ar_amount_written_units:
                    ws[f"F{current_row}"] = ar_data[key]["ar_0_30"]
                    ws[f"G{current_row}"] = ar_data[key]["total_charges"]
                    ar_amount_written_units.add(key)

            current_row += 1

        def fetch_rent_amounts(wb_obj, sheet_name):
            ws_local = wb_obj[sheet_name]
            unit_rent_amount_local = []
            for row in range(2, ws_local.max_row + 1):
                g_value = ws_local[f"G{row}"].value
                h_value = ws_local[f"H{row}"].value
                g_text = str(g_value).strip().lower() if g_value is not None else ""
                if "market" in g_text:
                    break
                if "rent" in g_text and h_value is not None:
                    unit_rent_amount_local.append(h_value)
            return unit_rent_amount_local

        unit_rent_amount = fetch_rent_amounts(wb, "Rent Roll w. Lease Charges")

        for col in range(len(unit_rent_amount)):
            char = chr(84)
            ws[f"{char}{col + 3}"] = unit_rent_amount[col]

        st.success("✅ Processing complete. Download your updated report below:")

        # Convert the workbook to bytes
        output = io.BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Download Updated Excel File",
            data=output.getvalue(),
            file_name="Updated_Custom_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ An error occurred while processing the file: {e}")
