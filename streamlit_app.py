import streamlit as st
import tempfile
import calendar

from datetime import datetime
from zoneinfo import ZoneInfo

from excel_handler import ExcelHandler
from session_manager import initialize
from ui_helpers import format_value, highlight

from auth import (
    login,
    logout
)

from audit import verify_audit_password

from database import save_collection

from workbook_storage import (
    upload_workbook,
    download_workbook,
    sync_workbook
)

from audit_db import (
    initialize_audit_db,
    add_audit_record,
    get_audit_records
)

DATE_EDIT_PASSWORD = "date123"
# -------------------------------------------------------
# Page Configuration
# -------------------------------------------------------

st.set_page_config(
    page_title="Tea Collection System",
    page_icon="🍃",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------------------------------------------
# Login Authentication
# -------------------------------------------------------

if not login():
    st.stop()

# -------------------------------------------------------
# Initialize Session
# -------------------------------------------------------

initialize()

# -------------------------------------------------------
# Initialize Audit Database
# -------------------------------------------------------

initialize_audit_db()

# -------------------------------------------------------
# Create Excel Handler
# -------------------------------------------------------

handler = ExcelHandler()

# -------------------------------------------------------
# Title
# -------------------------------------------------------

st.title("🍃 Tea Collection Entry System")
st.markdown("---")

# -------------------------------------------------------
# Current Sri Lanka Date
# -------------------------------------------------------

current_date = datetime.now(
    ZoneInfo("Asia/Colombo")
)

current_year = current_date.year
current_month = current_date.strftime("%B")

# -------------------------------------------------------
# Initialize Year / Month
# -------------------------------------------------------

if st.session_state["selected_year"] is None:
    st.session_state["selected_year"] = current_year

if st.session_state["selected_month"] is None:
    st.session_state["selected_month"] = current_month

# -------------------------------------------------------
# Upload / Restore Excel
# -------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload / Replace Monthly Tea Collection Excel File",
    type=["xlsx"],
    key="workbook_uploader"
)


# -------------------------------------------------------
# New Workbook Uploaded
# -------------------------------------------------------

if uploaded_file is not None:

    # Only process if this is a different upload
    if (
        st.session_state["uploaded_filename"]
        != uploaded_file.name
    ):

        file_bytes = uploaded_file.getvalue()

        # Save permanently to Supabase Storage
        upload_workbook(
            file_bytes,
            st.session_state["selected_year"],
            st.session_state["selected_month"]
        )

        # Create local working copy
        temp = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )

        temp.write(file_bytes)
        temp.close()

        st.session_state["temp_file"] = (
            temp.name
        )

        st.session_state[
            "uploaded_filename"
        ] = uploaded_file.name


# -------------------------------------------------------
# No Workbook in Current Session
# -------------------------------------------------------

if st.session_state["temp_file"] is None:

    # Try restoring last workbook from Supabase
    restored_file = download_workbook(
        st.session_state["selected_year"],
        st.session_state["selected_month"]
    )

    if restored_file is not None:

        st.session_state["temp_file"] = (
            restored_file
        )

        st.session_state[
            "uploaded_filename"
        ] = "Restored Workbook"

        st.success(
            "✅ Previous workbook restored."
        )

    else:

        st.info(
            "Please upload an Excel workbook to begin."
        )

        st.stop()

# -------------------------------------------------------
# Load workbook
# -------------------------------------------------------

workbook, sheet = handler.load(
    st.session_state["temp_file"]
)

df = handler.dataframe()

customer_df, customer_dict = handler.get_customer_data()

# -------------------------------------------------------
# Entry Form
# -------------------------------------------------------

st.markdown("---")
st.subheader("Tea Collection Entry")


# -------------------------------------------------------
# Year / Month Access Control
# -------------------------------------------------------

date_col1, date_col2, date_col3 = st.columns(3)

month_names = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
]


# -------------------------------------------------------
# Year
# -------------------------------------------------------

with date_col1:

    if st.session_state.get(
        "date_fields_unlocked",
        False
    ):

        selected_year = st.number_input(
            "Year",
            min_value=2020,
            max_value=2100,
            value=int(
                st.session_state["selected_year"]
            ),
            step=1
        )

        st.session_state["selected_year"] = (
            selected_year
        )

    else:

        selected_year = st.number_input(
            "Year",
            value=int(
                st.session_state["selected_year"]
            ),
            disabled=True
        )


# -------------------------------------------------------
# Month
# -------------------------------------------------------

with date_col2:

    if st.session_state.get(
        "date_fields_unlocked",
        False
    ):

        current_month_index = month_names.index(
            st.session_state["selected_month"]
        )

        selected_month = st.selectbox(
            "Month",
            month_names,
            index=current_month_index
        )

        st.session_state["selected_month"] = (
            selected_month
        )

    else:

        selected_month = st.text_input(
            "Month",
            value=st.session_state[
                "selected_month"
            ],
            disabled=True
        )


# -------------------------------------------------------
# Day
# -------------------------------------------------------


with date_col3:

    # Convert month name to month number
    month_number = datetime.strptime(
        selected_month,
        "%B"
    ).month

    # Get number of days in selected month
    days_in_month = calendar.monthrange(
        selected_year,
        month_number
    )[1]

    # Display only valid days
    selected_day = st.selectbox(
        "Day",
        list(range(1, days_in_month + 1))
    )

# -------------------------------------------------------
# Unlock Year / Month
# -------------------------------------------------------

if not st.session_state.get(
    "date_fields_unlocked",
    False
):

    if st.button(
        "🔐 Change Year / Month",
        key="change_year_month"
    ):

        st.session_state[
            "show_date_password"
        ] = True


# -------------------------------------------------------
# Password Verification
# -------------------------------------------------------

if (
    st.session_state.get(
        "show_date_password",
        False
    )
    and not st.session_state.get(
        "date_fields_unlocked",
        False
    )
):

    date_password = st.text_input(
        "Enter password to change Year / Month",
        type="password",
        key="date_edit_password"
    )

    if st.button(
        "Unlock",
        key="unlock_date_fields"
    ):

        if date_password == DATE_EDIT_PASSWORD:

            st.session_state[
                "date_fields_unlocked"
            ] = True

            st.session_state[
                "show_date_password"
            ] = False

            st.rerun()

        else:

            st.error(
                "Incorrect password."
            )


# -------------------------------------------------------
# Lock Year / Month
# -------------------------------------------------------

if st.session_state.get(
    "date_fields_unlocked",
    False
):

    if st.button(
        "🔒 Lock Year / Month",
        key="lock_date_fields"
    ):

        st.session_state[
            "date_fields_unlocked"
        ] = False

        st.session_state[
            "show_date_password"
        ] = False

        # Force app to load workbook
        # for newly selected Year / Month
        st.session_state[
            "temp_file"
        ] = None

        st.session_state[
            "uploaded_filename"
        ] = None

        st.rerun()

# Customer Row
col1, col2 = st.columns([1, 2])

with col1:
    customer_code = st.selectbox(
        "Customer Code",
        customer_df["Customer Code"]
    )

with col2:
    customer_name = customer_dict.get(customer_code, "")

    st.text_input(
        "Customer Name",
        value=customer_name,
        disabled=True
    )

# Amount
amount = st.number_input(
    "Amount (Kg)",
    min_value=0.0,
    step=0.5,
    format="%.2f"
)

# -------------------------------------------------------
# Submit Entry
# -------------------------------------------------------

if st.button(
    "Submit Entry",
    use_container_width=True
):

    try:

        old_value = sheet.cell(
            row=handler.find_customer_row(customer_code),
            column=handler.find_date_column(selected_day)
        ).value

        row, column = handler.update(
            customer_code,
            selected_day,
            amount
        )

        # -------------------------------------------------------
        # Save Collection Permanently
        # -------------------------------------------------------

        save_collection(
            customer_code=customer_code,
            customer_name=customer_name,
            year=selected_year,
            month=selected_month,
            day=selected_day,
            amount=amount
        )

        # -------------------------------------------------------
        # Save Audit Record to SQLite
        # -------------------------------------------------------

        add_audit_record(
            customer_code=customer_code,
            customer_name=customer_name,
            year=selected_year,
            month=selected_month,
            day=selected_day,
            old_amount=old_value,
            new_amount=amount,
            excel_row=row,
            excel_column=column
        )

        handler.save(
            st.session_state["temp_file"]
        )

        # -------------------------------------------------------
        # Backup Updated Excel to Supabase Storage
        # -------------------------------------------------------

        sync_workbook(
            st.session_state["temp_file"],
            selected_year,
            selected_month
        )

        st.session_state["updated_cells"].append(
            {
                "row": row,
                "column": column,
                "year": selected_year,
                "month": selected_month,
                "day": selected_day,
                "amount": amount,
                "time": datetime.now()
            }
        )

        success = st.empty()

        success.success(
            "✅ Amount Added Successfully"
        )

        st.rerun()

    except Exception as e:

        st.error(str(e))
# -------------------------------------------------------
# Current Excel Sheet
# -------------------------------------------------------

st.markdown("---")

st.markdown(
    f"""
    <h3>
        Current Excel Sheet → [ {selected_year} - {selected_month} ]
    </h3>
    """,
    unsafe_allow_html=True
)


styled_df = (
    df.style
    .format(
        format_value,
        subset=df.columns[2:]
    )
    .apply(
        lambda data:
        highlight(
            data,
            st.session_state["updated_cells"]
        ),
        axis=None
    )
)


st.dataframe(
    styled_df,
    use_container_width=True,
    height=500
)

# -------------------------------------------------------
# Navigation Drawer
# -------------------------------------------------------

with st.sidebar:

    st.write(
    f"👤 {st.session_state.get('username', '')}"
    )

    if st.button(
        "🚪 Logout",
        use_container_width=True,
        key="logout_button"
    ):

        logout()

        st.rerun()

    st.title("🍃 Menu")
    
    st.markdown("---")

    # ---------------------------------------------------
    # Download Workbook
    # ---------------------------------------------------

    with open(
        st.session_state["temp_file"],
        "rb"
    ) as file:

        st.download_button(
            label="📥 Download Workbook",

            data=file.read(),

            file_name=(
                f"Tea_Collection_"
                f"{selected_year}_"
                f"{selected_month}.xlsx"
            ),

            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True
        )

    st.markdown("---")

    # ---------------------------------------------------
    # Save Workbook
    # ---------------------------------------------------
    
    if st.button(
            "💾 Save Workbook",
            use_container_width=True
        ):
    
            temp_save = tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".xlsx"
            )
    
            temp_save.close()
    
            handler.save(
                temp_save.name
            )
    
            saved_item = {
                "name": (
                    f"Tea_Collection_"
                    f"{selected_year}_"
                    f"{selected_month}_"
                    f"{len(st.session_state['saved_workbooks']) + 1}"
                    ".xlsx"
                ),
    
                "path": temp_save.name,
    
                "time": datetime.now(
                    ZoneInfo("Asia/Colombo")
                )
            }
    
            st.session_state[
                "saved_workbooks"
            ].insert(
                0,
                saved_item
            )
    
            # Keep only last 10 files
            st.session_state[
                "saved_workbooks"
            ] = (
                st.session_state[
                    "saved_workbooks"
                ][:10]
            )
    
            st.success(
                "✅ Workbook Saved Successfully"
            )

    # ---------------------------------------------------
    # Saved Excel Sheets
    # ---------------------------------------------------

    if st.button(
        "📂 Saved Excel Sheets",
        use_container_width=True
    ):

        st.session_state["show_saved_panel"] = (
            not st.session_state["show_saved_panel"]
        )


    if st.session_state["show_saved_panel"]:

        st.subheader("Last 10 Saved Files")

        saved_files = (
            st.session_state["saved_workbooks"]
        )

        if not saved_files:

            st.info(
                "No saved files available."
            )

        else:

            selected_file = st.selectbox(
                "Select Workbook",
                saved_files,

                format_func=lambda item:
                    item["time"].strftime(
                        "%d-%m-%Y %I:%M:%S %p"
                    )
            )


            if st.button(
                "📖 Open Selected",
                use_container_width=True
            ):

                st.session_state[
                    "temp_file"
                ] = selected_file["path"]

                st.session_state[
                    "uploaded_filename"
                ] = selected_file["name"]

                st.success(
                    f"Opened {selected_file['name']}"
                )

                st.rerun()

    st.markdown("---")

    # ---------------------------------------------------
    # Audit Panel
    # ---------------------------------------------------

    if st.button(
        "🔐 Audit Panel",
        use_container_width=True,
        key="audit_panel_button"
    ):

        st.session_state["show_audit_panel"] = (
            not st.session_state.get(
                "show_audit_panel",
                False
            )
        )


    if st.session_state.get(
        "show_audit_panel",
        False
    ):

        if not st.session_state.get(
            "audit_authenticated",
            False
        ):

            st.subheader("🔐 Audit Access")

            audit_password = st.text_input(
                "Password",
                type="password",
                key="audit_password"
            )

            if st.button(
                "Unlock",
                use_container_width=True,
                key="unlock_audit_button"
            ):

                if verify_audit_password(
                    audit_password
                ):

                    st.session_state[
                        "audit_authenticated"
                    ] = True

                    st.rerun()

                else:

                    st.error(
                        "Incorrect password."
                    )

        else:

            st.subheader("📋 Audit History")

            audit_records = get_audit_records()

            if not audit_records:

                st.info(
                    "No audit records available."
                )

            else:

                for record in audit_records:

                    (
                        record_id,
                        customer_code,
                        customer_name,
                        year,
                        month,
                        day,
                        old_amount,
                        new_amount,
                        updated_at
                    ) = record

                    updated_time = datetime.fromisoformat(
                        updated_at
                    )

                    # Format old amount
                    if old_amount is None:
                        old_amount_display = "-"
                    else:
                        old_amount_display = f"{old_amount:g} Kg"

                    # Format new amount
                    new_amount_display = f"{new_amount:g} Kg"

                    st.markdown(
                        f"""
                        **Customer:**  
                        {customer_code} - {customer_name}

                        **Collection Date:**  
                        {year} - {month} - {day}

                        **Old Amount:**  
                        {old_amount_display}

                        **New Amount:**  
                        {new_amount_display}

                        **Updated:**  
                        {updated_time.strftime('%d-%m-%Y %I:%M:%S %p')}
                        """
                    )

                    st.markdown("---")

        if st.button(
                "🔒 Lock Audit Panel",
                use_container_width=True,
                key="lock_audit_button"
            ):

                st.session_state[
                    "audit_authenticated"
                ] = False

                st.session_state[
                    "show_audit_panel"
                ] = False

                st.rerun()            