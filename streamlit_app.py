import streamlit as st
import tempfile

from datetime import datetime
from zoneinfo import ZoneInfo

from excel_handler import ExcelHandler
from session_manager import initialize
from ui_helpers import format_value, highlight

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
# Initialize Session
# -------------------------------------------------------

initialize()

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
# Upload Excel
# -------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Monthly Tea Collection Excel File",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("Please upload your Excel file to begin.")
    st.stop()

# -------------------------------------------------------
# Create temp workbook only when a new file is uploaded
# -------------------------------------------------------

if (
    st.session_state["temp_file"] is None
    or st.session_state["uploaded_filename"] != uploaded_file.name
):

    temp = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    )

    temp.write(uploaded_file.getvalue())
    temp.close()

    st.session_state["temp_file"] = temp.name
    st.session_state["uploaded_filename"] = uploaded_file.name

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
# Current Sri Lanka Date
# -------------------------------------------------------

current_date = datetime.now(
    ZoneInfo("Asia/Colombo")
)

current_year = current_date.year
current_month = current_date.strftime("%B")


# -------------------------------------------------------
# Year / Month / Day
# -------------------------------------------------------

date_col1, date_col2, date_col3 = st.columns(3)

with date_col1:

    selected_year = st.number_input(
        "Year",
        value=current_year,
        disabled=True
    )


with date_col2:

    selected_month = st.text_input(
        "Month",
        value=current_month,
        disabled=True
    )


with date_col3:

    selected_day = st.selectbox(
        "Day",
        list(range(1, 32))
    )

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

        row, column = handler.update(
            customer_code,
            selected_day,
            amount
        )

        handler.save(
            st.session_state["temp_file"]
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

        import time

        time.sleep(2)

        success.empty()

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


