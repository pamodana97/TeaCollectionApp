import streamlit as st
import tempfile

from supabase import create_client


BUCKET_NAME = "tea-workbooks"


# -------------------------------------------------------
# Supabase Client
# -------------------------------------------------------

def get_supabase_client():

    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_KEY"]
    )


# -------------------------------------------------------
# Build Workbook Path
# -------------------------------------------------------

def get_workbook_path(year, month):

    return f"{int(year)}/{month}.xlsx"


# -------------------------------------------------------
# Upload / Replace Monthly Workbook
# -------------------------------------------------------

def upload_workbook(
    file_bytes,
    year,
    month
):

    supabase = get_supabase_client()

    workbook_path = get_workbook_path(
        year,
        month
    )

    supabase.storage.from_(
        BUCKET_NAME
    ).upload(
        path=workbook_path,
        file=file_bytes,
        file_options={
            "content-type":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "upsert": "true"
        }
    )


# -------------------------------------------------------
# Download Monthly Workbook
# -------------------------------------------------------

def download_workbook(
    year,
    month
):

    supabase = get_supabase_client()

    workbook_path = get_workbook_path(
        year,
        month
    )

    try:

        file_bytes = (
            supabase.storage
            .from_(BUCKET_NAME)
            .download(workbook_path)
        )

        temp = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )

        temp.write(file_bytes)
        temp.close()

        return temp.name

    except Exception:

        return None


# -------------------------------------------------------
# Sync Updated Monthly Workbook
# -------------------------------------------------------

def sync_workbook(
    local_path,
    year,
    month
):

    with open(
        local_path,
        "rb"
    ) as file:

        file_bytes = file.read()

    upload_workbook(
        file_bytes,
        year,
        month
    )