import streamlit as st
import tempfile
import time

import httpx
from supabase import create_client


BUCKET_NAME = "tea-workbooks"

# Number of attempts for temporary network/DNS problems
MAX_RETRIES = 4
RETRY_DELAY = 2


# -------------------------------------------------------
# Supabase Client
# -------------------------------------------------------

def get_supabase_client():

    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_KEY"]
    )


# -------------------------------------------------------
# Retry Supabase Operation
# -------------------------------------------------------

def run_with_retry(operation):

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            return operation()

        except (
            httpx.ConnectError,
            httpx.ConnectTimeout,
            httpx.ReadTimeout
        ) as e:

            last_error = e

            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY * attempt)

    # All attempts failed
    raise last_error


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

    workbook_path = get_workbook_path(
        year,
        month
    )

    def upload():

        # Create a fresh client for each retry
        supabase = get_supabase_client()

        return (
            supabase.storage
            .from_(BUCKET_NAME)
            .upload(
                path=workbook_path,
                file=file_bytes,
                file_options={
                    "content-type":
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "upsert": "true"
                }
            )
        )

    return run_with_retry(upload)


# -------------------------------------------------------
# Download Monthly Workbook
# -------------------------------------------------------

def download_workbook(
    year,
    month
):

    workbook_path = get_workbook_path(
        year,
        month
    )

    def download():

        # Create a fresh client for each retry
        supabase = get_supabase_client()

        return (
            supabase.storage
            .from_(BUCKET_NAME)
            .download(workbook_path)
        )

    try:

        file_bytes = run_with_retry(download)

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

    return upload_workbook(
        file_bytes,
        year,
        month
    )