import streamlit as st
import tempfile
import time
import httpx

from supabase import create_client
from storage3.exceptions import StorageApiError


BUCKET_NAME = "tea-workbooks"
MASTER_TEMPLATE_OBJECT = "master/New Worksheet.xlsx"

MAX_RETRIES = 4
RETRY_DELAY = 2
CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# -------------------------------------------------------
# Supabase Client
# -------------------------------------------------------

def get_supabase_client():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_KEY"]
    )


# -------------------------------------------------------
# Retry temporary network failures only
# -------------------------------------------------------

def run_with_retry(operation):
    last_error = None

    for attempt in range(MAX_RETRIES):
        try:
            return operation()
        except (
            httpx.ConnectError,
            httpx.ConnectTimeout,
            httpx.ReadTimeout,
            httpx.WriteTimeout
        ) as error:
            last_error = error
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY * (attempt + 1))

    raise last_error


# -------------------------------------------------------
# Storage helpers
# -------------------------------------------------------

def _is_object_missing(error):
    status = str(getattr(error, "status", "")).strip()
    message = str(error).lower()
    # Do not mistake a network error or a missing bucket for a missing file.
    return status == "404" and (
        "object not found" in message
        or "not_found" in message
    )


def _download_object(object_path):
    def operation():
        return (
            get_supabase_client()
            .storage.from_(BUCKET_NAME)
            .download(object_path)
        )

    try:
        return run_with_retry(operation)
    except StorageApiError as error:
        if _is_object_missing(error):
            return None
        raise


def _upload_object(object_path, file_bytes):
    def operation():
        return (
            get_supabase_client()
            .storage.from_(BUCKET_NAME)
            .upload(
                path=object_path,
                file=file_bytes,
                file_options={
                    "content-type": CONTENT_TYPE,
                    "upsert": "true"
                }
            )
        )

    return run_with_retry(operation)


# -------------------------------------------------------
# Monthly workbooks (unchanged storage locations)
# -------------------------------------------------------

def get_workbook_path(year, month):
    return f"{int(year)}/{month}.xlsx"


def upload_workbook(file_bytes, year, month):
    return _upload_object(get_workbook_path(year, month), file_bytes)


def download_workbook(year, month):
    file_bytes = _download_object(get_workbook_path(year, month))

    if file_bytes is None:
        return None

    with tempfile.NamedTemporaryFile(
        delete=False, suffix=".xlsx"
    ) as temp:
        temp.write(file_bytes)
        return temp.name


def sync_workbook(local_path, year, month):
    with open(local_path, "rb") as file:
        return upload_workbook(file.read(), year, month)


# -------------------------------------------------------
# Persistent master template (separate from monthly files)
# -------------------------------------------------------

def download_master_template():
    # Returns None only if this particular object does not exist.
    # Network/auth/bucket errors are raised, not treated as missing.
    return _download_object(MASTER_TEMPLATE_OBJECT)


def upload_master_template(file_bytes):
    return _upload_object(MASTER_TEMPLATE_OBJECT, file_bytes)


def get_master_template_bytes(bundled_template_path):
    template_bytes = download_master_template()
    if template_bytes is not None:
        return template_bytes

    # First-time setup: before the first master sync, use the repo template.
    with open(bundled_template_path, "rb") as file:
        return file.read()
