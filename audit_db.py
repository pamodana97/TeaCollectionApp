import sqlite3
from datetime import datetime
from zoneinfo import ZoneInfo


DATABASE_NAME = "audit.db"


# -------------------------------------------------------
# Create Audit Database
# -------------------------------------------------------

def initialize_audit_db():

    connection = sqlite3.connect(DATABASE_NAME)
    cursor = connection.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_log (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            customer_code TEXT NOT NULL,
            customer_name TEXT,

            year INTEGER NOT NULL,
            month TEXT NOT NULL,
            day INTEGER NOT NULL,

            old_amount REAL,
            new_amount REAL NOT NULL,

            excel_row INTEGER,
            excel_column INTEGER,

            updated_at TEXT NOT NULL
        )
        """
    )

    connection.commit()
    connection.close()


# -------------------------------------------------------
# Add Audit Record
# -------------------------------------------------------

def add_audit_record(
    customer_code,
    customer_name,
    year,
    month,
    day,
    old_amount,
    new_amount,
    excel_row,
    excel_column
):

    connection = sqlite3.connect(DATABASE_NAME)
    cursor = connection.cursor()

    current_time = datetime.now(
        ZoneInfo("Asia/Colombo")
    )

    cursor.execute(
        """
        INSERT INTO audit_log (
            customer_code,
            customer_name,
            year,
            month,
            day,
            old_amount,
            new_amount,
            excel_row,
            excel_column,
            updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            str(customer_code),
            str(customer_name),
            int(year),
            str(month),
            int(day),

            old_amount,
            float(new_amount),

            int(excel_row),
            int(excel_column),

            current_time.isoformat()
        )
    )

    connection.commit()
    connection.close()


# -------------------------------------------------------
# Get Audit Records
# -------------------------------------------------------

def get_audit_records():

    connection = sqlite3.connect(DATABASE_NAME)
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT
            id,
            customer_code,
            customer_name,
            year,
            month,
            day,
            old_amount,
            new_amount,
            updated_at
        FROM audit_log
        ORDER BY id DESC
        """
    )

    records = cursor.fetchall()

    connection.close()

    return records