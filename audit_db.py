import sqlite3
from datetime import datetime
from zoneinfo import ZoneInfo

DATABASE_NAME = "audit.db"


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
            username TEXT,
            action TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )

    # Upgrade existing audit databases without deleting existing records.
    for column_name, column_type in (("username", "TEXT"), ("action", "TEXT")):
        try:
            cursor.execute(
                f"ALTER TABLE audit_log ADD COLUMN {column_name} {column_type}"
            )
        except sqlite3.OperationalError:
            pass

    connection.commit()
    connection.close()


def add_audit_record(
    customer_code,
    customer_name,
    year,
    month,
    day,
    old_amount,
    new_amount,
    excel_row,
    excel_column,
    username,
    action
):

    connection = sqlite3.connect(DATABASE_NAME)
    cursor = connection.cursor()

    current_time = datetime.now(ZoneInfo("Asia/Colombo"))

    cursor.execute(
        """
        INSERT INTO audit_log (
            customer_code, customer_name,
            year, month, day,
            old_amount, new_amount,
            excel_row, excel_column,
            username, action, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            str(username),
            str(action),
            current_time.isoformat()
        )
    )

    connection.commit()
    connection.close()


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
            COALESCE(username, 'Unknown') AS username,
            COALESCE(
                action,
                CASE
                    WHEN old_amount IS NULL THEN 'Added'
                    WHEN new_amount = 0 THEN 'Deleted'
                    ELSE 'Updated'
                END
            ) AS action,
            updated_at
        FROM audit_log
        ORDER BY id DESC
        """
    )

    records = cursor.fetchall()
    connection.close()
    return records