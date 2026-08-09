import sqlite3
from datetime import datetime
from zoneinfo import ZoneInfo


DATABASE_NAME = "audit.db"


# -------------------------------------------------------
# Create / Upgrade Audit Database
# -------------------------------------------------------

def initialize_audit_db():

    connection = sqlite3.connect(
        DATABASE_NAME
    )

    cursor = connection.cursor()

    # ---------------------------------------------------
    # Collection Audit
    # ---------------------------------------------------

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

    # ---------------------------------------------------
    # Upgrade Existing Collection Audit Database
    # ---------------------------------------------------

    for column_name, column_type in (
        ("username", "TEXT"),
        ("action", "TEXT")
    ):

        try:

            cursor.execute(
                f"""
                ALTER TABLE audit_log
                ADD COLUMN {column_name}
                {column_type}
                """
            )

        except sqlite3.OperationalError:

            # Column already exists
            pass


    # ---------------------------------------------------
    # Customer Audit
    # ---------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS customer_audit_log (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            customer_code TEXT NOT NULL,

            old_name TEXT,
            new_name TEXT,

            action TEXT NOT NULL,

            username TEXT NOT NULL,

            updated_at TEXT NOT NULL
        )
        """
    )


    connection.commit()
    connection.close()


# -------------------------------------------------------
# Add Collection Audit Record
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
    excel_column,
    username,
    action
):

    connection = sqlite3.connect(
        DATABASE_NAME
    )

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

            username,
            action,

            updated_at
        )

        VALUES (
            ?, ?, ?, ?, ?,
            ?, ?, ?, ?,
            ?, ?, ?
        )
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


# -------------------------------------------------------
# Get Collection Audit Records
# -------------------------------------------------------

def get_audit_records():

    connection = sqlite3.connect(
        DATABASE_NAME
    )

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

            COALESCE(
                username,
                'Unknown'
            ) AS username,

            COALESCE(
                action,

                CASE

                    WHEN old_amount IS NULL
                        THEN 'Added'

                    WHEN new_amount = 0
                        THEN 'Deleted'

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


# =======================================================
# CUSTOMER AUDIT
# =======================================================


# -------------------------------------------------------
# Add Customer Audit Record
# -------------------------------------------------------

def add_customer_audit_record(
    customer_code,
    old_name,
    new_name,
    action,
    username
):

    connection = sqlite3.connect(
        DATABASE_NAME
    )

    cursor = connection.cursor()


    current_time = datetime.now(
        ZoneInfo("Asia/Colombo")
    )


    cursor.execute(
        """
        INSERT INTO customer_audit_log (

            customer_code,
            old_name,
            new_name,
            action,
            username,
            updated_at
        )

        VALUES (?, ?, ?, ?, ?, ?)
        """,

        (
            str(customer_code),

            (
                str(old_name)
                if old_name is not None
                else None
            ),

            (
                str(new_name)
                if new_name is not None
                else None
            ),

            str(action),
            str(username),

            current_time.isoformat()
        )
    )


    connection.commit()
    connection.close()


# -------------------------------------------------------
# Get Customer Audit Records
# -------------------------------------------------------

def get_customer_audit_records():

    connection = sqlite3.connect(
        DATABASE_NAME
    )

    cursor = connection.cursor()


    cursor.execute(
        """
        SELECT

            id,
            customer_code,
            old_name,
            new_name,
            action,
            username,
            updated_at

        FROM customer_audit_log

        ORDER BY id DESC
        """
    )


    records = cursor.fetchall()

    connection.close()

    return records