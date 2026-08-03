import streamlit as st
import psycopg2


# -------------------------------------------------------
# Database Connection
# -------------------------------------------------------

def get_connection():

    return psycopg2.connect(
        st.secrets["DATABASE_URL"]
    )


# -------------------------------------------------------
# Save / Update Tea Collection
# -------------------------------------------------------

def save_collection(
    customer_code,
    customer_name,
    year,
    month,
    day,
    amount
):

    connection = get_connection()

    try:

        cursor = connection.cursor()

        cursor.execute(
            """
            INSERT INTO tea_collections (
                customer_code,
                customer_name,
                collection_year,
                collection_month,
                collection_day,
                amount,
                updated_at
            )
            VALUES (
                %s, %s, %s, %s, %s, %s, NOW()
            )

            ON CONFLICT (
                customer_code,
                collection_year,
                collection_month,
                collection_day
            )

            DO UPDATE SET

                customer_name = EXCLUDED.customer_name,
                amount = EXCLUDED.amount,
                updated_at = NOW()
            """,
            (
                str(customer_code),
                str(customer_name),
                int(year),
                str(month),
                int(day),
                float(amount)
            )
        )

        connection.commit()

    finally:

        connection.close()

# -------------------------------------------------------
# Save Customers
# -------------------------------------------------------

def save_customers(customer_df):

    connection = get_connection()

    try:

        cursor = connection.cursor()

        for _, row in customer_df.iterrows():

            customer_code = str(
                row["Customer Code"]
            ).strip()

            customer_name = str(
                row["Customer Name"]
            ).strip()

            cursor.execute(
                """
                INSERT INTO customers (
                    customer_code,
                    customer_name,
                    updated_at
                )
                VALUES (
                    %s, %s, NOW()
                )

                ON CONFLICT (customer_code)

                DO UPDATE SET
                    customer_name = EXCLUDED.customer_name,
                    updated_at = NOW()
                """,
                (
                    customer_code,
                    customer_name
                )
            )

        connection.commit()

    finally:

        connection.close()

# -------------------------------------------------------
# Save Monthly Workbook
# -------------------------------------------------------

def save_monthly_workbook(
    year,
    month,
    filename
):

    connection = get_connection()

    try:

        cursor = connection.cursor()

        cursor.execute(
            """
            INSERT INTO monthly_workbooks (
                collection_year,
                collection_month,
                original_filename,
                updated_at
            )

            VALUES (
                %s, %s, %s, NOW()
            )

            ON CONFLICT (
                collection_year,
                collection_month
            )

            DO UPDATE SET
                original_filename =
                    EXCLUDED.original_filename,
                updated_at = NOW()
            """,
            (
                int(year),
                str(month),
                str(filename)
            )
        )

        connection.commit()

    finally:

        connection.close()

# -------------------------------------------------------
# Get Saved Collections
# -------------------------------------------------------

def get_collections(
    year,
    month
):

    connection = get_connection()

    try:

        cursor = connection.cursor()

        cursor.execute(
            """
            SELECT
                customer_code,
                collection_day,
                amount
            FROM tea_collections

            WHERE
                collection_year = %s
                AND collection_month = %s

            ORDER BY
                customer_code,
                collection_day
            """,
            (
                int(year),
                str(month)
            )
        )

        return cursor.fetchall()

    finally:

        connection.close()

# -------------------------------------------------------
# Delete Tea Collection
# -------------------------------------------------------

def delete_collection(
    customer_code,
    year,
    month,
    day
):

    connection = get_connection()

    try:

        cursor = connection.cursor()

        cursor.execute(
            """
            DELETE FROM tea_collections
            WHERE
                customer_code = %s
                AND collection_year = %s
                AND collection_month = %s
                AND collection_day = %s
            """,
            (
                str(customer_code),
                int(year),
                str(month),
                int(day)
            )
        )

        connection.commit()

    finally:

        connection.close()

# -------------------------------------------------------
# Get Last Collection Entry
# -------------------------------------------------------

def get_last_collection(
    year,
    month
):

    connection = get_connection()

    try:

        cursor = connection.cursor()

        cursor.execute(
            """
            SELECT
                customer_name,
                amount,
                collection_year,
                collection_month,
                collection_day,
                updated_at
            FROM tea_collections
            WHERE
                collection_year = %s
                AND collection_month = %s
            ORDER BY updated_at DESC
            LIMIT 1
            """,
            (
                int(year),
                str(month)
            )
        )

        return cursor.fetchone()

    finally:

        connection.close()