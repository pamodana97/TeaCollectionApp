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