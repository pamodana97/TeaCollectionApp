import streamlit as st
import extra_streamlit_components as stx

from datetime import datetime, timedelta, timezone


# -------------------------------------------------------
# Configuration
# -------------------------------------------------------

COOKIE_NAME = "tea_collection_session"

SESSION_TIMEOUT_MINUTES = 15

VALID_USERNAME = "admin"
VALID_PASSWORD = "admin123"


# -------------------------------------------------------
# ONE Cookie Manager
# -------------------------------------------------------

cookie_manager = stx.CookieManager(
    key="tea_collection_cookie_manager"
)


# -------------------------------------------------------
# Current UTC Time
# -------------------------------------------------------

def current_timestamp():

    return int(
        datetime.now(timezone.utc).timestamp()
    )


# -------------------------------------------------------
# Create Login Cookie
# -------------------------------------------------------

def create_login_cookie(username):

    now = current_timestamp()

    cookie_value = (
        f"{username}|{now}"
    )

    # Keep cookie available across browser refreshes.
    # Inactivity is checked separately using timestamp.
    expires_at = (
        datetime.now(timezone.utc)
        + timedelta(days=1)
    )

    cookie_manager.set(
        COOKIE_NAME,
        cookie_value,
        expires_at=expires_at
    )


# -------------------------------------------------------
# Delete Login Cookie
# -------------------------------------------------------

def delete_login_cookie():

    try:

        cookie_manager.delete(
            COOKIE_NAME
        )

    except Exception:
        pass


# -------------------------------------------------------
# Read Existing Login Cookie
# -------------------------------------------------------

def restore_login_from_cookie():

    cookie_value = cookie_manager.get(
        COOKIE_NAME
    )

    if not cookie_value:
        return False

    try:

        username, timestamp = (
            cookie_value.split("|")
        )

        timestamp = int(timestamp)

    except (ValueError, AttributeError):

        delete_login_cookie()

        return False

    now = current_timestamp()

    inactive_seconds = (
        now - timestamp
    )

    timeout_seconds = (
        SESSION_TIMEOUT_MINUTES * 60
    )

    # Session expired
    if inactive_seconds > timeout_seconds:

        delete_login_cookie()

        st.session_state["logged_in"] = False
        st.session_state["username"] = ""

        return False

    # Restore session
    st.session_state["logged_in"] = True
    st.session_state["username"] = username

    return True


# -------------------------------------------------------
# Refresh Activity Timestamp
# -------------------------------------------------------

def refresh_login_cookie():

    if not st.session_state.get(
        "logged_in",
        False
    ):
        return

    username = st.session_state.get(
        "username",
        ""
    )

    if username:

        create_login_cookie(
            username
        )


# -------------------------------------------------------
# Logout
# -------------------------------------------------------

def logout():

    try:
        cookie_manager.delete(
            COOKIE_NAME
        )
    except Exception:
        pass

    st.session_state["logged_in"] = False
    st.session_state["username"] = ""


# -------------------------------------------------------
# Login
# -------------------------------------------------------

def login():

    # ---------------------------------------------------
    # Already Logged In
    # ---------------------------------------------------

    if st.session_state.get(
        "logged_in",
        False
    ):

        refresh_login_cookie()

        return True


    # ---------------------------------------------------
    # Try Restoring Browser Session
    # ---------------------------------------------------

    if restore_login_from_cookie():

        return True


    # ---------------------------------------------------
    # Login Page
    # ---------------------------------------------------

    st.title(
        "🍃 Tea Collection System"
    )

    st.subheader(
        "Login"
    )

    # ---------------------------------------------------
    # Login Form
    # ---------------------------------------------------

    with st.form(
        "login_form"
    ):

        username = st.text_input(
            "Username",
            placeholder="Enter username",
            key="login_username"
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password",
            key="login_password"
        )

        login_clicked = st.form_submit_button(
            "Login",
            use_container_width=True
        )


    # ---------------------------------------------------
    # Login Verification
    # ---------------------------------------------------

    if login_clicked:

        if (
            username == VALID_USERNAME
            and password == VALID_PASSWORD
        ):

            st.session_state[
                "logged_in"
            ] = True

            st.session_state[
                "username"
            ] = username

            create_login_cookie(
                username
            )

            st.rerun()

        else:

            st.error(
                "Invalid username or password."
            )
            
    return False