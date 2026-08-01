import streamlit as st


def login():

    # Already logged in
    if st.session_state.get("logged_in", False):
        return True

    # ---------------------------------------------------
    # Login Page
    # ---------------------------------------------------

    st.title("🍃 Tea Collection System")

    st.subheader("Login")

    username = st.text_input(
        "Username",
        placeholder="Enter username"
    )

    password = st.text_input(
        "Password",
        type="password",
        placeholder="Enter password"
    )

    if st.button(
        "Login",
        use_container_width=True
    ):

        if username == "admin" and password == "admin123":

            st.session_state["logged_in"] = True
            st.session_state["username"] = username

            st.rerun()

        else:

            st.error(
                "Invalid username or password."
            )

    return False