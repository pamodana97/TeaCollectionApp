
import streamlit as st
def initialize():
    defaults={
        "updated_cells":[],
        "saved_workbooks":[],
        "show_saved_panel":False,
        "temp_file":None,
        "uploaded_filename":None
    }
    for k,v in defaults.items():
        st.session_state.setdefault(k,v)

    # ---------------------------------------------------
    # Audit Panel
    # ---------------------------------------------------

    if "show_audit_panel" not in st.session_state:
        st.session_state["show_audit_panel"] = False

    if "audit_authenticated" not in st.session_state:
        st.session_state["audit_authenticated"] = False
