
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
