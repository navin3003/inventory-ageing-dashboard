from pathlib import Path


import streamlit as st

import streamlit.components.v1 as components




st.set_page_config(

    page_title="Inventory Ageing Dashboard",

    page_icon="📊",

    layout="wide",

    initial_sidebar_state="collapsed",

)


html_file = Path(__file__).parent / "dashboard.html"


if not html_file.exists():

    st.error("dashboard.html was not found in the repository.")

    st.stop()


dashboard_html = html_file.read_text(encoding="utf-8")


components.html(

    dashboard_html,

    height=1600,

    scrolling=True,

)
