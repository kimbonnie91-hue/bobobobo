import streamlit as st
import streamlit.components.v1 as components
import pathlib

st.set_page_config(
    page_title="LF Mall CRM 자동화 메시지 — 고객 여정 대시보드",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
#MainMenu, header, footer { visibility: hidden; }
.stApp { margin: 0; padding: 0; }
.block-container { padding: 0 !important; max-width: 100% !important; }
div[data-testid="stVerticalBlock"] { padding: 0 !important; gap: 0 !important; }
iframe { border: none; }
</style>
""", unsafe_allow_html=True)

html_path = pathlib.Path(__file__).parent / "crm_journey.html"
html_content = html_path.read_text(encoding="utf-8")

components.html(html_content, height=1100, scrolling=True)
