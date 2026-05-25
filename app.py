import streamlit as st
import os
import pandas as pd
from io import BytesIO

import streamlit as st

st.markdown("""
<style>
div[data-baseweb="select"] > div {
    font-size: 16px;
}

div[role="listbox"] {
    max-height: 250px !important;
    overflow-y: auto !important;
}
</style>
""", unsafe_allow_html=True)

ARQUIVO_EXCEL = r"C:\Users\gabri\OneDrive\Área de Trabalho\automacao_precos\PRECIFICAÇÃO IPHONE.xlsx"
LOGIN_JSON = r"C:\Users\gabri\OneDrive\Área de Trabalho\automacao_precos\login.json"
ARQUIVO_EXCEL = r"C:\Users\gabri\OneDrive\Área de Trabalho\automacao_precos\PRECIFICAÇÃO IPHONE.xlsx"

st.set_page_config(
    page_title="Zenite Pro",
    page_icon="📱",
    layout="wide"
)

st.title("📱 Zenite Pro")
st.subheader("Sistema Inteligente de Preços Apple")

if st.button("🔄 Atualizar preços"):
    os.system(r'python "C:\Users\gabri\OneDrive\Área de Trabalho\automacao_precos\app.py"')
    st.success("Preços atualizados com sucesso!")

st.divider()

st.subheader("📊 Produtos atualizados")

df = pd.read_excel(ARQUIVO_EXCEL)

st.dataframe(df, use_container_width=True)

buffer = BytesIO()
df.to_excel(buffer, index=False)
buffer.seek(0)

st.download_button(
    label="📥 Baixar planilha atualizada",
    data=buffer,
    file_name="precos_atualizados_zenite.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)