import os
import subprocess
from io import BytesIO

import pandas as pd
import streamlit as st


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_EXCEL = os.path.join(
    BASE_DIR,
    "app_zenite",
    "PRECIFICAÇÃO IPHONE.xlsx"
)

SCRIPT_ROBO = os.path.join(BASE_DIR, "app_zenite", "robo.py")


st.set_page_config(
    page_title="Zenite Pro",
    page_icon="📱",
    layout="wide"
)

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

st.title("📱 Zenite Pro")
st.subheader("Sistema Inteligente de Preços Apple")

if st.button("🔄 Atualizar preços"):
    resultado = subprocess.run(
        ["python", SCRIPT_ROBO],
        capture_output=True,
        text=True
    )

    if resultado.returncode == 0:
        st.success("Preços atualizados com sucesso!")
    else:
        st.error("Erro ao atualizar os preços.")
        st.code(resultado.stderr)

st.divider()
st.subheader("📊 Produtos atualizados")

df = pd.read_excel(ARQUIVO_EXCEL)
df = df.fillna(0)

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