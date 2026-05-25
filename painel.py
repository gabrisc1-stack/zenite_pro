import streamlit as st
import pandas as pd
import os
from io import BytesIO


ARQUIVO_EXCEL = r"C:\Users\gabri\OneDrive\Área de Trabalho\automacao_precos\PRECIFICAÇÃO IPHONE.xlsx"
SCRIPT_ROBO = r"C:\Users\gabri\OneDrive\Área de Trabalho\automacao_precos\app.py"

st.set_page_config(page_title="Zenite Pro", page_icon="📱", layout="wide")

st.title("📱 Zenite Pro")
st.subheader("Sistema Inteligente de Preços Apple")

if st.button("🔄 Atualizar preços"):
    os.system(f'python "{SCRIPT_ROBO}"')
    st.success("Atualização finalizada!")

st.divider()
st.subheader("📊 Produtos atualizados")

df = pd.read_excel(ARQUIVO_EXCEL)
df = df.fillna(0)
colunas_num = [
    "Menor_Preço",
    "SEDEX",
    "CUSTO INTERMEDIADOR",
    "CUSTO EMBALAGEM",
    "LUCRO"
]

for coluna in colunas_num:
    df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

df["VALOR DE VENDA"] = (
    df["Menor_Preço"]
    + df["SEDEX"]
    + df["CUSTO INTERMEDIADOR"]
    + df["CUSTO EMBALAGEM"]
    + df["LUCRO"]
)

df["CUSTO NF(1%)"] = df["VALOR DE VENDA"] * 0.01

st.dataframe(df)

buffer = BytesIO()
df.to_excel(buffer, index=False)
buffer.seek(0)

st.download_button(
    label="📥 Baixar planilha atualizada",
    data=buffer,
    file_name="precos_atualizados_zenite.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)