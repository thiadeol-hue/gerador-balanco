import streamlit as st
import pandas as pd

st.title("Gerador de Balanço CEAGESP")

capital = st.file_uploader("CSV CAPITAL")
interior = st.file_uploader("CSV INTERIOR")

mes = st.number_input("Mês",1,12,2)
ano = st.number_input("Ano",2020,2035,2026)

def ler_csv_robusto(arquivo):
    try:
        df = pd.read_csv(arquivo, sep=";", encoding="latin1")
    except:
        try:
            df = pd.read_csv(arquivo, sep=",", encoding="latin1")
        except:
            df = pd.read_csv(arquivo, encoding="latin1")
    return df

if st.button("GERAR BALANÇO"):

    if capital and interior:

        df_cap = ler_csv_robusto(capital)
        df_int = ler_csv_robusto(interior)

        df = pd.concat([df_cap, df_int], ignore_index=True)

        # limpar nomes de colunas
        df.columns = df.columns.str.strip()

        # detectar coluna produto
        produto_col = None
        for c in df.columns:
            if "prod" in c.lower():
                produto_col = c

        # detectar coluna quantidade
        quant_col = None
        for c in df.columns:
            if "ton" in c.lower() or "quant" in c.lower() or "kg" in c.lower():
                quant_col = c

        if produto_col is None or quant_col is None:
            st.error("Não foi possível identificar automaticamente as colunas de produto e quantidade.")
        else:

            resumo = (
                df.groupby(produto_col)[quant_col]
                .sum()
                .reset_index()
            )

            nome = f"Balanco_{mes:02d}_{ano}.xlsx"

            resumo.to_excel(nome, index=False)

            with open(nome, "rb") as f:
                st.download_button("Baixar Balanço", f, file_name=nome)
