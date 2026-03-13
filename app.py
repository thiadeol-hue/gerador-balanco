import streamlit as st
import pandas as pd

st.title("Gerador de Balanço CEAGESP")

capital = st.file_uploader("CSV CAPITAL")
interior = st.file_uploader("CSV INTERIOR")

mes = st.number_input("Mês",1,12,2)
ano = st.number_input("Ano",2020,2035,2026)

if st.button("GERAR BALANÇO"):

    if capital and interior:

        df_cap = pd.read_csv(capital, sep=";", encoding="latin1")
        df_int = pd.read_csv(interior, sep=";", encoding="latin1")

        df = pd.concat([df_cap, df_int])

        resumo = (
            df.groupby("Produto")["Tonelada"]
            .sum()
            .reset_index()
        )

        nome = f"Balanco_{mes:02d}_{ano}.xlsx"

        resumo.to_excel(nome, index=False)

        with open(nome, "rb") as f:
            st.download_button("Baixar Balanço", f, file_name=nome)
