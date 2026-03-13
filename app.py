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

        df.columns = df.columns.str.strip()

        produto_col = None
        quant_col = None

        for c in df.columns:
            if "prod" in c.lower():
                produto_col = c
            if "ton" in c.lower() or "quant" in c.lower() or "kg" in c.lower():
                quant_col = c

        if produto_col is None or quant_col is None:
            st.error("Não foi possível identificar as colunas de produto e quantidade.")
        else:

            ranking = (
                df.groupby(produto_col)[quant_col]
                .sum()
                .reset_index()
                .sort_values(quant_col, ascending=False)
            )

            top20 = ranking.head(20)

            resumo = pd.DataFrame({
                "Indicador": [
                    "Total de produtos",
                    "Volume total"
                ],
                "Valor": [
                    ranking.shape[0],
                    ranking[quant_col].sum()
                ]
            })

            nome = f"Balanco_{mes:02d}_{ano}.xlsx"

            with pd.ExcelWriter(nome, engine="openpyxl") as writer:

                df.to_excel(writer, sheet_name="Dados Consolidados", index=False)
                ranking.to_excel(writer, sheet_name="Ranking Produtos", index=False)
                top20.to_excel(writer, sheet_name="Top 20", index=False)
                resumo.to_excel(writer, sheet_name="Resumo", index=False)

            with open(nome, "rb") as f:
                st.download_button("Baixar Balanço", f, file_name=nome)
