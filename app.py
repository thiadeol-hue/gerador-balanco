import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("Gerador de Balanço CEAGESP")

st.write("Envie os arquivos necessários")

template = st.file_uploader("Template Excel", type=["xlsx"])
csv_capital = st.file_uploader("CSV Capital", type=["csv"])
csv_interior = st.file_uploader("CSV Interior", type=["csv"])

if st.button("Gerar Balanço"):

```
if template is None or csv_capital is None or csv_interior is None:
    st.error("Envie todos os arquivos.")
else:

    df_cap = pd.read_csv(csv_capital, sep=";", encoding="latin1")
    df_int = pd.read_csv(csv_interior, sep=";", encoding="latin1")

    df = pd.concat([df_cap, df_int])

    col_produto = None
    col_ton = None

    for c in df.columns:
        if "prod" in c.lower():
            col_produto = c
        if "ton" in c.lower() or "quant" in c.lower():
            col_ton = c

    if col_produto is None or col_ton is None:
        st.error("Colunas Produto ou Tonelada não encontradas")
    else:

        resumo = (
            df.groupby(col_produto)[col_ton]
            .sum()
            .reset_index()
            .sort_values(col_ton, ascending=False)
        )

        wb = load_workbook(template)
        ws = wb.create_sheet("BALANCO_GERADO")

        ws.append(["Produto", "Toneladas"])

        for _, row in resumo.iterrows():
            ws.append([row[col_produto], float(row[col_ton])])

        buffer = BytesIO()
        wb.save(buffer)

        st.success("Balanço gerado!")

        st.download_button(
            label="Baixar Excel",
            data=buffer.getvalue(),
            file_name="balanco_gerado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
```
