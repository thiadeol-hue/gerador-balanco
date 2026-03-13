import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("Gerador de Balanço")

template = st.file_uploader("Template Excel", type=["xlsx"])
csv_capital = st.file_uploader("CSV Capital", type=["csv"])
csv_interior = st.file_uploader("CSV Interior", type=["csv"])

def detectar_coluna(colunas, palavras):
for c in colunas:
for p in palavras:
if p in c.lower():
return c
return None

if st.button("Gerar balanço"):

```
if template is None or csv_capital is None or csv_interior is None:
    st.error("Envie todos os arquivos")
else:

    df_cap = pd.read_csv(csv_capital, sep=";", encoding="latin1")
    df_int = pd.read_csv(csv_interior, sep=";", encoding="latin1")

    df = pd.concat([df_cap, df_int])

    col_produto = detectar_coluna(df.columns, ["prod"])
    col_ton = detectar_coluna(df.columns, ["ton", "quant", "peso"])

    if col_produto is None or col_ton is None:
        st.error("Não encontrei colunas de produto ou tonelada")
    else:

        resumo = (
            df.groupby(col_produto)[col_ton]
            .sum()
            .reset_index()
            .sort_values(col_ton, ascending=False)
        )

        wb = load_workbook(template)
        ws = wb.create_sheet("BALANCO")

        ws.append(["Produto", "Tonelada"])

        for _, row in resumo.iterrows():
            ws.append([row[col_produto], float(row[col_ton])])

        buffer = BytesIO()
        wb.save(buffer)

        st.success("Balanço gerado")

        st.download_button(
            "Baixar Excel",
            data=buffer.getvalue(),
            file_name="balanco_gerado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
```
