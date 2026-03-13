import streamlit as st
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

st.title("Gerador de Balanço CEAGESP")

csv_capital = st.file_uploader("CSV Capital", type=["csv"])
csv_interior = st.file_uploader("CSV Interior", type=["csv"])

def encontrar_coluna(colunas, palavras):
for c in colunas:
nome = c.lower()
for p in palavras:
if p in nome:
return c
return None

if st.button("Gerar balanço"):

```
if csv_capital is None or csv_interior is None:
    st.error("Envie os dois arquivos CSV")
    st.stop()

df_cap = pd.read_csv(csv_capital, sep=";", encoding="latin1")
df_int = pd.read_csv(csv_interior, sep=";", encoding="latin1")

df = pd.concat([df_cap, df_int], ignore_index=True)

col_produto = encontrar_coluna(df.columns, ["prod"])
col_ton = encontrar_coluna(df.columns, ["ton", "quant", "peso"])

if col_produto is None:
    st.error("Coluna de produto não encontrada")
    st.write(df.columns)
    st.stop()

if col_ton is None:
    st.error("Coluna de tonelada não encontrada")
    st.write(df.columns)
    st.stop()

resumo = (
    df.groupby(col_produto)[col_ton]
    .sum()
    .reset_index()
    .sort_values(col_ton, ascending=False)
)

wb = Workbook()
ws = wb.active
ws.title = "BALANCO"

ws.append(["Produto", "Tonelada"])

for _, row in resumo.iterrows():
    ws.append([row[col_produto], float(row[col_ton])])

buffer = BytesIO()
wb.save(buffer)

st.success("Balanço gerado com sucesso!")

st.download_button(
    label="Baixar Excel",
    data=buffer.getvalue(),
    file_name="balanco_ceagesp.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
```
