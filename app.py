import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("Gerador Profissional de Balanço CEAGESP")

ano = st.number_input("Ano", value=2026)
mes = st.selectbox(
"Mês",
["01","02","03","04","05","06","07","08","09","10","11","12"]
)

csv_capital = st.file_uploader("CSV Capital", type=["csv"])
csv_interior = st.file_uploader("CSV Interior", type=["csv"])
template = st.file_uploader("Template Excel", type=["xlsx"])
ano_anterior = st.file_uploader("Balanço Ano Anterior (opcional)", type=["xlsx"])

def detectar_coluna(colunas, palavras):
for c in colunas:
for p in palavras:
if p in c.lower():
return c
return None

if st.button("GERAR BALANÇO"):

```
if not csv_capital or not csv_interior or not template:
    st.error("Envie todos os arquivos obrigatórios")
    st.stop()

df_cap = pd.read_csv(csv_capital, sep=";", encoding="latin1")
df_int = pd.read_csv(csv_interior, sep=";", encoding="latin1")

df = pd.concat([df_cap, df_int])

col_prod = detectar_coluna(df.columns, ["prod"])
col_ton = detectar_coluna(df.columns, ["ton","quant","peso"])

if col_prod is None or col_ton is None:
    st.error("Colunas de produto ou tonelada não encontradas")
    st.stop()

resumo = (
    df.groupby(col_prod)[col_ton]
    .sum()
    .reset_index()
    .sort_values(col_ton, ascending=False)
)

wb = load_workbook(template)

aba_nome = f"BALANCO_{mes}_{ano}"
ws = wb.create_sheet(aba_nome)

ws.append(["Produto","Tonelada"])

for _,row in resumo.iterrows():
    ws.append([row[col_prod], float(row[col_ton])])

buffer = BytesIO()
wb.save(buffer)

nome_arquivo = f"balanco_{mes}_{ano}.xlsx"

st.success("Balanço gerado!")

st.download_button(
    "Baixar Excel",
    data=buffer.getvalue(),
    file_name=nome_arquivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
```
