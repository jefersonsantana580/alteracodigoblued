
import streamlit as st
import pandas as pd
import re
from itertools import count
from io import BytesIO
from pathlib import Path

MES_MAP = {
    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
}

def run_saldo_mestre(input_excel, output_buffer):
    xls = pd.ExcelFile(input_excel)

    filas = xls.parse('FILAS')
    delta_sheet = [s for s in xls.sheet_names if 'delta' in s.lower()][0]
    delta = xls.parse(delta_sheet)

    filas.columns = filas.columns.str.strip()
    delta.columns = delta.columns.str.strip()

    for col in ['PRODUCT PROPOSTO', 'NR_FILA']:
        if col not in filas.columns:
            filas[col] = ''

    filas[['PRODUCT PROPOSTO', 'NR_FILA']] = filas[
        ['PRODUCT PROPOSTO', 'NR_FILA']
    ].fillna('')

    filas['MES'] = pd.to_datetime(filas.iloc[:, 0]).dt.to_period('M')

    records = []

    for c in delta.columns:
        if c in ['PRODUCT', 'PRODUCT SERIES']:
            continue

        m = re.match(r"([A-Za-zçÇ]{3})[/\-](\d{2,4})", str(c), re.I)
        if not m:
            continue

        mes_num = MES_MAP.get(m.group(1).lower())
        if not mes_num:
            continue

        ano = int(m.group(2))
        if ano < 100:
            ano += 2000

        per = pd.Period(f"{ano}-{mes_num:02d}", freq='M')

        tmp = delta[['PRODUCT', 'PRODUCT SERIES', c]].copy()
        tmp.rename(columns={c: 'Delta'}, inplace=True)
        tmp['MES'] = per
        records.append(tmp)

    long_delta = pd.concat(records, ignore_index=True)
    long_delta = long_delta.dropna(subset=['Delta'])
    long_delta = long_delta[long_delta['Delta'] != 0]

    saldo = (
        long_delta.groupby('PRODUCT')['Delta']
        .apply(lambda s: int(s[s > 0].sum()))
        .to_dict()
    )

    for _, neg in long_delta[long_delta['Delta'] < 0].iterrows():
        cod_neg = neg['PRODUCT']
        modelo = neg['PRODUCT SERIES']
        mes = neg['MES']
        qtd = abs(int(neg['Delta']))

        filas_idx = filas[
            (filas['PRODUCT'] == cod_neg) &
            (filas['PRODUCT SERIES'] == modelo) &
            (filas['MES'] == mes) &
            (filas['PRODUCT PROPOSTO'] == '')
        ].index.tolist()

        for idx in filas_idx:
            if qtd <= 0:
                break

            candidatos = [p for p, s in saldo.items() if s > 0]
            escolhido = candidatos[0] if candidatos else None

            if escolhido:
                filas.at[idx, 'PRODUCT PROPOSTO'] = escolhido
                saldo[escolhido] -= 1
            else:
                filas.at[idx, 'PRODUCT PROPOSTO'] = 'corte'

            qtd -= 1

    inc = count(1)
    increment_rows = []

    for prod, s in saldo.items():
        if s <= 0:
            continue

        modelo = delta.loc[
            delta['PRODUCT'] == prod, 'PRODUCT SERIES'
        ].iloc[0]

        for _ in range(s):
            n = next(inc)
            row = {c: '' for c in filas.columns}
            row['NR_FILA'] = f'incremento {n}'
            row['PRODUCT PROPOSTO'] = prod
            row['PRODUCT SERIES'] = modelo
            increment_rows.append(row)

    if increment_rows:
        filas = pd.concat([filas, pd.DataFrame(increment_rows)], ignore_index=True)

    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        filas.drop(columns='MES', errors='ignore').to_excel(
            writer, sheet_name='FILAS', index=False
        )
        delta.to_excel(writer, sheet_name=delta_sheet, index=False)

# ======================
# UI
# ======================

st.set_page_config(page_title='Saldo Mestre', layout='wide')

st.title('📊 Sugestão de alteração de código em forecast')

st.info('📌 Importante: o Excel deve manter o formato padrão para o cálculo')

col1, col2 = st.columns([0.6, 1.4])

# ESQUERDA
with col1:

    st.markdown("## 📥 Baixar arquivo padrão")

    ARQUIVO_PADRAO = Path("ARQUIVO PADRAO - Ajuste de código conforme PR.xlsx")

    if ARQUIVO_PADRAO.exists():
        with open(ARQUIVO_PADRAO, "rb") as f:
            st.download_button("⬇️ Baixar arquivo padrão", f)
    else:
        st.warning("Arquivo padrão não encontrado.")

    uploaded_file = st.file_uploader('📂 Selecione o Excel', type='xlsx')
    processar = st.button('▶️ Processar', use_container_width=True)

# DIREITA
with col2:
    st.markdown("## 📊 Resumo da análise")
    resumo_container = st.empty()

# PROCESSAMENTO
if uploaded_file and processar:
    try:
        with st.spinner('Processando...'):
            output_buffer = BytesIO()
            run_saldo_mestre(uploaded_file, output_buffer)
            output_buffer.seek(0)

        st.success('✅ Arquivo gerado!')

        st.download_button(
            '⬇️ Baixar Excel Final',
            data=output_buffer,
            file_name='resultado.xlsx'
        )

        
# ===== RESUMO SIMPLES E ESTÁVEL =====

xls = pd.ExcelFile(uploaded_file)
delta_sheet = [s for s in xls.sheet_names if 'delta' in s.lower()][0]
delta = xls.parse(delta_sheet)

delta.columns = delta.columns.str.strip()

lista = []

for col in delta.columns:
    if col in ['PRODUCT', 'PRODUCT SERIES', 'PRODUCT NEED']:
        continue

    df_tmp = delta[['PRODUCT SERIES', 'PRODUCT NEED', col]].copy()
    df_tmp.columns = ['PRODUCT SERIES', 'PRODUCT NEED', 'VALOR']

    df_tmp['VALOR'] = pd.to_numeric(df_tmp['VALOR'], errors='coerce').fillna(0)
    df_tmp['MES'] = col

    lista.append(df_tmp)

df = pd.concat(lista, ignore_index=True)

# Agrupa
df_resumo = df.groupby(
    ['PRODUCT SERIES', 'PRODUCT NEED', 'MES']
)['VALOR'].sum().reset_index()

# Ordenar meses corretamente
def chave_mes(m):
    match = re.match(r"([a-zA-Z]{3})[/\-](\d{2,4})", str(m))
    if match:
        mes = MES_MAP.get(match.group(1).lower(), 0)
        ano = int(match.group(2))
        if ano < 100:
            ano += 2000
        return (ano, mes)
    return (9999, 99)

ordem_meses = sorted(df_resumo['MES'].unique(), key=chave_mes)

# Pivot
pivot = df_resumo.pivot(
    index=['PRODUCT SERIES', 'PRODUCT NEED'],
    columns='MES',
    values='VALOR'
).fillna(0)

pivot = pivot[ordem_meses]

# Total
pivot['TOTAL'] = pivot.sum(axis=1)

pivot = pivot.astype(int)

# ===== COLORAÇÃO =====
def pintar(v):
    if v > 0:
        return 'color: green'
    elif v < 0:
        return 'color: red'
    else:
        return 'color: white'

styled = pivot.style.map(pintar)

# Mostrar
    with col2:
    resumo_container.write(styled)


except Exception as e:
        st.error('❌ Erro')
        st.exception(e)
