
import streamlit as st
import pandas as pd
import re
from itertools import count
from io import BytesIO

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

    # Garantir colunas
    for col in ['PRODUCT PROPOSTO', 'NR_FILA']:
        if col not in filas.columns:
            filas[col] = ''

    filas[['PRODUCT PROPOSTO', 'NR_FILA']] = filas[
        ['PRODUCT PROPOSTO', 'NR_FILA']
    ].fillna('')

    # Mês da FILAS (primeira coluna)
    filas['MES'] = pd.to_datetime(filas.iloc[:, 0]).dt.to_period('M')

    # Explodir delta mensal
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
    long_delta = long_delta.sort_values('MES')

    # Saldo mestre
    saldo = (
        long_delta.groupby('PRODUCT')['Delta']
        .apply(lambda s: int(s[s > 0].sum()))
        .to_dict()
    )

    # Substituições e cortes
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
            candidatos_same = [
                p for p in candidatos
                if p in delta[delta['PRODUCT SERIES'] == modelo]['PRODUCT'].values
            ]

            escolhido = (
                candidatos_same[0]
                if candidatos_same
                else (candidatos[0] if candidatos else None)
            )

            if escolhido:
                filas.at[idx, 'PRODUCT PROPOSTO'] = escolhido
                saldo[escolhido] -= 1
            else:
                filas.at[idx, 'PRODUCT PROPOSTO'] = 'corte'

            qtd -= 1

    # Incrementos finais
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
        filas = pd.concat(
            [filas, pd.DataFrame(increment_rows)],
            ignore_index=True
        )

    # Exportar
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        filas.drop(columns='MES', errors='ignore').to_excel(
            writer, sheet_name='FILAS', index=False
        )
        delta.to_excel(writer, sheet_name=delta_sheet, index=False)


# ======================
# STREAMLIT UI
# ======================
st.set_page_config(
    page_title='Saldo Mestre — Sugestão de Filas',
    layout='centered'
)

st.title('📊 Saldo Mestre – Sugestão de Filas')
st.info(
    '📌 Importante: o Excel deve conter a aba **FILAS** '
    'e uma aba **DELTA** do ciclo.'
)

uploaded_file = st.file_uploader(
    '📂 Selecione o Excel',
    type='xlsx'
)

if uploaded_file and st.button('▶️ Processar'):
    try:
        with st.spinner('Processando arquivo...'):
            output_buffer = BytesIO()
            run_saldo_mestre(uploaded_file, output_buffer)
            output_buffer.seek(0)

        st.success('✅ Arquivo gerado com sucesso!')

        st.download_button(
            '⬇️ Baixar Excel Final',
            data=output_buffer,
            file_name='FILAS_DEFINITIVO_SALDO_MESTRE.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        st.error('❌ Erro ao processar o arquivo')
        st.exception(e)
