import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Provisões Encargos Folha de Pagamento", layout="wide")


def carregar_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, header=0)
    df = df.astype(str).fillna('').applymap(lambda x: str(x).strip())

    def normalizar_nome(nome):
        import unicodedata
        n = str(nome).strip()
        n = ''.join(c for c in unicodedata.normalize(
            'NFKD', n) if not unicodedata.combining(c))
        n = n.replace(' ', '_').replace('-', '_').upper()
        return n

    df.columns = [normalizar_nome(c) for c in df.columns]

    counts = {}
    new_cols = []
    for c in df.columns:
        if c in counts:
            counts[c] += 1
            new_cols.append(f"{c}_{counts[c]}")
        else:
            counts[c] = 0
            new_cols.append(c)
    df.columns = new_cols

    return df


def limpar_valor(x):
    if pd.isna(x):
        return 0.0
    s = str(x).strip().replace('\xa0', '').replace('\n', '').replace(' ', '')
    s = s.replace(',', '.')
    try:
        v = float(s)
    except:
        v = 0.0
    return abs(v)


def tratar_planilha(df):
    df_tratado = df.copy()
    for col in df.columns:
        if 'COD' not in col.upper():
            if df[col].astype(str).str.contains(r"\d").any():
                try:
                    df_tratado[col] = df[col].apply(limpar_valor)
                except:
                    pass
    total_debito = 0.0
    total_credito = 0.0
    for col in df.columns:
        if 'PARTIDA' in col.upper():
            try:
                val_col = df.columns[df.columns.get_loc(col) - 1]
                partida_vals = df[col].fillna('').astype(str).str.upper()
                val_vals = df[val_col].apply(limpar_valor)
                total_debito += val_vals[partida_vals == 'D'].sum()
                total_credito += val_vals[partida_vals == 'C'].sum()
            except:
                pass
    return df_tratado, total_debito, total_credito


def gerar_layout_final(df, lote, competencia, cpf_cnpj, complemento):
    header = [
        'TIPO', 'COD LOTE', 'VLR CONTABIL LOTE', 'COMPETENCIA', 'COD FILIAL', 'COD CONTA CONTABIL',
        'VLR CONTABIL', 'PARTIDA', 'COD HISTORICO', 'COMPLEMENTO', 'COD CENTRO CUSTO', 'VLR CENTRO CUSTO',
        'CPF/CNPJ', 'IMOBILIZADO', 'VLR IMOBILIZADO'
    ]

    out_rows = []
    cols = list(df.columns)
    value_labels = ['FERIAS', 'INSS', 'FGTS', 'PIS']

    cod_historico_map = {
        'FERIAS': '868',
        'INSS': '869',
        'FGTS': '870',
        'PIS': '871'
    }

    grupos = {}
    for lbl in value_labels:
        val_col = next((c for c in cols if lbl in c), None)
        if val_col:
            idx_val = cols.index(val_col)
            cod_col = None
            for i in range(idx_val - 1, -1, -1):
                if 'COD' in cols[i].upper() and 'CC' not in cols[i].upper():
                    cod_col = cols[i]
                    break
                if any(x in cols[i] for x in value_labels):
                    break
            partida_col = next((cols[i] for i in range(
                idx_val + 1, len(cols)) if 'PARTIDA' in cols[i]), None)
            grupos[lbl] = {'valor': val_col,
                           'codigo': cod_col, 'partida': partida_col}

    vlr_contabil_lote = 0.0
    for g in grupos.values():
        if g['valor'] and g['partida']:
            partida_vals = df[g['partida']].fillna('').astype(str).str.upper()
            valores = df[g['valor']].apply(limpar_valor)
            vlr_contabil_lote += valores[partida_vals == 'D'].sum()

    lot_line = {h: '' for h in header}
    lot_line['TIPO'] = 'LOT'
    lot_line['COD LOTE'] = lote
    lot_line['COMPETENCIA'] = pd.to_datetime(
        competencia).strftime('%d/%m/%Y') if competencia else ''
    lot_line['VLR CONTABIL LOTE'] = f"{vlr_contabil_lote:.2f}"
    out_rows.append(lot_line)

    total_idx = None
    for i, r in df.iterrows():
        if any(isinstance(x, str) and 'TOTAL' in x.upper() for x in r.values if pd.notna(x)):
            total_idx = i
            break
    iter_df = df.loc[:(total_idx - 1)] if total_idx is not None else df

    for _, row in iter_df.iterrows():
        # Forçar COD FILIAL (UNIDADE) a inteiro quando aplicável
        unidade_raw = row.get('UNIDADE', '')
        try:
            unidade = int(float(str(unidade_raw).replace(',', '.'))) if str(
                unidade_raw).strip() not in ['', 'nan', 'None'] else ''
        except:
            unidade = str(unidade_raw).strip()

        cc_raw = row.get('CC', '')
        cc = '' if pd.isna(cc_raw) else str(cc_raw).strip().replace(',', '.')

        for lbl, g in grupos.items():
            valor = limpar_valor(row.get(g['valor'], 0))
            if valor == 0:
                continue

            cod_conta = str(row.get(g['codigo'], '')).strip()
            if cod_conta in ['nan', 'None', '0', '']:
                continue
            partida = row.get(g['partida'], '') if g['partida'] else ''
            cod_hist = cod_historico_map.get(lbl, '')

            con_line = {h: '' for h in header}
            con_line['TIPO'] = 'CON'
            con_line['COD FILIAL'] = unidade
            con_line['COD CONTA CONTABIL'] = cod_conta
            con_line['VLR CONTABIL'] = f"{valor:.2f}"
            con_line['PARTIDA'] = partida
            con_line['COD HISTORICO'] = cod_hist
            con_line['COMPLEMENTO'] = complemento
            con_line['CPF/CNPJ'] = cpf_cnpj
            out_rows.append(con_line)

            if unidade not in ['99', 'TOTAL'] and cc != '':
                cus_line = {h: '' for h in header}
                cus_line['TIPO'] = 'CUS'
                cus_line['COD CENTRO CUSTO'] = cc
                cus_line['VLR CENTRO CUSTO'] = f"{valor:.2f}"
                out_rows.append(cus_line)

    if total_idx is not None and total_idx < len(df):
        total_row = df.loc[total_idx]
        unidade_raw = total_row.get('UNIDADE', '')
        try:
            unidade_total = int(float(str(unidade_raw).replace(',', '.'))) if str(
                unidade_raw).strip() not in ['', 'nan', 'None'] else ''
        except:
            unidade_total = str(unidade_raw).strip()

        for lbl, g in grupos.items():
            valor_total = limpar_valor(total_row.get(g['valor'], 0))
            if valor_total == 0:
                continue
            cod_conta_total = str(total_row.get(
                g['codigo'], '')).strip() if g['codigo'] else ''
            if cod_conta_total.lower() in ['nan', 'none', '0', '']:
                continue
            partida_total = total_row.get(
                g['partida'], '') if g['partida'] else ''
            cod_hist_total = cod_historico_map.get(lbl, '')

            con_line = {h: '' for h in header}
            con_line['TIPO'] = 'CON'
            con_line['COD FILIAL'] = unidade_total
            con_line['COD CONTA CONTABIL'] = cod_conta_total
            con_line['VLR CONTABIL'] = f"{valor_total:.2f}"
            con_line['PARTIDA'] = partida_total
            con_line['COD HISTORICO'] = cod_hist_total
            con_line['COMPLEMENTO'] = complemento
            con_line['CPF/CNPJ'] = cpf_cnpj
            out_rows.append(con_line)

    return pd.DataFrame(out_rows, columns=header)


def to_excel_bytes(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue()


# ----------------------- Streamlit UI -----------------------
st.title('Provisões Encargos Folha de Pagamento')

st.markdown('#### Passos para uso:')
st.markdown('1️⃣ Faça upload do arquivo Excel original.  \n2️⃣ Confira a visualização conforme a planilha.  \n3️⃣ Clique em **Tratar Planilha** e depois em **Gerar Arquivo Contábil**.')

uploaded_file = st.file_uploader(
    'Carregar arquivo Excel', type=['xlsx', 'xls'])

if uploaded_file is not None:
    df_raw = carregar_excel(uploaded_file)
    st.subheader('Visualização Original (conforme planilha importada)')
    st.dataframe(df_raw, use_container_width=True)

    lote = st.number_input('Código do Lote', min_value=0, step=1, value=0)
    competencia = st.date_input('Competência', value=datetime.today())
    cpf_cnpj = st.text_input('CPF/CNPJ')
    complemento = st.text_input('Complemento (descrição)')

    # Armazenar complemento no session_state para garantir que o valor digitado seja usado
    st.session_state['complemento'] = complemento

    if st.button('Tratar Planilha'):
        with st.spinner('Processando planilha...'):
            df_tratado, debito, credito = tratar_planilha(df_raw)
            st.session_state['df_tratado'] = df_tratado
            st.session_state['totais'] = (debito, credito)

        st.success('Planilha tratada com sucesso!')
        st.metric('Total Débito', f"R$ {debito:,.2f}")
        st.metric('Total Crédito', f"R$ {credito:,.2f}")
        st.subheader('Visualização Tratada')
        st.dataframe(df_tratado, use_container_width=True)

    if 'df_tratado' in st.session_state:
        st.markdown('---')
        st.subheader('Gerar Arquivo Contábil')
        if st.button('Gerar Arquivo Contábil'):
            with st.spinner('Gerando layout final...'):
                # pegar complemento do session_state (garante o valor digitado)
                complemento_to_use = st.session_state.get('complemento', '')
                df_final = gerar_layout_final(
                    st.session_state['df_tratado'], lote, competencia, cpf_cnpj, complemento_to_use
                )
                st.session_state['df_final'] = df_final
            st.success('Layout contábil gerado com sucesso!')
            st.dataframe(st.session_state['df_final'].head(
                200), use_container_width=True)

        if 'df_final' in st.session_state:
            excel_bytes = to_excel_bytes(st.session_state['df_final'])
            st.download_button(
                'Download Arquivo Contábil',
                data=excel_bytes,
                file_name='layout_contabil.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
else:
    st.info('Aguardando upload do arquivo Excel...')

# fim do arquivo
