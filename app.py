import json
import pandas as pd
import streamlit as st
from datetime import date
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(layout="wide", page_title="Posição Financeira Diária")
st.title("Dashboard Financeira Diária")

# ================= FUNCOES AUXILIARES =================
def formatar_br(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def converter_valor_br(valor):
    if pd.isna(valor): return 0.0
    val = str(valor).strip()
    if val == '' or val == '-' or val.upper() == 'NAN': return 0.0
    val = val.replace('R$', '').replace(' ', '')
    if ',' in val: val = val.replace('.', '').replace(',', '.')
    try: return float(val)
    except: return 0.0

def detectar_empresa(nome):
    nome = str(nome).upper()
    if 'MATRIZ' in nome: return 'MATRIZ'
    if 'WS' in nome: return 'WS'
    if 'EUSEBIO' in nome: return 'EUSEBIO'
    return 'OUTROS'

def normalizar_chave_manual(chave):
    c = str(chave).upper().strip()
    if 'NOVOS' in c: return 'NOVOS PAGOS'
    if 'USADOS' in c: return 'USADOS PAGOS'
    if 'HB' in c or 'H.B' in c: return 'H.B.PECAS'
    if 'FIDIC' in c: return 'FIDIC'
    if 'ESTOQUE' in c or 'EST.PECAS' in c: return 'ESTOQUE PECAS'
    return chave

def normalizar_tipo(titulo):
    t = str(titulo).upper().strip()
    if 'CARTEIRA' in t: return 'CARTEIRA'
    if 'MERCADO.PAGO' in t or 'MERCADO PAGO' in t: return 'MERCADO.PAGO'
    if 'VEICULO' in t or 'V-VEICULO' in t or 'VENDA VEIC' in t: return 'VEICULO'
    if 'SEGURADORA' in t or 'SEGURO' in t: return 'SEGURADORA'
    if 'GARANTIA' in t or 'VENDA GARANTIA' in t: return 'GARANTIA'
    if 'BANCO' in t: return 'BANCO'
    if 'CARTAO' in t or 'CARTÕES' in t or 'CARTAO DE CREDITO' in t or 'CREDITO' in t: return 'CARTAO'
    if 'NOVOS PAGOS' in t or 'NOVOS.PAGOS' in t: return 'NOVOS PAGOS'
    if 'USADOS PAGOS' in t or 'USADOS.PAGOS' in t or 'USADOS' in t: return 'USADOS PAGOS'
    if 'FIDIC' in t: return 'FIDIC'
    if 'H.B.PECAS' in t or 'HB PECAS' in t or 'PECAS' in t: return 'H.B.PECAS'
    if 'ESTOQUE PECAS' in t or 'EST.PECAS' in t: return 'ESTOQUE PECAS'
    if 'OBRIG' in t: return 'OBRIGACOES'
    if 'TRANSITORIA' in t: return 'TRANSITORIA'
    if 'ADIANTAMENTO' in t: return 'ADIANTAMENTO'
    if 'DIF_TRANS_ADIANT' in t: return 'DIF_TRANS_ADIANT'
    return 'OUTROS'

ITENS = [
    ('CARTEIRA', 'CARTEIRA'), ('MERCADO.PAGO', 'MERCADO.PAGO'),('VEICULO', 'VEICULO'), ('SEGURADORA', 'SEGURADORA'),
    ('GARANTIA', 'GARANTIA'), ('BANCO', 'BANCOS'), ('CARTAO', 'CARTOES'),
    ('NOVOS PAGOS', 'NOVOS.PAGOS'),('USADOS PAGOS', 'USADOS.PAGOS'),
    ('H.B.PECAS', 'H.B.PECAS'), ('FIDIC', 'FIDIC'),
    ('ESTOQUE PECAS', 'EST.PECAS'), ('OBRIGACOES', 'OBRIG. A PAGA'),
    ('TRANSITORIA', 'TRANSITORIA'), ('ADIANTAMENTO', 'ADIANTAMENTOS'), ('DIF_TRANS_ADIANT', 'DIF_TRANS_ADIANT')
]

# ================= FUNCAO DO EXCEL - PRECISA ESTAR AQUI EM CIMA =================
def gerar_excel(df_para_exportar, empresas_selecionadas):
    output = BytesIO()
    data_hoje = date.today().strftime('%d/%m/%Y')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        worksheet = workbook.create_sheet('POSICAO DIARIA')
        writer.sheets['POSICAO DIARIA'] = worksheet
        from openpyxl.styles import Font, Alignment, Border, Side
        border_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        bold = Font(bold=True, size=11); center = Alignment(horizontal='center', vertical='center'); right = Alignment(horizontal='right', vertical='center')
        
        linha_atual = 1
        worksheet.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=8)
        cell_titulo = worksheet.cell(row=linha_atual, column=1, value="POSIÇÃO FINANCEIRA DIÁRIA"); cell_titulo.font = Font(bold=True, size=14); cell_titulo.alignment = center
        worksheet.cell(row=linha_atual, column=9, value='DATA').font = bold; worksheet.cell(row=linha_atual, column=10, value=data_hoje)
        
        linha_atual += 1
        empresas_ordem = ['MATRIZ', 'WS', 'EUSEBIO']; col_inicio = 1
        for emp in empresas_ordem:
            if emp not in empresas_selecionadas: col_inicio += 3; continue
            
            linha_temp = linha_atual
            worksheet.merge_cells(start_row=linha_temp, start_column=col_inicio, end_row=linha_temp, end_column=col_inicio+1)
            worksheet.cell(row=linha_temp, column=col_inicio, value=emp).font = bold; worksheet.cell(row=linha_temp, column=col_inicio).alignment = center
            
            linha_temp += 1; total_geral = 0.0; valores_por_item = {}
            for item_chave, item_nome in ITENS:
                total = df_para_exportar[(df_para_exportar['Tipo de Título'] == item_chave) & (df_para_exportar['Empresa'] == emp)]['Saldo'].sum()
                if item_chave == 'DIF_TRANS_ADIANT':
                    trans_valor = valores_por_item.get('TRANSITORIA', 0); adiant_valor = valores_por_item.get('ADIANTAMENTO', 0)
                    total = trans_valor - adiant_valor if trans_valor > 0 else 0.0
                valores_por_item[item_chave] = total
                if item_chave == 'OBRIGACOES': total_geral -= total
                else: total_geral += total
                
                worksheet.cell(row=linha_temp, column=col_inicio, value=item_nome).border = border_fina
                cell_valor = worksheet.cell(row=linha_temp, column=col_inicio+1, value=total); cell_valor.alignment = right; cell_valor.number_format = 'R$ #,##0.00'
                
                worksheet.column_dimensions[get_column_letter(col_inicio)].width = 22
                worksheet.column_dimensions[get_column_letter(col_inicio+1)].width = 18
                
                linha_temp += 1
            
            worksheet.cell(row=linha_temp, column=col_inicio, value='TOTAL').font = bold
            cell_total = worksheet.cell(row=linha_temp, column=col_inicio+1, value=total_geral); cell_total.font = bold; cell_total.alignment = right; cell_total.number_format = 'R$ #,##0.00'
            
            col_inicio += 3

    return output.getvalue()

# ================= LEITURA DOS ARQUIVOS VIA UPLOAD =================
uploaded_files = st.file_uploader(
    "Arraste os 4 arquivos RFN aqui",
    type=['xlsx', 'xls'],
    accept_multiple_files=True
)

manual_file = st.file_uploader("Opcional: Suba o valores_manuais.json", type=['json'])

valores_iniciais = {
    'MATRIZ': {'NOVOS PAGOS': '0,00', 'USADOS PAGOS': '0,00', 'H.B.PECAS': '0,00', 'FIDIC': '0,00', 'ESTOQUE PECAS': '0,00'},
    'WS': {'NOVOS PAGOS': '0,00', 'USADOS PAGOS': '0,00', 'H.B.PECAS': '0,00', 'FIDIC': '0,00', 'ESTOQUE PECAS': '0,00'},
    'EUSEBIO': {'NOVOS PAGOS': '0,00', 'USADOS PAGOS': '0,00', 'H.B.PECAS': '0,00', 'FIDIC': '0,00', 'ESTOQUE PECAS': '0,00'}
}
if manual_file is not None:
    try:
        dados = json.load(manual_file)
        dados_norm = {}
        for emp, itens in dados.items():
            dados_norm[emp] = {}
            for chave, valor in itens.items():
                chave_norm = normalizar_chave_manual(chave)
                dados_norm[emp][chave_norm] = valor
        valores_iniciais.update(dados_norm)
    except: pass

if uploaded_files:
    dfs = {file.name: pd.read_excel(file) for file in uploaded_files}

    #... aqui vão todas as suas 4 funções carregar_... [mantém igual]
    def carregar_posicao_analitica():
        if 'financeiro.xls' not in dfs and 'financeiro.xlsx' not in dfs: return pd.DataFrame()
        df_raw = dfs.get('financeiro.xls', dfs.get('financeiro.xlsx'))
        df_raw.columns = df_raw.columns.str.strip()
        dados = []
        for empresa in df_raw['Empresa'].dropna().unique():
            df_empresa = df_raw[df_raw['Empresa'] == empresa]
            empresa_norm = detectar_empresa(empresa)
            for _, row in df_empresa.iterrows():
                try:
                    agente = str(row['Agente Cobrador']).strip()
                    titulo = str(row['Nro Titulo']).strip()
                    saldo = converter_valor_br(row['Saldo'])
                    if titulo == '' or titulo.upper() == 'NAN' or saldo <= 0: continue
                    tipo = normalizar_tipo(agente)
                    dados.append({'Tipo de Título': tipo, 'Empresa': empresa_norm, 'Saldo': saldo})
                except: continue
        return pd.DataFrame(dados) if dados else pd.DataFrame()

    def carregar_obrigacoes():
        if 'RFN003_PosicaoAnaliticoPagar.xlsx' not in dfs: return pd.DataFrame()
        df_obrig_raw = dfs['RFN003_PosicaoAnaliticoPagar.xlsx']
        obrig_dict = {'MATRIZ': 0.0, 'WS': 0.0, 'EUSEBIO': 0.0}
        for i in range(len(df_obrig_raw)):
            linha_toda = " ".join([str(x) for x in df_obrig_raw.iloc[i] if pd.notna(x)]).upper()
            if 'RESUMO' in linha_toda and 'EMPRESA' in linha_toda:
                for j in range(1, 6):
                    if i+j >= len(df_obrig_raw): break
                    emp = str(df_obrig_raw.iloc[i + j, 0])
                    val = converter_valor_br(df_obrig_raw.iloc[i + j, 9])
                    if 'EUSEBIO' in emp.upper(): obrig_dict['EUSEBIO'] = val
                    elif 'MATRIZ' in emp.upper(): obrig_dict['MATRIZ'] = val
                    elif 'WS' in emp.upper(): obrig_dict['WS'] = val
                break
        linhas = [{'Tipo de Título': 'OBRIGACOES', 'Empresa': e, 'Saldo': v} for e,v in obrig_dict.items() if v > 0]
        return pd.DataFrame(linhas)

    def carregar_creditos_nao_identificados():
        if 'RFN024_SaldoCreditosNaoIdentificados.xlsx' not in dfs: return pd.DataFrame()
        df_cred_raw = dfs['RFN024_SaldoCreditosNaoIdentificados.xlsx']
        credito_dict = {'MATRIZ': 0.0, 'WS': 0.0, 'EUSEBIO': 0.0}
        for i in range(len(df_cred_raw)):
            if not str(df_cred_raw.iloc[i, 0]).isdigit(): continue
            emp = detectar_empresa(str(df_cred_raw.iloc[i, 2]))
            sal = converter_valor_br(df_cred_raw.iloc[i, df_cred_raw.shape[1]-1])
            if emp in credito_dict and sal > 0: credito_dict[emp] += sal
        linhas = [{'Tipo de Título': 'TRANSITORIA', 'Empresa': e, 'Saldo': v} for e,v in credito_dict.items() if v > 0]
        return pd.DataFrame(linhas)

    def carregar_adiantamentos():
        if 'RFN013_FichaRazaoSaldoExcel.xlsx' not in dfs and 'RFN013_FichaRazaoSaldo_Excel.xls' not in dfs: return pd.DataFrame()
        df_ad_raw = dfs.get('RFN013_FichaRazaoSaldoExcel.xlsx', dfs.get('RFN013_FichaRazaoSaldo_Excel.xls'))
        adiant_dict = {'MATRIZ': 0.0, 'WS': 0.0, 'EUSEBIO': 0.0}
        for i in range(1, len(df_ad_raw)):
            emp = detectar_empresa(str(df_ad_raw.iloc[i, 3]))
            sal = converter_valor_br(str(df_ad_raw.iloc[i, 6]))
            if emp in adiant_dict: adiant_dict[emp] += sal
        linhas = [{'Tipo de Título': 'ADIANTAMENTO', 'Empresa': e, 'Saldo': v} for e,v in adiant_dict.items() if v > 0]
        return pd.DataFrame(linhas)

    def carregar_manuais(valores_digitados):
        dados = []
        for empresa, itens in valores_digitados.items():
            for tipo, valor_str in itens.items():
                valor = converter_valor_br(valor_str)
                if valor > 0: dados.append({'Tipo de Título': tipo, 'Empresa': empresa, 'Saldo': valor})
        return pd.DataFrame(dados)

    # ================= TELA DE LANCAMENTO MANUAL =================
    st.markdown("#### Lançamento Manual")
    col_m, col_ws, col_e = st.columns(3)
    valores_digitados = {'MATRIZ': {}, 'WS': {}, 'EUSEBIO': {}}

    with col_m:
        st.markdown("**MATRIZ**")
        valores_digitados['MATRIZ']['NOVOS PAGOS'] = st.text_input("NOVOS PAGOS", valores_iniciais['MATRIZ']['NOVOS PAGOS'], key="m_novos_pagos")
        valores_digitados['MATRIZ']['USADOS PAGOS'] = st.text_input("USADOS PAGOS", valores_iniciais['MATRIZ']['USADOS PAGOS'], key="m_usados_pagos")
        valores_digitados['MATRIZ']['H.B.PECAS'] = st.text_input("H.B.PECAS", valores_iniciais['MATRIZ']['H.B.PECAS'], key="m_hb_pecas")
        valores_digitados['MATRIZ']['FIDIC'] = st.text_input("FIDIC", valores_iniciais['MATRIZ']['FIDIC'], key="m_fidic")
        valores_digitados['MATRIZ']['ESTOQUE PECAS'] = st.text_input("EST.PECAS", valores_iniciais['MATRIZ']['ESTOQUE PECAS'], key="m_est_pecas")

    with col_ws:
        st.markdown("**WS**")
        valores_digitados['WS']['NOVOS PAGOS'] = st.text_input("NOVOS PAGOS", valores_iniciais['WS']['NOVOS PAGOS'], key="ws_novos_pagos")
        valores_digitados['WS']['USADOS PAGOS'] = st.text_input("USADOS PAGOS", valores_iniciais['WS']['USADOS PAGOS'], key="ws_usados_pagos")
        valores_digitados['WS']['H.B.PECAS'] = st.text_input("H.B.PECAS", valores_iniciais['WS']['H.B.PECAS'], key="ws_hb_pecas")
        valores_digitados['WS']['FIDIC'] = st.text_input("FIDIC", valores_iniciais['WS']['FIDIC'], key="ws_fidic")
        valores_digitados['WS']['ESTOQUE PECAS'] = st.text_input("EST.PECAS", valores_iniciais['WS']['ESTOQUE PECAS'], key="ws_est_pecas")

    with col_e:
        st.markdown("**EUSEBIO**")
        valores_digitados['EUSEBIO']['NOVOS PAGOS'] = st.text_input("NOVOS PAGOS", valores_iniciais['EUSEBIO']['NOVOS PAGOS'], key="eus_novos_pagos")
        valores_digitados['EUSEBIO']['USADOS PAGOS'] = st.text_input("USADOS PAGOS", valores_iniciais['EUSEBIO']['USADOS PAGOS'], key="eus_usados_pagos")
        valores_digitados['EUSEBIO']['H.B.PECAS'] = st.text_input("H.B.PECAS", valores_iniciais['EUSEBIO']['H.B.PECAS'], key="eus_hb_pecas")
        valores_digitados['EUSEBIO']['FIDIC'] = st.text_input("FIDIC", valores_iniciais['EUSEBIO']['FIDIC'], key="eus_fidic")
        valores_digitados['EUSEBIO']['ESTOQUE PECAS'] = st.text_input("EST.PECAS", valores_iniciais['EUSEBIO']['ESTOQUE PECAS'], key="eus_est_pecas")

    if st.button("💾 Carregar Dados e Calcular"):
        lista_df = [carregar_posicao_analitica(), carregar_obrigacoes(), carregar_creditos_nao_identificados(), carregar_adiantamentos(), carregar_manuais(valores_digitados)]
        lista_df = [df for df in lista_df if not df.empty]
        if lista_df:
            df = pd.concat(lista_df, ignore_index=True)
            df['Saldo'] = pd.to_numeric(df['Saldo'], errors='coerce').fillna(0.0)

            empresas = df['Empresa'].unique()
            novas_linhas = []
            for emp in empresas:
                trans = df[(df['Empresa'] == emp) & (df['Tipo de Título'] == 'TRANSITORIA')]['Saldo'].sum()
                adiant = df[(df['Empresa'] == emp) & (df['Tipo de Título'] == 'ADIANTAMENTO')]['Saldo'].sum()
                if trans > 0:
                    dif = trans - adiant
                    if dif!= 0: novas_linhas.append({'Tipo de Título': 'DIF_TRANS_ADIANT', 'Empresa': emp, 'Saldo': dif})
            if novas_linhas: df = pd.concat([df, pd.DataFrame(novas_linhas)], ignore_index=True)

            st.session_state['df_final'] = df
            st.success("Dados carregados!")
        else:
            st.error("Não consegui ler os dados dos arquivos")

    # ================= TABELA + DOWNLOAD =================
    if 'df_final' in st.session_state:
        df = st.session_state['df_final']
    
        st.markdown("### POSIÇÃO FINANCEIRA DIÁRIA")
        c1, c2, c3 = st.columns([3, 1, 1])
        with c2: st.markdown("**DATA**")
        with c3: st.markdown(f"**{date.today().strftime('%d/%m/%Y')}**")
        st.divider()
    
        # AQUI: MONTA A TABELA PARA VISUALIZAR NA TELA
        empresas_ordem = ['MATRIZ', 'WS', 'EUSEBIO']
        for emp in empresas_ordem:
            if emp not in empresas_selecionadas: continue
            
            st.markdown(f"#### {emp}")
            valores_emp = {}
            total_geral = 0.0
            
            for item_chave, item_nome in ITENS:
                total = df[(df['Tipo de Título'] == item_chave) & (df['Empresa'] == emp)]['Saldo'].sum()
                
                # Recalcula DIF_TRANS_ADIANT na hora de mostrar
                if item_chave == 'DIF_TRANS_ADIANT':
                    trans_valor = df[(df['Tipo de Título'] == 'TRANSITORIA') & (df['Empresa'] == emp)]['Saldo'].sum()
                    adiant_valor = df[(df['Tipo de Título'] == 'ADIANTAMENTO') & (df['Empresa'] == emp)]['Saldo'].sum()
                    total = trans_valor - adiant_valor if trans_valor > 0 else 0.0
                
                valores_emp[item_chave] = total
                if item_chave == 'OBRIGACOES': total_geral -= total
                else: total_geral += total
                
                col1, col2 = st.columns([3,1])
                with col1: st.write(item_nome)
                with col2: st.write(formatar_br(total))
            
            st.markdown(f"**TOTAL {emp}: {formatar_br(total_geral)}**")
            st.divider()
    
        with st.sidebar:
            st.markdown("### Filtros")
            empresas_selecionadas = st.multiselect("Empresas", ['MATRIZ', 'WS', 'EUSEBIO'], default=['MATRIZ', 'WS', 'EUSEBIO'])
            st.divider()
            st.markdown("### Exportar")
            
            excel_data = gerar_excel(df, empresas_selecionadas)
            st.download_button(
                label="📥 Baixar Excel",
                data=excel_data,
                file_name=f"Posicao_Financeira_{date.today().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
