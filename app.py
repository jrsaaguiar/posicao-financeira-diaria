import pandas as pd
import streamlit as st
from datetime import date, timedelta
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter
from database import SessionLocal, PosicaoDiaria, Base # <- adiciona Base
from sqlalchemy import Column, Integer, String, Boolean # <- adiciona isso
import hashlib

class Usuarios(Base): # <- agora Base existe
    __tablename__ = 'usuarios'
    id = Column(Integer, primary_key=True)
    email = Column(String)
    senha_hash = Column(String)
    nome = Column(String)
    ativo = Column(Boolean)

# Empresas
EMPRESAS = ["MATRIZ", "WS", "EUSEBIO"]

st.set_page_config(layout="wide", page_title="Posição Financeira Diária")
def tela_login():
    st.title("🔒 Acesso Restrito - Posição Diária")
    
    with st.form("login"):
        email = st.text_input("Email")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar")
        
        if entrar:
            db = SessionLocal()
            senha_hash = hashlib.sha256(senha.encode()).hexdigest()
            user = db.query(Usuarios).filter_by(email=email, senha_hash=senha_hash, ativo=True).first()
            db.close()
            
            if user:
                st.session_state['logado'] = True
                st.session_state['usuario'] = user.nome
                st.session_state['email'] = user.email
                st.rerun()
            else:
                st.error("Email ou senha inválidos")

# TRAVA
if 'logado' not in st.session_state:
    st.session_state['logado'] = False

if not st.session_state['logado']:
    tela_login()
    st.stop() # para tudo aqui se não logou

# SIDEBAR
st.sidebar.success(f"Logado: {st.session_state['usuario']}")
if st.sidebar.button("Sair"):
    for key in ['logado', 'usuario', 'email']:
        st.session_state.pop(key, None)
    st.rerun()
st.title("Dashboard Financeira Diária")
st.markdown("""
<style>
[data-testid="stMetricValue"] {
    font-size: 20px!important;
    display: flex!important;
    align-items: baseline!important;
    gap: 3px!important;
}
[data-testid="stMetricValue"]::before {
    content: "R$ ";
    font-size: 12px!important;
    font-weight: 600;
}
[data-testid="stMetricLabel"] {
    font-size: 11px!important;
}
</style>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["Lançamento", "Histórico"])

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

@st.cache_data
def get_total_empresa(data, empresa):
    """Calcula o total líquido da empresa na data"""
    db = SessionLocal()
    regs = db.query(PosicaoDiaria).filter(
        PosicaoDiaria.data == data,
        PosicaoDiaria.empresa == empresa
    ).all()
    db.close()
    if not regs: return 0.0

    total = 0.0
    for r in regs:
        if r.tipo_titulo == 'OBRIG. A PAGA':
            total -= r.valor
        elif r.tipo_titulo not in ['TRANSITORIA', 'DIF_TRANS_ADIANT']:
            total += r.valor
    return total

@st.cache_data
def get_variacao_empresa(data_hoje, data_ontem, empresa):
    """Calcula a variação % vs dia anterior"""
    total_hoje = get_total_empresa(data_hoje, empresa)
    total_ontem = get_total_empresa(data_ontem, empresa)

    if total_ontem == 0:
        return None

    variacao = ((total_hoje - total_ontem) / abs(total_ontem)) * 100
    return variacao

def formatar_compacto(valor):
    """Formata para caber no card: 1.2M, 589K"""
    if valor >= 1_000_000:
        return f"{valor/1_000_000:.2f}M"
    elif valor >= 1_000:
        return f"{valor/1_000:.1f}K"
    else:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def normalizar_tipo(titulo):
    t = str(titulo).upper().strip()
    if 'CARTEIRA' in t: return 'CARTEIRA'
    if 'MERCADO.PAGO' in t or 'MERCADO PAGO' in t: return 'MERCADO PAGO'
    if 'VEICULO' in t or 'V-VEICULO' in t or 'VENDA VEIC' in t: return 'VEICULO'
    if 'SEGURADORA' in t or 'SEGURO' in t: return 'SEGURADORA'
    if 'GARANTIA' in t or 'VENDA GARANTIA' in t: return 'GARANTIA'
    if 'BANCO' in t: return 'BANCOS'
    if 'CARTAO' in t or 'CARTÕES' in t or 'CARTAO DE CREDITO' in t or 'CREDITO' in t: return 'CARTOES'
    if 'FUNDAO NOVOS' in t or 'FUNDAO' in t: return 'FUNDAO NOVOS'
    if 'NOVOS PAGOS' in t or 'NOVOS.PAGOS' in t: return 'NOVOS PAGOS'
    if 'USADOS PAGOS' in t or 'USADOS.PAGOS' in t or 'USADOS' in t: return 'USADOS PAGOS'
    if 'FIDIC' in t: return 'FIDIC'
    if 'H.B.PECAS' in t or 'HB PECAS' in t or 'PECAS' in t: return 'H.B.PECAS'
    if 'ESTOQUE PECAS' in t or 'EST.PECAS' in t: return 'ESTOQUE PECAS'
    if 'OBRIG' in t: return 'OBRIG. A PAGA'
    if 'ADIANTAMENTO' in t: return 'ADIANTAMENTOS'
    if 'TRANSITORIA' in t: return 'TRANSITORIA'
    if 'DIF_TRANS_ADIANT' in t: return 'DIF_TRANS_ADIANT'
    return 'OUTROS'

ITENS = [
    ('CARTEIRA', 'CARTEIRA'), ('MERCADO PAGO', 'MERCADO PAGO'),('VEICULO', 'VEICULO'), ('SEGURADORA', 'SEGURADORA'),
    ('GARANTIA', 'GARANTIA'), ('BANCOS', 'BANCOS'), ('CARTOES', 'CARTOES'),
    ('NOVOS PAGOS', 'NOVOS PAGOS'),('USADOS PAGOS', 'USADOS PAGOS'),('FUNDAO NOVOS', 'FUNDAO NOVOS'),
    ('FIDIC', 'FIDIC'), ('H.B.PECAS', 'H.B.PECAS'),
    ('ESTOQUE PECAS', 'ESTOQUE PECAS'), ('OBRIG. A PAGA', 'OBRIG. A PAGA'),('ADIANTAMENTOS', 'ADIANTAMENTOS'),
    ('TRANSITORIA', 'TRANSITORIA'), ('DIF_TRANS_ADIANT', 'DIF_TRANS_ADIANT')
]

ITENS_MANUAIS = [
    ('NOVOS PAGOS', 'NOVOS PAGOS'),
    ('USADOS PAGOS', 'USADOS PAGOS'),
    ('FUNDAO NOVOS', 'FUNDAO NOVOS'),
    ('FIDIC', 'FIDIC'), ('H.B.PECAS', 'H.B.PECAS'),
    ('ESTOQUE PECAS', 'ESTOQUE PECAS')
]
# Carrega Valores manuais do banco
def carregar_valores_manuais_do_banco(data_ref):
    db = SessionLocal()
    valores = {
        'MATRIZ': {k: '0,00' for k,_ in ITENS_MANUAIS},
        'WS': {k: '0,00' for k,_ in ITENS_MANUAIS},
        'EUSEBIO': {k: '0,00' for k,_ in ITENS_MANUAIS}
    }
    valores_qtd = {
        'MATRIZ': {'NOVOS PAGOS': 0, 'USADOS PAGOS': 0},
        'WS': {'NOVOS PAGOS': 0, 'USADOS PAGOS': 0},
        'EUSEBIO': {'NOVOS PAGOS': 0, 'USADOS PAGOS': 0}
    }
    registros = db.query(PosicaoDiaria).filter(
        PosicaoDiaria.data == data_ref,
        PosicaoDiaria.tipo_titulo.in_([k for k,_ in ITENS_MANUAIS])
    ).all()
    for reg in registros:
        valores[reg.empresa][reg.tipo_titulo] = formatar_br(reg.valor).replace('R$ ', '')
        if reg.tipo_titulo in ['NOVOS PAGOS', 'USADOS PAGOS']:
            valores_qtd[reg.empresa][reg.tipo_titulo] = int(reg.qtd_veiculos or 0) # <- GARANTE INT
    db.close()
    return valores, valores_qtd

def salvar_posicao_no_banco(df, data_ref, modo='novo'):
    db = SessionLocal()
    if modo == 'manutencao':
        for _, row in df.iterrows():
            reg = db.query(PosicaoDiaria).filter(
                PosicaoDiaria.data == data_ref,
                PosicaoDiaria.empresa == row['Empresa'],
                PosicaoDiaria.tipo_titulo == row['Tipo de Título']
            ).first()
            if reg:
                reg.valor = row['Saldo']
                reg.qtd_veiculos = row.get('Qtd', 0)
                reg.valor_medio = row.get('ValorMedio', 0.0)
            else:
                novo = PosicaoDiaria(
                    data=data_ref,
                    empresa=row['Empresa'],
                    tipo_titulo=row['Tipo de Título'],
                    valor=row['Saldo'],
                    qtd_veiculos=row.get('Qtd', 0),
                    valor_medio=row.get('ValorMedio', 0.0)
                )
                db.add(novo)
    else:
        db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_ref).delete()
        for _, row in df.iterrows():
            qtd = row.get('Qtd', 0)
            valor_medio = row.get('ValorMedio', 0.0)
            novo = PosicaoDiaria(
                data=data_ref,
                empresa=row['Empresa'],
                tipo_titulo=row['Tipo de Título'],
                valor=row['Saldo'],
                qtd_veiculos=qtd,
                valor_medio=valor_medio
            )
            db.add(novo)
    db.commit()
    db.close()

def zerar_banco():
    db = SessionLocal()
    qtd = db.query(PosicaoDiaria).count()
    db.query(PosicaoDiaria).delete()
    db.commit()
    db.close()
    return qtd
# Gerar excel exporta empresas selecionadas.
def gerar_excel(df_para_exportar, empresas_selecionadas, data_titulo_str):
    output = BytesIO()
    data_excel = date.fromisoformat(data_titulo_str).strftime('%d/%m/%Y')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        worksheet = workbook.create_sheet('POSICAO DIARIA')
        writer.sheets['POSICAO DIARIA'] = worksheet
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        border_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        bold = Font(bold=True, size=11); center = Alignment(horizontal='center', vertical='center'); right = Alignment(horizontal='right', vertical='center')
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        cell_titulo = worksheet.cell(row=1, column=1, value="POSIÇÃO FINANCEIRA DIÁRIA"); cell_titulo.font = Font(bold=True, size=14); cell_titulo.alignment = center; cell_titulo.border = border_fina
        worksheet.cell(row=1, column=7, value='DATA').font = bold; worksheet.cell(row=1, column=7).alignment = center; worksheet.cell(row=1, column=7).border = border_fina
        worksheet.cell(row=1, column=8, value=data_excel).border = border_fina

        col_inicio = 1
        for emp in ['MATRIZ', 'WS', 'EUSEBIO']:
            if emp not in empresas_selecionadas: col_inicio += 3; continue
            worksheet.merge_cells(start_row=2, start_column=col_inicio, end_row=2, end_column=col_inicio+1)
            cell_emp = worksheet.cell(row=2, column=col_inicio, value=emp); cell_emp.font = bold; cell_emp.alignment = center; cell_emp.fill = header_fill; cell_emp.border = border_fina
            worksheet.cell(row=2, column=col_inicio+1).fill = header_fill; worksheet.cell(row=2, column=col_inicio+1).border = border_fina
            col_inicio += 3

        linha_temp = 3; col_inicio = 1
        for emp in ['MATRIZ', 'WS', 'EUSEBIO']:
            if emp not in empresas_selecionadas: col_inicio += 3; continue
            worksheet.column_dimensions[get_column_letter(col_inicio)].width = 25
            worksheet.column_dimensions[get_column_letter(col_inicio+1)].width = 18
            worksheet.cell(row=linha_temp, column=col_inicio, value='DESCRICAO').font = bold; worksheet.cell(row=linha_temp, column=col_inicio).fill = header_fill; worksheet.cell(row=linha_temp, column=col_inicio).border = border_fina
            worksheet.cell(row=linha_temp, column=col_inicio+1, value='VALORES').font = bold; worksheet.cell(row=linha_temp, column=col_inicio+1).fill = header_fill; worksheet.cell(row=linha_temp, column=col_inicio+1).border = border_fina
            col_inicio += 3

        linha_dados = 4
        for item_chave, item_nome in ITENS:
            col_inicio = 1
            for emp in ['MATRIZ', 'WS', 'EUSEBIO']:
                if emp not in empresas_selecionadas: col_inicio += 3; continue
                total_valor = df_para_exportar[(df_para_exportar['Tipo de Título'] == item_chave) & (df_para_exportar['Empresa'] == emp)]['Saldo'].sum()
                total_qtd = df_para_exportar[(df_para_exportar['Tipo de Título'] == item_chave) & (df_para_exportar['Empresa'] == emp)]['Qtd'].sum()
                if item_chave == 'DIF_TRANS_ADIANT':
                    trans_valor = df_para_exportar[(df_para_exportar['Tipo de Título'] == 'TRANSITORIA') & (df_para_exportar['Empresa'] == emp)]['Saldo'].sum()
                    adiant_valor = df_para_exportar[(df_para_exportar['Tipo de Título'] == 'ADIANTAMENTOS') & (df_para_exportar['Empresa'] == emp)]['Saldo'].sum()
                    total_valor = adiant_valor - trans_valor if trans_valor > 0 else 0.0
                cell_desc = worksheet.cell(row=linha_dados, column=col_inicio, value=item_nome); cell_desc.border = border_fina
                if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS']:
                    valor_formatado = formatar_br(total_valor).replace('R$ ', '') # <- CORRIGIDO
                    cell_valor = worksheet.cell(row=linha_dados, column=col_inicio+1, value=f"{valor_formatado} - {int(total_qtd)}")
                    cell_valor.alignment = right; cell_valor.border = border_fina
                else:
                    cell_valor = worksheet.cell(row=linha_dados, column=col_inicio+1, value=total_valor); cell_valor.alignment = right; cell_valor.number_format = 'R$ #,##0.00'; cell_valor.border = border_fina
                col_inicio += 3
            linha_dados += 1

        col_inicio = 1
        for emp in ['MATRIZ', 'WS', 'EUSEBIO']:
            if emp not in empresas_selecionadas: col_inicio += 3; continue
            total_geral = 0.0
            for item_chave, item_nome in ITENS:
                total = df_para_exportar[(df_para_exportar['Tipo de Título'] == item_chave) & (df_para_exportar['Empresa'] == emp)]['Saldo'].sum()
                if item_chave == 'DIF_TRANS_ADIANT':
                    trans_valor = df_para_exportar[(df_para_exportar['Tipo de Título'] == 'TRANSITORIA') & (df_para_exportar['Empresa'] == emp)]['Saldo'].sum()
                    adiant_valor = df_para_exportar[(df_para_exportar['Tipo de Título'] == 'ADIANTAMENTOS') & (df_para_exportar['Empresa'] == emp)]['Saldo'].sum()
                    total = adiant_valor - trans_valor if trans_valor > 0 else 0.0
                if item_chave == 'OBRIG. A PAGA':
                    total_geral -= total
                elif item_chave not in ['TRANSITORIA', 'DIF_TRANS_ADIANT']:
                    total_geral += total
            worksheet.cell(row=linha_dados, column=col_inicio, value='TOTAL').font = bold; worksheet.cell(row=linha_dados, column=col_inicio).border = border_fina
            cell_total = worksheet.cell(row=linha_dados, column=col_inicio+1, value=total_geral); cell_total.font = bold; cell_total.alignment = right; cell_total.number_format = 'R$ #,##0.00'; cell_total.border = border_fina
            col_inicio += 3
    return output.getvalue()

with tab1:
    with st.sidebar:
        st.markdown("### Data de Referência do Lançamento")
        DATA_REF_DATE = st.date_input("Selecione a Data", value=date.today(), format="DD/MM/YYYY", key="data_lancamento")
        DATA_REF = DATA_REF_DATE.strftime("%Y-%m-%d")

        st.markdown("### Filtros")
        empresas_selecionadas = st.multiselect("Empresas", ['MATRIZ', 'WS', 'EUSEBIO'], default=['MATRIZ', 'WS', 'EUSEBIO'])
        st.divider()
        st.markdown("### Exportar")
        st.markdown("### Resumo do Dia")
        col_m, col_ws, col_e = st.columns(3)
        empresas_cards = {'MATRIZ': col_m, 'WS': col_ws, 'EUSEBIO': col_e}

        for emp, col in empresas_cards.items():
            if emp not in empresas_selecionadas: continue
            with col:
                total_hoje = get_total_empresa(DATA_REF, emp)
                data_ontem_date = DATA_REF_DATE - timedelta(days=1)
                data_ontem = data_ontem_date.strftime("%Y-%m-%d")
                variacao = get_variacao_empresa(DATA_REF, data_ontem, emp)

                st.metric(
                    label=f"TOTAL {emp}",
                    value=formatar_compacto(total_hoje),
                    delta=f"{variacao:.2f}%" if variacao is not None else None
                )
        st.divider()

    with st.expander("📝 Lançamento Manual / Manutenção - Editável", expanded=True):
        col_cal, col_btn = st.columns([3,1])
        with col_cal:
            DATA_MANUTENCAO_DATE = st.date_input("📅 Selecione a Data para Manutenção", value=DATA_REF_DATE, format="DD/MM/YYYY", key="data_manutencao")
            DATA_MANUTENCAO = DATA_MANUTENCAO_DATE.strftime("%Y-%m-%d")
        with col_btn:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📂 Carregar Dados da Data", key="btn_carregar_manut"):
                db = SessionLocal()
                registros = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == DATA_MANUTENCAO).all()
                db.close()
                if registros:
                    dados = [{'Tipo de Título': r.tipo_titulo, 'Empresa': r.empresa, 'Saldo': r.valor, 'Qtd': r.qtd_veiculos, 'ValorMedio': r.valor_medio} for r in registros]
                    st.session_state['df_carregado_manut'] = pd.DataFrame(dados)
                    st.success(f"Dados de {DATA_MANUTENCAO_DATE.strftime('%d/%m/%Y')} carregados.")
                else:
                    st.warning("Nenhum dado encontrado para esta data.")
                    st.session_state['df_carregado_manut'] = pd.DataFrame()
                st.rerun()

        valores_iniciais, valores_qtd_iniciais = carregar_valores_manuais_do_banco(DATA_MANUTENCAO)

        if 'df_carregado_manut' in st.session_state and not st.session_state['df_carregado_manut'].empty:
            df_temp = st.session_state['df_carregado_manut']
            for emp in ['MATRIZ', 'WS', 'EUSEBIO']:
                for item_chave, _ in ITENS_MANUAIS:
                    val = df_temp[(df_temp['Empresa']==emp) & (df_temp['Tipo de Título']==item_chave)]['Saldo'].sum()
                    qtd = df_temp[(df_temp['Empresa']==emp) & (df_temp['Tipo de Título']==item_chave)]['Qtd'].sum()
                    valores_iniciais[emp][item_chave] = formatar_br(val).replace('R$ ', '')
                    if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS']:
                        valores_qtd_iniciais[emp][item_chave] = int(qtd)

        uploaded_files = st.file_uploader("📁 Arraste os 4 arquivos RFN aqui - Só se for lançamento novo", type=['xlsx', 'xls'], accept_multiple_files=True, key="uploader_tab1")
        dfs = {}
        if uploaded_files:
            dfs = {file.name: pd.read_excel(file) for file in uploaded_files}

        def carregar_posicao_analitica():
            if not uploaded_files: return pd.DataFrame()
            if 'RFN003_PosicaoAnaliticoReceber_Excel.xls' not in dfs and 'RFN003_PosicaoAnaliticoReceber_Excel.xlsx' not in dfs: return pd.DataFrame()
            df_raw = dfs.get('RFN003_PosicaoAnaliticoReceber_Excel.xls', dfs.get('RFN003_PosicaoAnaliticoReceber_Excel.xlsx'))
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
            if not uploaded_files: return pd.DataFrame()
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
            linhas = [{'Tipo de Título': 'OBRIG. A PAGA', 'Empresa': e, 'Saldo': v} for e,v in obrig_dict.items() if v > 0]
            return pd.DataFrame(linhas)

        def carregar_creditos_nao_identificados():
            if not uploaded_files: return pd.DataFrame()
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
            if not uploaded_files: return pd.DataFrame()
            if 'RFN013_FichaRazaoSaldoExcel.xlsx' not in dfs and 'RFN013_FichaRazaoSaldo_Excel.xls' not in dfs: return pd.DataFrame()
            df_ad_raw = dfs.get('RFN013_FichaRazaoSaldoExcel.xlsx', dfs.get('RFN013_FichaRazaoSaldo_Excel.xls'))
            adiant_dict = {'MATRIZ': 0.0, 'WS': 0.0, 'EUSEBIO': 0.0}
            for i in range(1, len(df_ad_raw)):
                emp = detectar_empresa(str(df_ad_raw.iloc[i, 3]))
                sal = converter_valor_br(str(df_ad_raw.iloc[i, 6]))
                if emp in adiant_dict: adiant_dict[emp] += sal
            linhas = [{'Tipo de Título': 'ADIANTAMENTOS', 'Empresa': e, 'Saldo': v} for e,v in adiant_dict.items() if v > 0]
            return pd.DataFrame(linhas)
        
        # def carregar_manuais_adiantamentos()
        def carregar_manuais(valores_digitados, valores_qtd_digitados):
            dados = []
            for empresa, itens in valores_digitados.items():
                for tipo, valor_str in itens.items():
                    valor = converter_valor_br(valor_str)
                    qtd = valores_qtd_digitados[empresa].get(tipo, 0)
                    valor_medio = valor / qtd if qtd > 0 else 0.0
                    dados.append({'Tipo de Título': tipo, 'Empresa': empresa, 'Saldo': valor, 'Qtd': qtd, 'ValorMedio': valor_medio})
            return pd.DataFrame(dados)
        
        col_m, col_ws, col_e = st.columns(3)
        valores_digitados = {'MATRIZ': {}, 'WS': {}, 'EUSEBIO': {}}
        valores_qtd_digitados = {'MATRIZ': {}, 'WS': {}, 'EUSEBIO': {}}
        empresas_col = {'MATRIZ': col_m, 'WS': col_ws, 'EUSEBIO': col_e}
        
        for emp, col in empresas_col.items():
            with col:
                st.markdown(f"**{emp}**")
                for item_chave, item_nome in ITENS_MANUAIS:
                    st.markdown(f"{item_nome}")
                    key_valor = f"{emp}_{item_chave}_edit_{DATA_MANUTENCAO.replace('-', '')}"
                    key_qtd = f"{emp}_QTD_{item_chave}_edit_{DATA_MANUTENCAO.replace('-', '')}"
        
                    valores_digitados[emp][item_chave] = st.text_input(
                        label="",
                        value=valores_iniciais[emp][item_chave],
                        key=key_valor,
                        label_visibility="collapsed"
                    )
                    if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS']:
                        valores_qtd_digitados[emp][item_chave] = st.number_input(
                            "Qtd",
                            value=int(valores_qtd_iniciais[emp][item_chave]), # <- INT
                            key=key_qtd,
                            min_value=0,
                            step=1 # <- STEP
                        )
        
        if st.button(f"💾 Salvar / Atualizar {DATA_MANUTENCAO_DATE.strftime('%d/%m/%Y')}", key="btn_salvar_manut"):
            if uploaded_files:
                lista_df = [carregar_posicao_analitica(), carregar_obrigacoes(), carregar_creditos_nao_identificados(), carregar_adiantamentos(), carregar_manuais(valores_digitados, valores_qtd_digitados)]
                lista_df = [df for df in lista_df if not df.empty]
                if lista_df:
                    df = pd.concat(lista_df, ignore_index=True)
                    df['Saldo'] = pd.to_numeric(df['Saldo'], errors='coerce').fillna(0.0)
                    df['Qtd'] = pd.to_numeric(df['Qtd'], errors='coerce').fillna(0)
                    empresas = df['Empresa'].unique()
                    novas_linhas = []
                    for emp in empresas:
                        trans = df[(df['Empresa'] == emp) & (df['Tipo de Título'] == 'TRANSITORIA')]['Saldo'].sum()
                        adiant = df[(df['Empresa'] == emp) & (df['Tipo de Título'] == 'ADIANTAMENTOS')]['Saldo'].sum()
                        if trans > 0:
                            dif = adiant - trans
                            if dif!= 0: novas_linhas.append({'Tipo de Título': 'DIF_TRANS_ADIANT', 'Empresa': emp, 'Saldo': dif, 'Qtd': 0, 'ValorMedio': 0.0})
                    if novas_linhas: df = pd.concat([df, pd.DataFrame(novas_linhas)], ignore_index=True)
                    salvar_posicao_no_banco(df, DATA_MANUTENCAO, modo='novo')
                    st.session_state['df_final'] = df
                    st.success(f"Dados de {DATA_MANUTENCAO_DATE.strftime('%d/%m/%Y')} ATUALIZADOS no banco!")
            else:
                df_manual = carregar_manuais(valores_digitados, valores_qtd_digitados)
                salvar_posicao_no_banco(df_manual, DATA_MANUTENCAO, modo='manutencao')
                if 'df_carregado_manut' in st.session_state:
                    del st.session_state['df_carregado_manut']
                st.success(f"Manutenção salva! Data: {DATA_MANUTENCAO_DATE.strftime('%d/%m/%Y')}")
            st.rerun()
        
        if 'df_final' in st.session_state:
            df = st.session_state['df_final']
            with st.sidebar:
                excel_data = gerar_excel(df, empresas_selecionadas, DATA_REF)
                st.download_button(label="📊 Gerar Planilha Única", data=excel_data, file_name=f"Posicao_Financeira_{DATA_REF.replace('-', '')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            col1, col2, col3 = st.columns(3)
            empresas_cols = {'MATRIZ': col1, 'WS': col2, 'EUSEBIO': col3}
            for emp, col in empresas_cols.items():
                if emp not in empresas_selecionadas: continue
                with col:
                    dados_tabela = []
                    total_geral = 0.0
                    for item_chave, item_nome in ITENS:
                        total = df[(df['Tipo de Título'] == item_chave) & (df['Empresa'] == emp)]['Saldo'].sum()
                        total_qtd = df[(df['Tipo de Título'] == item_chave) & (df['Empresa'] == emp)]['Qtd'].sum()
                        if item_chave == 'DIF_TRANS_ADIANT':
                            trans_valor = df[(df['Tipo de Título'] == 'TRANSITORIA') & (df['Empresa'] == emp)]['Saldo'].sum()
                            adiant_valor = df[(df['Tipo de Título'] == 'ADIANTAMENTOS') & (df['Empresa'] == emp)]['Saldo'].sum()
                            total = adiant_valor - trans_valor if trans_valor > 0 else 0.0
                        if item_chave == 'OBRIG. A PAGA':
                            total_geral -= total
                        elif item_chave not in ['TRANSITORIA', 'DIF_TRANS_ADIANT']:
                            total_geral += total
                        if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS']:
                            valor_formatado = formatar_br(total).replace('R$ ', '')
                            dados_tabela.append({"DESCRICAO": item_nome, "VALORES": f"{valor_formatado} - {int(total_qtd)}"})
                        else:
                            dados_tabela.append({"DESCRICAO": item_nome, "VALORES": formatar_br(total)})
                    dados_tabela.append({"DESCRICAO": "TOTAL", "VALORES": formatar_br(total_geral)})
                    df_mostrar = pd.DataFrame(dados_tabela)
                    st.dataframe(df_mostrar, hide_index=True, use_container_width=True, height=680, column_config={
                        "DESCRICAO": st.column_config.TextColumn("DESCRICAO"),
                        "VALORES": st.column_config.TextColumn("VALORES")
                    })
        
        with tab2:
            st.header("Consulta de Posições Salvas")
            with st.expander("⚠️ Zona Perigosa - Apagar Dados"):
                st.warning("Isso vai apagar TODAS as datas do banco. Use só pra zerar e recomeçar.")
                if st.button("🗑️ ZERAR BANCO COMPLETO"):
                    qtd_apagada = zerar_banco()
                    st.success(f"Banco zerado! {qtd_apagada} registros apagados.")
                    st.rerun()
        
            col1, col2 = st.columns([2,3])
            with col1:
                data_consulta_date = st.date_input("Selecione a Data", value=date.today(), format="DD/MM/YYYY")
                data_consulta = data_consulta_date.strftime("%Y-%m-%d")
        
            if st.button("Carregar Dados da Data"):
                db = SessionLocal()
                registros = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_consulta).all()
                db.close()
                if not registros:
                    st.warning(f"Nenhum dado encontrado para {data_consulta_date.strftime('%d/%m/%Y')}")
                else:
                    dados = [{'Tipo de Título': r.tipo_titulo, 'Empresa': r.empresa, 'Saldo': r.valor, 'Qtd': r.qtd_veiculos, 'Valor Medio': r.valor_medio} for r in registros]
                    df_hist = pd.DataFrame(dados)
                    st.success(f"{len(df_hist)} registros encontrados para {data_consulta_date.strftime('%d/%m/%Y')}")
                    col1, col2, col3 = st.columns(3)
                    empresas_cols = {'MATRIZ': col1, 'WS': col2, 'EUSEBIO': col3}
                    for emp, col in empresas_cols.items():
                        with col:
                            dados_tabela = []
                            total_geral = 0.0
                            for item_chave, item_nome in ITENS:
                                total = df_hist[(df_hist['Tipo de Título'] == item_chave) & (df_hist['Empresa'] == emp)]['Saldo'].sum()
                                total_qtd = df_hist[(df_hist['Tipo de Título'] == item_chave) & (df_hist['Empresa'] == emp)]['Qtd'].sum()
                                if item_chave == 'DIF_TRANS_ADIANT':
                                    trans_valor = df_hist[(df_hist['Tipo de Título'] == 'TRANSITORIA') & (df_hist['Empresa'] == emp)]['Saldo'].sum()
                                    adiant_valor = df_hist[(df_hist['Tipo de Título'] == 'ADIANTAMENTOS') & (df_hist['Empresa'] == emp)]['Saldo'].sum()
                                    total = adiant_valor - trans_valor if trans_valor > 0 else 0.0
                                if item_chave == 'OBRIG. A PAGA':
                                    total_geral -= total
                                elif item_chave not in ['TRANSITORIA', 'DIF_TRANS_ADIANT']:
                                    total_geral += total
                                if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS']:
                                    valor_formatado = formatar_br(total).replace('R$ ', '')
                                    dados_tabela.append({"DESCRICAO": item_nome, "VALORES": f"{valor_formatado} - {int(total_qtd)}"})
                                else:
                                    dados_tabela.append({"DESCRICAO": item_nome, "VALORES": formatar_br(total)})
                            dados_tabela.append({"DESCRICAO": "TOTAL", "VALORES": formatar_br(total_geral)})
                            df_mostrar = pd.DataFrame(dados_tabela)
                            st.dataframe(df_mostrar, hide_index=True, use_container_width=True, height=680, column_config={
                                "DESCRICAO": st.column_config.TextColumn("DESCRICAO"),
                                "VALORES": st.column_config.TextColumn("VALORES")
                            })
                    excel_data_hist = gerar_excel(df_hist, ['MATRIZ', 'WS', 'EUSEBIO'], data_consulta)
                    st.download_button(
                        label=f"📊 Baixar Planilha de {data_consulta_date.strftime('%d/%m/%Y')}",
                        data=excel_data_hist,
                file_name=f"Posicao_Financeira_{data_consulta.replace('-', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
