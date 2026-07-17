import pandas as pd
import streamlit as st
from datetime import date, timedelta
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter
from database import SessionLocal, PosicaoDiaria, Usuarios
import hashlib
import random
import string
import numpy as np
from sqlalchemy import func
from auth import verificar_login, tela_cadastro_usuario

def gerar_hash(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def gerar_senha_aleatoria(tamanho=8):
    return ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(tamanho))

def formatar_br(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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
                st.session_state['perfil'] = user.perfil
                st.rerun()
            else:
                st.error("Email ou senha inválidos")

if 'logado' not in st.session_state:
    st.session_state['logado'] = False

if not st.session_state['logado']:
    tela_login()
    st.stop()

if st.session_state['logado']:
    db = SessionLocal()
    user_db = db.query(Usuarios).filter_by(email=st.session_state['email']).first()
    st.session_state['perfil'] = user_db.perfil if user_db else "Usuario"
    db.close()

    st.sidebar.write(f"Perfil: {st.session_state['perfil']}")

    if st.session_state['perfil'] == 'Admin':
        with st.sidebar.expander("👥 Gerenciar Usuários"):
            st.markdown("#### Usuários Cadastrados")
            db = SessionLocal()
            users = db.query(Usuarios).order_by(Usuarios.id.desc()).all()
            db.close()
            selected_id = None
            if users:
                df_users = pd.DataFrame([{'ID': u.id, 'Nome': u.nome, 'Email': u.email, 'Perfil': u.perfil, 'Ativo': 'Sim' if u.ativo else 'Não'} for u in users])
                selected = st.dataframe(df_users, use_container_width=True, hide_index=True, height=200, on_select="rerun", selection_mode="single-row")
                if selected['selection']['rows']:
                    idx = selected['selection']['rows'][0]
                    selected_id = int(df_users.iloc[idx]['ID'])
            else:
                st.warning("Nenhum usuário cadastrado")
            st.divider()
            st.markdown("#### Dados do Usuário")
            user_edit = None
            if selected_id:
                db = SessionLocal()
                user_edit = db.query(Usuarios).filter(Usuarios.id == selected_id).first()
                db.close()
            with st.form("form_usuario"):
                nome = st.text_input("Nome", value=user_edit.nome if user_edit else "")
                email = st.text_input("Email", value=user_edit.email if user_edit else "", disabled=bool(user_edit))
                perfil = st.selectbox("Perfil", ["Usuario", "Admin"], index=0 if not user_edit or user_edit.perfil=="Usuario" else 1)
                ativo = st.checkbox("Ativo", value=user_edit.ativo if user_edit else True)
                col1, col2 = st.columns(2)
                btn_salvar = col1.form_submit_button("💾 Salvar", use_container_width=True)
                btn_senha = col2.form_submit_button("🔑 Nova Senha", use_container_width=True, disabled=True)
            db = SessionLocal()
            if btn_salvar:
                if user_edit:
                    user_edit.nome, user_edit.perfil, user_edit.ativo = nome, perfil, ativo
                    msg = "Usuário atualizado!"
                else:
                    if db.query(Usuarios).filter_by(email=email).first():
                        st.error("Erro: Este email já está cadastrado")
                        db.close()
                        st.stop()
                    senha_temp = gerar_senha_aleatoria()
                    novo = Usuarios(nome=nome, email=email, senha_hash=gerar_hash(senha_temp), perfil=perfil, ativo=ativo)
                    db.add(novo)
                    msg = f"Usuário cadastrado! Senha padrão: `{senha_temp}`"
                db.commit()
                st.success(msg)
                db.close()
                st.rerun()
            else:
                db.close()

    st.sidebar.divider()
    st.sidebar.markdown("### Data de Referência")
    DATA_REF_DATE = st.sidebar.date_input("Selecione a Data", value=date.today(), format="DD/MM/YYYY", key="data_lancamento")
    st.sidebar.markdown("### Filtros")
    empresas_selecionadas = st.sidebar.multiselect("Empresas", ['MATRIZ', 'WS', 'EUSEBIO'], default=['MATRIZ', 'WS', 'EUSEBIO'])
    st.sidebar.divider()
    st.sidebar.markdown("### Resumo do Dia")

    col_m, col_ws, col_e = st.sidebar.columns(3)
    empresas_cards = {'MATRIZ': col_m, 'WS': col_ws, 'EUSEBIO': col_e}

    @st.cache_data(ttl=10)
    def get_total_empresa(data, empresa):
        db = SessionLocal()
        try:
            total = db.query(func.sum(PosicaoDiaria.valor)).filter(PosicaoDiaria.data == data, PosicaoDiaria.empresa == empresa).scalar()
            return float(total or 0.0)
        finally:
            db.close()

    @st.cache_data(ttl=10)
    def get_variacao_empresa(data_atual, data_anterior, empresa):
        if not data_anterior: return 0.0
        total_atual = get_total_empresa(data_atual, empresa)
        total_anterior = get_total_empresa(data_anterior, empresa)
        if total_anterior == 0: return 0.0
        return ((total_atual - total_anterior) / total_anterior) * 100

    for emp, col in empresas_cards.items():
        if emp not in empresas_selecionadas: continue
        with col:
            total_hoje = get_total_empresa(DATA_REF_DATE, emp)
            data_ontem_date = DATA_REF_DATE - timedelta(days=1)
            variacao = get_variacao_empresa(DATA_REF_DATE, data_ontem_date, emp)
            delta_cor = "normal" if variacao >= 0 else "inverse"
            st.metric(label=f"TOTAL {emp}", value=formatar_br(total_hoje), delta=f"{variacao:.2f}%" if variacao is not None else None, delta_color=delta_cor)

st.title("Dashboard Financeira Diária")
st.markdown("""<style>[data-testid="stMetricValue"] {font-size: 20px!important;display: flex!important;align-items: baseline!important;gap: 3px!important;}[data-testid="stMetricValue"]::before {content: "R$ ";font-size: 12px!important;font-weight: 600;}[data-testid="stMetricLabel"] {font-size: 11px!important;}</style>""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["Lançamento", "Histórico"])

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

ITENS = [('CARTEIRA', 'CARTEIRA'), ('MERCADO PAGO', 'MERCADO PAGO'),('VEICULO', 'VEICULO'), ('SEGURADORA', 'SEGURADORA'),('GARANTIA', 'GARANTIA'), ('BANCOS', 'BANCOS'), ('CARTOES', 'CARTOES'),('NOVOS PAGOS', 'NOVOS PAGOS'),('USADOS PAGOS', 'USADOS PAGOS'),('FUNDAO NOVOS', 'FUNDAO NOVOS'),('FIDIC', 'FIDIC'), ('H.B.PECAS', 'H.B.PECAS'),('ESTOQUE PECAS', 'ESTOQUE PECAS'), ('OBRIG. A PAGA', 'OBRIG. A PAGA'),('ADIANTAMENTOS', 'ADIANTAMENTOS'),('TRANSITORIA', 'TRANSITORIA'), ('DIF_TRANS_ADIANT', 'DIF_TRANS_ADIANT')]
ITENS_MANUAIS = [('NOVOS PAGOS', 'NOVOS PAGOS'),('USADOS PAGOS', 'USADOS PAGOS'),('FUNDAO NOVOS', 'FUNDAO NOVOS'),('FIDIC', 'FIDIC'), ('H.B.PECAS', 'H.B.PECAS'),('ESTOQUE PECAS', 'ESTOQUE PECAS')]

def carregar_valores_manuais_do_banco(data_ref):
    if isinstance(data_ref, str): data_ref = date.fromisoformat(data_ref)
    db = SessionLocal()
    valores = {'MATRIZ': {k: '0,00' for k,_ in ITENS_MANUAIS}, 'WS': {k: '0,00' for k,_ in ITENS_MANUAIS}, 'EUSEBIO': {k: '0,00' for k,_ in ITENS_MANUAIS}}
    valores_qtd = {'MATRIZ': {'NOVOS PAGOS': 0, 'USADOS PAGOS': 0}, 'WS': {'NOVOS PAGOS': 0, 'USADOS PAGOS': 0}, 'EUSEBIO': {'NOVOS PAGOS': 0, 'USADOS PAGOS': 0}}
    registros = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_ref, PosicaoDiaria.tipo_titulo.in_([k for k,_ in ITENS_MANUAIS])).all()
    for reg in registros:
        valores[reg.empresa][reg.tipo_titulo] = formatar_br(reg.valor).replace('R$ ', '')
        if reg.tipo_titulo in ['NOVOS PAGOS', 'USADOS PAGOS']:
            valores_qtd[reg.empresa][reg.tipo_titulo] = int(reg.qtd_veiculos or 0)
    db.close()
    return valores, valores_qtd

def salvar_posicao_no_banco(df, data_ref, modo='novo'):
    db = SessionLocal()
    usuario_logado = st.session_state.get('email', 'sistema')
    df = df.copy()
    if 'Qtd' not in df.columns: df['Qtd'] = 0
    if 'ValorMedio' not in df.columns: df['ValorMedio'] = 0.0
    df['Saldo'] = pd.to_numeric(df['Saldo'], errors='coerce').fillna(0.0)
    df['Qtd'] = pd.to_numeric(df['Qtd'], errors='coerce').fillna(0).astype(int)
    df['ValorMedio'] = pd.to_numeric(df['ValorMedio'], errors='coerce').fillna(0.0).replace([np.inf, -np.inf], 0.0)
    df['Empresa'] = df['Empresa'].astype(str).str.strip()
    df['Tipo de Título'] = df['Tipo de Título'].astype(str).str.strip()
    df = df[(df['Empresa']!= '') & (df['Empresa']!= 'nan') & (df['Empresa']!= 'None')]
    df = df[(df['Tipo de Título']!= '') & (df['Tipo de Título']!= 'nan') & (df['Tipo de Título']!= 'None')]
    if isinstance(data_ref, str): data_ref = date.fromisoformat(data_ref)
    if df.empty: st.warning("Nenhuma linha válida para salvar"); db.close(); return
    try:
        if modo == 'manutencao':
            for _, row in df.iterrows():
                reg = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_ref, PosicaoDiaria.empresa == row['Empresa'], PosicaoDiaria.tipo_titulo == row['Tipo de Título']).first()
                if reg:
                    reg.valor, reg.qtd_veiculos, reg.valor_medio, reg.criado_por = float(row['Saldo']), int(row['Qtd']), float(row['ValorMedio']), usuario_logado
                else:
                    db.add(PosicaoDiaria(data=data_ref, empresa=row['Empresa'], tipo_titulo=row['Tipo de Título'], valor=float(row['Saldo']), qtd_veiculos=int(row['Qtd']), valor_medio=float(row['ValorMedio']), criado_por=usuario_logado))
        else:
            db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_ref).delete()
            objs = [PosicaoDiaria(data=data_ref, empresa=row['Empresa'], tipo_titulo=row['Tipo de Título'], valor=float(row['Saldo']), qtd_veiculos=int(row['Qtd']), valor_medio=float(row['ValorMedio']), criado_por=usuario_logado) for _, row in df.iterrows()]
            db.add_all(objs)
        db.commit()
        st.cache_data.clear()
        st.success(f"{len(df)} registros salvos com sucesso!")
    except Exception as e:
        db.rollback(); st.error(f"Erro ao salvar: {e}"); st.dataframe(df)
    finally:
        db.close()

def zerar_banco():
    db = SessionLocal(); qtd = db.query(PosicaoDiaria).count(); db.query(PosicaoDiaria).delete(); db.commit(); db.close(); return qtd

def gerar_excel(df_para_exportar, empresas_selecionadas, data_titulo_str):
    return BytesIO().getvalue()

# ========== ABA 1: LANÇAMENTO ==========
with tab1:
    st.markdown("### Lançamento e Manutenção")
    col_up, col_manut = st.columns([1,2])
    with col_up:
        st.markdown("#### 1. Upload RFN")
        uploaded_files = st.file_uploader("Arraste os 4 arquivos RFN aqui", type=['xlsx', 'xls'], accept_multiple_files=True, key="uploader_tab1")
        dfs = {file.name: pd.read_excel(file) for file in uploaded_files} if uploaded_files else {}
    with col_manut:
        st.markdown("#### 2. Lançamento Manual / Manutenção")
        col_cal, col_btn = st.columns([3,1])
        with col_cal:
            DATA_MANUTENCAO_DATE = st.date_input("📅 Selecione a Data para Manutenção", value=DATA_REF_DATE, format="DD/MM/YYYY", key="data_manutencao")
            DATA_MANUTENCAO = DATA_MANUTENCAO_DATE.strftime("%Y-%m-%d")
        with col_btn:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📂 Carregar Dados da Data", key="btn_carregar_manut"):
                db = SessionLocal()
                registros = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == DATA_MANUTENCAO_DATE).all()
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
                valores_digitados[emp][item_chave] = st.text_input(label="", value=valores_iniciais[emp][item_chave], key=key_valor, label_visibility="collapsed")
                if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS']:
                    valores_qtd_digitados[emp][item_chave] = st.number_input("Qtd", value=int(valores_qtd_iniciais[emp][item_chave]), key=key_qtd, min_value=0, step=1)

    if st.button(f"💾 Salvar / Atualizar {DATA_MANUTENCAO_DATE.strftime('%d/%m/%Y')}", key="btn_salvar_manut"):

        # 1. Pega os valores digitados na tela e monta o DF
        linhas = []
        for emp in ['MATRIZ', 'WS', 'EUSEBIO']:
            for item_chave, item_nome in ITENS_MANUAIS:
                valor_str = valores_digitados[emp][item_chave]
                valor_float = converter_valor_br(valor_str)
                qtd = valores_qtd_digitados[emp].get(item_chave, 0) if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS'] else 0
                valor_medio = valor_float / qtd if qtd > 0 else 0.0
    
                linhas.append({
                    'Empresa': emp,
                    'Tipo de Título': item_chave,
                    'Saldo': valor_float,
                    'Qtd': qtd,
                    'ValorMedio': valor_medio
                })
    
        df_para_salvar = pd.DataFrame(linhas)
    
        # 2. Chama a função que já salva no banco
        salvar_posicao_no_banco(df_para_salvar, DATA_MANUTENCAO_DATE, modo='manutencao')
        st.rerun()

# ========== ABA 2: HISTÓRICO ==========
with tab2:
    st.header("📊 Relatórios e Auditoria")
    with st.expander("⚠️ Zona Perigosa - Apagar Dados"):
        if st.button("🗑️ ZERAR BANCO COMPLETO"):
            qtd_apagada = zerar_banco()
            st.success(f"Banco zerado! {qtd_apagada} registros apagados.")
            st.rerun()

    st.divider()
    st.subheader("Gerar Relatório de Auditoria por Período")

    col1, col2, col3 = st.columns(3)
    with col1:
        data_inicio = st.date_input("Data Inicial", value=date.today() - timedelta(days=7))
    with col2:
        data_fim = st.date_input("Data Final", value=date.today())
    with col3:
        db = SessionLocal()
        usuarios = [u[0] for u in db.query(Usuarios.email).filter(Usuarios.ativo==True).all()]
        db.close()
        usuario_filtro = st.selectbox("Filtrar por Usuário", ["Todos"] + usuarios)

    if st.button("📄 Gerar Relatório PDF", type="primary"):
        db = SessionLocal()
        query = db.query(PosicaoDiaria).filter(
            PosicaoDiaria.data >= data_inicio,
            PosicaoDiaria.data <= data_fim
        )
        if usuario_filtro!= "Todos":
            query = query.filter(PosicaoDiaria.criado_por == usuario_filtro)
        registros = query.all()
        db.close()

        if not registros:
            st.warning("Nenhum dado encontrado no período.")
        else:
            df_rel = pd.DataFrame([{
                'Data': r.data, 'Empresa': r.empresa, 'Item': r.tipo_titulo,
                'Valor': r.valor, 'Qtd': r.qtd_veiculos, 'Lançado por': r.criado_por
            } for r in registros])

            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import cm
            from reportlab.platypus import Table, TableStyle
            from reportlab.lib import colors

            buffer = BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            width, height = A4

            c.setFont("Helvetica-Bold", 16)
            c.drawString(2*cm, height - 2*cm, "RELATÓRIO DE AUDITORIA - POSIÇÃO FINANCEIRA")
            c.setFont("Helvetica", 10)
            c.drawString(2*cm, height - 2.8*cm, f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
            c.drawString(2*cm, height - 3.3*cm, f"Usuário Filtro: {usuario_filtro}")

            resumo = df_rel.groupby('Empresa')['Valor'].sum().reset_index()
            data_resumo = [['Empresa', 'Total Lançado']] + [[row['Empresa'], formatar_br(row['Valor'])] for _, row in resumo.iterrows()]
            t_resumo = Table(data_resumo, colWidths=[6*cm, 4*cm])
            t_resumo.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.grey), ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),('ALIGN',(0,0),(-1,-1),'CENTER'), ('GRID', (0,0), (-1,-1), 0.5, colors.black)]))
            t_resumo.wrapOn(c, width, height)
            t_resumo.drawOn(c, 2*cm, height - 6*cm)

            c.showPage()
            c.save()
            pdf_data = buffer.getvalue()

            st.success(f"{len(df_rel)} registros encontrados.")
            st.dataframe(df_rel, use_container_width=True, hide_index=True)

            st.download_button(
                label="⬇️ Baixar Relatório PDF",
                data=pdf_data,
                file_name=f"Relatorio_Auditoria_{data_inicio}_{data_fim}.pdf",
                mime="application/pdf"
            )
