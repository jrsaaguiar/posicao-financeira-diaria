import pandas as pd
import streamlit as st
from datetime import date, timedelta
import numpy as np
from sqlalchemy import func
import hashlib
import random
import string
import plotly.express as px

# Juntar tudo do database.py
from database import Usuario, SessionLocal, engine, PosicaoDiaria, init_db 

from auth import verificar_login, tela_cadastro_usuario
from exportar import gerar_excel_dashboard
from processamento_rfn import carregar_posicao_analitica, carregar_obrigacoes, carregar_creditos_nao_identificados, carregar_adiantamentos
from utils import converter_valor_br, detectar_empresa, normalizar_tipo
st.set_page_config(page_title="Posição Financeira Diária", layout="wide") # <-- deixa só 1

#init_db() # <-- cria as tabelas na primeira vez

def gerar_hash(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def gerar_senha_aleatoria(tamanho=8):
    return ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(tamanho))

def formatar_br(valor):
    if valor is None or pd.isna(valor):
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

EMPRESAS = ["MATRIZ", "WS", "EUSEBIO"]
ITENS_ORDEM = [
    'CARTEIRA', 'MERCADO PAGO', 'VEICULO', 'SEGURADORA', 'GARANTIA', 'BANCOS', 'CARTOES',
    'NOVOS PAGOS', 'USADOS PAGOS', 'FUNDAO NOVOS', 'FIDIC', 'H.B.PECAS', 'ESTOQUE PECAS',
    'OBRIG. A PAGA', 'ADIANTAMENTOS', 'TRANSITORIA', 'DIF_TRANS_ADIANT'
]

st.set_page_config(layout="wide", page_title="Posição Financeira Diária")

@st.cache_data(ttl=10)
def montar_df_dashboard(data_ref, empresas):
    db = SessionLocal()
    dados = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_ref, PosicaoDiaria.empresa.in_(empresas)).all()
    db.close()
    if not dados:
        return None

    df_pivot = pd.DataFrame(0.0, index=ITENS_ORDEM, columns=empresas)
    df_qtd = pd.DataFrame(0, index=ITENS_ORDEM, columns=empresas)
    
    for d in dados:
        if d.tipo_titulo in ITENS_ORDEM and d.empresa in empresas:
            df_pivot.loc[d.tipo_titulo, d.empresa] = d.valor or 0.0
            df_qtd.loc[d.tipo_titulo, d.empresa] = d.qtd_veiculos or 0

    dados_tabela = []
    for item in ITENS_ORDEM:
        linha = {"DESCRICAO": item}
        for emp in empresas:
            valor = df_pivot.loc[item, emp]
            qtd = df_qtd.loc[item, emp]
            if item in ['NOVOS PAGOS', 'USADOS PAGOS']:
                linha[emp] = f"{formatar_br(valor).replace('R$ ', '')} - {qtd}"
            else:
                linha[emp] = formatar_br(valor)
        dados_tabela.append(linha)

    linha_total = {"DESCRICAO": "TOTAL"}
    for emp in empresas:
        obrig = df_pivot.loc['OBRIG. A PAGA', emp] if 'OBRIG. A PAGA' in df_pivot.index else 0
        trans = df_pivot.loc['TRANSITORIA', emp] if 'TRANSITORIA' in df_pivot.index else 0
        dif_trans = df_pivot.loc['DIF_TRANS_ADIANT', emp] if 'DIF_TRANS_ADIANT' in df_pivot.index else 0
        
        total_emp = df_pivot[emp].sum() - obrig - trans - dif_trans
        linha_total[emp] = formatar_br(total_emp)
    
    dados_tabela.append(linha_total)
    return pd.DataFrame(dados_tabela)

def tela_login():
    st.title("🔒 Acesso Restrito - Posição Diária")
    with st.form("login"):
        email = st.text_input("Email")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar")
        if entrar:
            db = SessionLocal()
            senha_hash = hashlib.sha256(senha.encode()).hexdigest()
            
            user = db.query(Usuario).filter_by(email=email, senha=senha_hash, ativo=True).first() # 1. Usuario sem S  2. senha e não senha_hash
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
    st.session_state['perfil'] = getattr(user_db, 'perfil', 'Usuario') if user_db else "Usuario"
    db.close()
    
    st.sidebar.write(f"Perfil: {st.session_state['perfil']}")

    if st.session_state['perfil'] == 'Admin':
        with st.sidebar.expander("👥 Gerenciar Usuários"):
            st.markdown("#### Usuários Cadastrados")
            db = SessionLocal()
            users = db.query(Usuario).order_by(Usuario.id.desc()).all()
            db.close()
            selected_id = None
            if users:
                df_users = pd.DataFrame([{
                    'ID': u.id, 
                    'Nome': u.nome, 
                    'Email': u.email, 
                    'Perfil': getattr(u, 'perfil', 'Usuario'), 
                    'Ativo': 'Sim' if u.ativo else 'Não'
                } for u in users])
                selected = st.dataframe(df_users, use_container_width=True, hide_index=True, height=200, on_select="rerun", selection_mode="single-row")
                if selected.get('selection', {}).get('rows'):
                    selected_id = int(df_users.iloc[selected['selection']['rows'][0]]['ID'])
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
                perfil_val = getattr(user_edit, 'perfil', 'Usuario') if user_edit else 'Usuario'
                perfil = st.selectbox("Perfil", ["Usuario", "Admin"], index=0 if perfil_val == "Usuario" else 1)
                ativo = st.checkbox("Ativo", value=user_edit.ativo if user_edit else True)
                
                col1, col2 = st.columns(2)
                btn_salvar = col1.form_submit_button("💾 Salvar", use_container_width=True)
                btn_senha = col2.form_submit_button("🔑 Nova Senha", use_container_width=True, disabled=not bool(user_edit))
                
            if btn_salvar:
                db = SessionLocal()
                if user_edit:
                    u_db = db.query(Usuarios).filter(Usuarios.id == user_edit.id).first()
                    if u_db:
                        u_db.nome = nome
                        if hasattr(u_db, 'perfil'): u_db.perfil = perfil
                        u_db.ativo = ativo
                        msg = "Usuário atualizado!"
                else:
                    if db.query(Usuarios).filter_by(email=email).first():
                        st.error("Erro: Este email já está cadastrado")
                        db.close()
                        st.stop()
                    senha_temp = gerar_senha_aleatoria()
                    novo = Usuarios(
                        nome=nome, 
                        email=email, 
                        senha_hash=gerar_hash(senha_temp), 
                        perfil=perfil, 
                        ativo=ativo
                    )
                    db.add(novo)
                    msg = f"Usuário cadastrado! Senha padrão: `{senha_temp}`"
                db.commit()
                db.close()
                st.success(msg)
                st.rerun()

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
            regs = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data, PosicaoDiaria.empresa == empresa).all()
            total = 0.0
            for r in regs:
                if r.tipo_titulo == 'OBRIG. A PAGA': 
                    total -= (r.valor or 0.0)
                elif r.tipo_titulo not in ['TRANSITORIA', 'DIF_TRANS_ADIANT']: 
                    total += (r.valor or 0.0)
            return total
        finally: 
            db.close()

    @st.cache_data(ttl=10)
    def get_variacao_empresa(data_atual, data_anterior, empresa):
        if not data_anterior: 
            return 0.0
        total_atual = get_total_empresa(data_atual, empresa)
        total_anterior = get_total_empresa(data_anterior, empresa)
        if total_anterior == 0: 
            return 0.0
        return ((total_atual - total_anterior) / total_anterior) * 100

    for emp, col in empresas_cards.items():
        if emp not in empresas_selecionadas: 
            continue
        with col:
            total_hoje = get_total_empresa(DATA_REF_DATE, emp)
            data_ontem_date = DATA_REF_DATE - timedelta(days=1)
            variacao = get_variacao_empresa(DATA_REF_DATE, data_ontem_date, emp)
            delta_cor = "normal" if variacao >= 0 else "inverse"
            st.metric(label=f"TOTAL {emp}", value=formatar_br(total_hoje), delta=f"{variacao:.2f}%", delta_color=delta_cor)

st.markdown("""
    <style>
        .main-title {
            font-size: 20px;       /* era 24px */
            font-weight: 600;
            margin-top: -25px;     /* sobe mais */
            margin-bottom: 8px;    /* menos espaço embaixo */
        }
        [data-testid="stMetricValue"] {
            font-size: 18px!important; /* era 20px */
            display: flex!important;
            align-items: baseline!important;
            gap: 3px!important;
        }
        [data-testid="stMetricValue"]::before {
            content: "R$ ";
            font-size: 11px!important; /* era 12px */
            font-weight: 600;
        }
        [data-testid="stMetricLabel"] {
            font-size: 10px!important; /* era 11px */
        }
    </style>
    <h1 class="main-title">Dashboard Financeira Diária</h1>
""", unsafe_allow_html=True)

def get_all_data():
    return pd.read_sql_table('posicoes_diarias', engine)

tab1, tab2, tab3, tab4 = st.tabs(["Lançamento", "Histórico", "Manutenção","Graficos"])
ITENS = [('CARTEIRA', 'CARTEIRA'), ('MERCADO PAGO', 'MERCADO PAGO'),('VEICULO', 'VEICULO'), ('SEGURADORA', 'SEGURADORA'),('GARANTIA', 'GARANTIA'), ('BANCOS', 'BANCOS'), ('CARTOES', 'CARTOES'),('NOVOS PAGOS', 'NOVOS PAGOS'),('USADOS PAGOS', 'USADOS PAGOS'),('FUNDAO NOVOS', 'FUNDAO NOVOS'),('FIDIC', 'FIDIC'), ('H.B.PECAS', 'H.B.PECAS'),('ESTOQUE PECAS', 'ESTOQUE PECAS'), ('OBRIG. A PAGA', 'OBRIG. A PAGA'),('ADIANTAMENTOS', 'ADIANTAMENTOS'),('TRANSITORIA', 'TRANSITORIA'), ('DIF_TRANS_ADIANT', 'DIF_TRANS_ADIANT')]
ITENS_MANUAIS = [('NOVOS PAGOS', 'NOVOS PAGOS'),('USADOS PAGOS', 'USADOS PAGOS'),('FUNDAO NOVOS', 'FUNDAO NOVOS'),('FIDIC', 'FIDIC'), ('H.B.PECAS', 'H.B.PECAS'),('ESTOQUE PECAS', 'ESTOQUE PECAS')]

def carregar_valores_manuais_do_banco(data_ref):
    if isinstance(data_ref, str): 
        data_ref = date.fromisoformat(data_ref)
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
    
    df = df[(df['Empresa'] != '') & (df['Empresa'] != 'nan') & (df['Empresa'] != 'None')]
    df = df[(df['Tipo de Título'] != '') & (df['Tipo de Título'] != 'nan') & (df['Tipo de Título'] != 'None')]
    
    if isinstance(data_ref, str): 
        data_ref = date.fromisoformat(data_ref)
        
    if df.empty: 
        st.warning("Nenhuma linha válida para salvar")
        db.close()
        return
        
    try:
        if modo == 'manutencao':
            for _, row in df.iterrows():
                reg = db.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_ref, PosicaoDiaria.empresa == row['Empresa'], PosicaoDiaria.tipo_titulo == row['Tipo de Título']).first()
                if reg: 
                    reg.valor = float(row['Saldo'])
                    reg.qtd_veiculos = int(row['Qtd'])
                    reg.valor_medio = float(row['ValorMedio'])
                    reg.criado_por = usuario_logado
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
        db.rollback()
        st.error(f"Erro ao salvar: {e}")
        st.dataframe(df)
    finally: 
        db.close()

# ========== ABA 1: LANÇAMENTO ==========
with tab1:
    st.markdown("### Lançamento e Manutenção")

    # 1. Cria contador pra resetar o uploader
    if 'uploader_key' not in st.session_state:
        st.session_state['uploader_key'] = 0

    DATA_MANUTENCAO_DATE = st.date_input("📅 Data de Referência", value=DATA_REF_DATE, format="DD/MM/YYYY", key="data_manutencao")
    
    # 2. Uploader com key dinâmica
    uploaded_files = st.file_uploader(
        "📁 Arraste os 4 arquivos RFN aqui. Se não subir nada, será apenas manutenção", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True, 
        key=f"uploader_tab1_{st.session_state['uploader_key']}"
    )
    

    valores_iniciais, valores_qtd_iniciais = carregar_valores_manuais_do_banco(DATA_MANUTENCAO_DATE)
    valores_digitados = {'MATRIZ': {}, 'WS': {}, 'EUSEBIO': {}}
    valores_qtd_digitados = {'MATRIZ': {}, 'WS': {}, 'EUSEBIO': {}}
    
    col_m, col_ws, col_e = st.columns(3)
    empresas_col = {'MATRIZ': col_m, 'WS': col_ws, 'EUSEBIO': col_e}
    
    for emp, col in empresas_col.items():
        with col: 
            st.markdown(f"**{emp}**")
            for item_chave, item_nome in ITENS_MANUAIS:
                st.markdown(f"{item_nome}")
                key_valor = f"{emp}_{item_chave}_edit_{DATA_MANUTENCAO_DATE.strftime('%Y%m%d')}_{st.session_state['uploader_key']}" # key dinamica tbm
                key_qtd = f"{emp}_QTD_{item_chave}_edit_{DATA_MANUTENCAO_DATE.strftime('%Y%m%d')}_{st.session_state['uploader_key']}" # key dinamica tbm
                valores_digitados[emp][item_chave] = st.text_input(label="", value=valores_iniciais[emp][item_chave], key=key_valor, label_visibility="collapsed")
                if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS']: 
                    valores_qtd_digitados[emp][item_chave] = st.number_input("Qtd", value=int(valores_qtd_iniciais[emp][item_chave]), key=key_qtd, min_value=0, step=1)

    if st.button(f"🚀 Processar e Salvar {DATA_MANUTENCAO_DATE.strftime('%d/%m/%Y')}", key="btn_processar_salvar", type="primary"):
        lista_df = []
        if uploaded_files:
            lista_df = [
                carregar_posicao_analitica(uploaded_files),
                carregar_obrigacoes(uploaded_files),
                carregar_creditos_nao_identificados(uploaded_files),
                carregar_adiantamentos(uploaded_files)
            ]
        
        linhas_manuais = []
        for emp in ['MATRIZ', 'WS', 'EUSEBIO']:
            for item_chave, item_nome in ITENS_MANUAIS:
                valor_str = valores_digitados[emp][item_chave]
                valor_float = converter_valor_br(valor_str)
                qtd = valores_qtd_digitados[emp].get(item_chave, 0) if item_chave in ['NOVOS PAGOS', 'USADOS PAGOS'] else 0
                valor_medio = valor_float / qtd if qtd > 0 else 0.0
                linhas_manuais.append({'Empresa': emp, 'Tipo de Título': item_chave, 'Saldo': valor_float, 'Qtd': qtd, 'ValorMedio': valor_medio})
        
        lista_df.append(pd.DataFrame(linhas_manuais))
        lista_df = [df for df in lista_df if not df.empty]
        
        if lista_df:
            df = pd.concat(lista_df, ignore_index=True)
            df['Saldo'] = pd.to_numeric(df['Saldo'], errors='coerce').fillna(0.0)
            df['Qtd'] = pd.to_numeric(df['Qtd'], errors='coerce').fillna(0)
            df['ValorMedio'] = pd.to_numeric(df['ValorMedio'], errors='coerce').fillna(0.0)
            
            empresas = df['Empresa'].unique()
            novas_linhas = []
            for emp in empresas:
                trans = df[(df['Empresa'] == emp) & (df['Tipo de Título'] == 'TRANSITORIA')]['Saldo'].sum()
                adiant = df[(df['Empresa'] == emp) & (df['Tipo de Título'] == 'ADIANTAMENTOS')]['Saldo'].sum()
                if trans > 0:
                    dif = adiant - trans
                    if dif!= 0: 
                        novas_linhas.append({'Tipo de Título': 'DIF_TRANS_ADIANT', 'Empresa': emp, 'Saldo': dif, 'Qtd': 0, 'ValorMedio': 0.0})
            
            if novas_linhas: 
                df = pd.concat([df, pd.DataFrame(novas_linhas)], ignore_index=True)
                
            salvar_posicao_no_banco(df, DATA_MANUTENCAO_DATE, modo='novo')

            # 3. LIMPA TUDO: uploader + campos manuais
            st.session_state['uploader_key'] += 1
            st.cache_data.clear()
            st.success(f"Processamento concluído para {DATA_MANUTENCAO_DATE.strftime('%d/%m/%Y')}!")
            st.rerun()
        else: 
            st.error("Nenhum dado para salvar. Envie arquivos ou preencha os manuais.")

# ========== ABA 2: HISTÓRICO E EXPORTAÇÃO ==========
with tab2:
    st.markdown("### Exportar Posição do Dia")
    col1, col2 = st.columns(2)
    with col1: 
        data_export = st.date_input("Selecione a Data", value=DATA_REF_DATE, key="data_export")
    with col2: 
        empresas_export = st.multiselect("Empresas", ['MATRIZ', 'WS', 'EUSEBIO'], default=empresas_selecionadas, key="emp_export")
        
    st.markdown("#### Preview da Planilha")
    df_preview = montar_df_dashboard(data_export, empresas_export)
    if df_preview is not None:
        st.dataframe(df_preview, use_container_width=True, hide_index=True)
        st.divider()
        if st.button("📊 Gerar Excel Dashboard", type="primary"):
            with st.spinner("Montando planilha..."): 
                excel_data = gerar_excel_dashboard(data_export, empresas_export)
            if excel_data:
                st.download_button(label="📥 Baixar Planilha", data=excel_data, file_name=f"Posicao_{data_export.strftime('%d%m%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.success("Planilha pronta para download!")
    else: 
        st.warning("Nenhum dado encontrado para a data e empresas selecionadas")

# ========== ABA 3: MANUTENÇÃO ==========
with tab3:
    st.subheader("Manutenção de Dados Carregados")
    st.warning("Edite ou exclua valores já lançados por data e empresa")

    col1, col2 = st.columns(2)
    with col1:
        data_manut = st.date_input("Selecione a Data", format="DD/MM/YYYY", key="data_manut_aba3")
    with col2:
        empresa_manut = st.selectbox("Empresa", ["TODAS", "MATRIZ", "WS", "EUSEBIO"], key="empresa_manut_aba3")

    if st.button("Carregar Dados da Data", key="btn_carregar_manut"):
        session = SessionLocal()
        query = session.query(PosicaoDiaria).filter(PosicaoDiaria.data == data_manut)

        if empresa_manut != "TODAS":
            query = query.filter(PosicaoDiaria.empresa == empresa_manut)

        registros = query.all()
        session.close()

        if not registros:
            st.info("Nenhum dado encontrado para essa data/empresa")
            st.session_state.pop('df_manut', None)
        else:
            df_manut = pd.DataFrame([{
                "id": r.id,
                "tipo_de_titulo": r.tipo_titulo,
                "empresa": r.empresa,
                "saldo": r.valor or 0.0
            } for r in registros])

            df_manut = df_manut.sort_values("tipo_de_titulo")
            st.session_state['df_manut'] = df_manut

    if 'df_manut' in st.session_state and not st.session_state['df_manut'].empty:
        st.write("Edite os valores e clique em Salvar")
        df_editado = st.data_editor(
            st.session_state['df_manut'],
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "tipo_de_titulo": st.column_config.TextColumn("Tipo", disabled=True),
                "empresa": st.column_config.TextColumn("Empresa", disabled=True),
                "saldo": st.column_config.NumberColumn("Saldo R$", format="R$ %.2f", step=0.01)
            },
            hide_index=True,
            use_container_width=True
        )

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("💾 Salvar Alterações", key="btn_salvar_manut"):
                session = SessionLocal()
                for _, row in df_editado.iterrows():
                    reg = session.query(PosicaoDiaria).filter(PosicaoDiaria.id == row['id']).first()
                    if reg:
                        reg.valor = float(row['saldo'])
                session.commit()
                session.close()
                st.cache_data.clear()
                st.success("Alterações salvas com sucesso!")
                del st.session_state['df_manut']
                st.rerun()

# st.info("Amanhã a gente monta os gráficos aqui 📊")
with tab4:
    st.header("Gráficos")
    
    df = get_all_data()
    df.columns = df.columns.str.lower()
    
    if not df.empty:
        df['data'] = pd.to_datetime(df['data'])
        
        st.subheader("Filtros")
        col1, col2, col3 = st.columns(3)
        with col1:
            data_inicio = st.date_input("Data Início", df['data'].min())
        with col2:
            data_fim = st.date_input("Data Fim", df['data'].max())
        with col3:
            empresas_filtro = st.multiselect("Filtrar Empresas", df['empresa'].unique(), default=df['empresa'].unique())
        
        mask = (df['data'] >= pd.to_datetime(data_inicio)) & (df['data'] <= pd.to_datetime(data_fim)) & (df['empresa'].isin(empresas_filtro))
        df_filtrado = df[mask]
        
        
        if not df_filtrado.empty:
            df_total_valor = df_filtrado.groupby(['data', 'empresa'])['valor'].sum().reset_index()
            df_total_qtd = df_filtrado.groupby(['data', 'empresa'])['qtd_veiculos'].sum().reset_index()
            data_hoje = df_filtrado['data'].max()
            df_dia = df_filtrado[df_filtrado['data'] == data_hoje]

            st.subheader("KPIs do Período")
            kpi1, kpi2, kpi3 = st.columns(3) # <-- 3 colunas só

            total_geral = df_filtrado['valor'].sum()
            total_dia = df_dia['valor'].sum() if not df_dia.empty else 0
            media_dia = df_total_valor.groupby('data')['valor'].sum().mean() if not df_total_valor.empty else 0

            with kpi1: st.metric("Total R$", f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            with kpi2: st.metric(f"Dia {data_hoje.strftime('%d/%m')}", f"R$ {total_dia:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            with kpi3: st.metric("Média/Dia", f"R$ {media_dia:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            # GRÁFICO 1
            st.subheader("1. Evolução do TOTAL R$ por Empresa")
            fig1 = px.line(df_total_valor, x='data', y='valor', color='empresa', markers=True)
            fig1.update_layout(height=400)
            fig1.update_yaxes(tickprefix="R$ ", tickformat=",.2f") # <-- CORRETO
            st.plotly_chart(fig1, use_container_width=True)

            # GRÁFICO 2
            st.subheader("2. Evolução da QTD de Veículos por Empresa")
            fig2 = px.line(df_total_qtd, x='data', y='qtd_veiculos', color='empresa', markers=True)
            fig2.update_layout(height=400)
            fig2.update_yaxes(tickformat=",.0f") # <-- CORRETO
            st.plotly_chart(fig2, use_container_width=True)

            # GRÁFICO 3
            st.subheader(f"3. Composição R$ por Conta - {data_hoje.strftime('%d/%m/%Y')}")
            fig3 = px.bar(df_dia, x='tipo_titulo', y='valor', color='empresa', barmode='group')
            fig3.update_layout(height=400, xaxis_tickangle=-45)
            fig3.update_yaxes(tickprefix="R$ ", tickformat=",.2f") # <-- CORRETO
            st.plotly_chart(fig3, use_container_width=True)

            # GRÁFICO 4
            st.subheader(f"4. Distribuição % por Conta - {data_hoje.strftime('%d/%m/%Y')}")
            df_pizza = df_dia.groupby('tipo_titulo')['valor'].sum().reset_index()
            fig4 = px.pie(df_pizza, names='tipo_titulo', values='valor', hole=0.4)
            fig4.update_traces(texttemplate='%{label}<br>R$ %{value:,.2f}<br>%{percent}')
            fig4.update_layout(height=400)
            st.plotly_chart(fig4, use_container_width=True)
        
        else:
            st.warning("Nenhum dado encontrado para os filtros selecionados.")
        
    else:
        st.warning("Não há dados no banco ainda.")
