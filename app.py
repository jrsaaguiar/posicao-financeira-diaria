# app inicializao
import hashlib
import random
import string
from datetime import date, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from sqlalchemy import func

from auth import tela_cadastro_usuario, verificar_login
from database import PosicaoDiaria, SessionLocal, Usuarios, engine, init_db
from exportar import gerar_excel_dashboard
from processamento_rfn import (
    carregar_adiantamentos,
    carregar_creditos_nao_identificados,
    carregar_obrigacoes,
    carregar_posicao_analitica,
)
from utils import converter_valor_br, detectar_empresa, normalizar_tipo

st.set_page_config(page_title="Posição Financeira Diária", layout="wide")

init_db()


def gerar_hash(senha):
    return hashlib.sha256(senha.encode()).hexdigest()


def gerar_senha_aleatoria(tamanho=8):
    return "".join(
        random.choice(string.ascii_letters + string.digits) for _ in range(tamanho)
    )


def formatar_br(valor):
    if valor is None or pd.isna(valor):
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


EMPRESAS = ["MATRIZ", "WS", "EUSEBIO"]
ITENS_ORDEM = [
    "CARTEIRA",
    "MERCADO PAGO",
    "VEICULO",
    "SEGURADORA",
    "GARANTIA",
    "BANCOS",
    "CARTOES",
    "NOVOS PAGOS",
    "USADOS PAGOS",
    "FUNDAO NOVOS",
    "FIDIC",
    "H.B.PECAS",
    "ESTOQUE PECAS",
    "OBRIG. A PAGA",
    "ADIANTAMENTOS",
    "TRANSITORIA",
    "DIF_TRANS_ADIANT",
]


@st.cache_data(ttl=10)
def montar_df_dashboard(data_ref, empresas):
    with SessionLocal() as db:
        dados = (
            db.query(PosicaoDiaria)
            .filter(
                PosicaoDiaria.data == data_ref, PosicaoDiaria.empresa.in_(empresas)
            )
            .all()
        )
        if not dados:
            return None

        df_pivot = pd.DataFrame(0.0, index=ITENS_ORDEM, columns=empresas)
        df_qtd = pd.DataFrame(0, index=ITENS_ORDEM, columns=empresas)

        for d in dados:
            if d.tipo_titulo in ITENS_ORDEM and d.empresa in empresas:
                valor_float = float(d.valor or 0.0)
                df_pivot.loc[d.tipo_titulo, d.empresa] = valor_float
                df_qtd.loc[d.tipo_titulo, d.empresa] = (
                    getattr(d, "qtd_veiculos", 0) or 0
                )

    dados_tabela = []
    for item in ITENS_ORDEM:
        linha = {"DESCRICAO": item}
        for emp in empresas:
            valor = df_pivot.loc[item, emp]
            qtd = df_qtd.loc[item, emp]
            if item in ["NOVOS PAGOS", "USADOS PAGOS"]:
                linha[emp] = f"{formatar_br(valor).replace('R$ ', '')} - {qtd}"
            else:
                linha[emp] = formatar_br(valor)
        dados_tabela.append(linha)

    linha_total = {"DESCRICAO": "TOTAL"}
    for emp in empresas:
        obrig = (
            df_pivot.loc["OBRIG. A PAGA", emp]
            if "OBRIG. A PAGA" in df_pivot.index
            else 0
        )
        trans = (
            df_pivot.loc["TRANSITORIA", emp]
            if "TRANSITORIA" in df_pivot.index
            else 0
        )
        dif_trans = (
            df_pivot.loc["DIF_TRANS_ADIANT", emp]
            if "DIF_TRANS_ADIANT" in df_pivot.index
            else 0
        )

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
            senha_hash = gerar_hash(senha)
            with SessionLocal() as db:
                user = (
                    db.query(Usuarios)
                    .filter_by(email=email, senha_hash=senha_hash, ativo=True)
                    .first()
                )

            if user:
                st.session_state["logado"] = True
                st.session_state["usuario"] = user.nome
                st.session_state["email"] = user.email
                st.session_state["perfil"] = user.perfil
                st.rerun()
            else:
                st.error("Email ou senha inválidos")


if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    tela_login()
    st.stop()

if st.session_state["logado"]:
    with SessionLocal() as db:
        user_db = (
            db.query(Usuarios)
            .filter_by(email=st.session_state["email"])
            .first()
        )
        st.session_state["perfil"] = (
            getattr(user_db, "perfil", "Usuario") if user_db else "Usuario"
        )

    st.sidebar.write(f"Perfil: {st.session_state['perfil']}")

    if st.session_state["perfil"] == "Admin":
        with st.sidebar.expander("👥 Gerenciar Usuários"):
            st.markdown("#### Usuários Cadastrados")

            with SessionLocal() as db:
                users = db.query(Usuarios).order_by(Usuarios.email.desc()).all()

            selected_email = None
            if users:
                df_users = pd.DataFrame(
                    [
                        {
                            "Email": u.email,
                            "Nome": u.nome,
                            "Perfil": getattr(u, "perfil", "Usuario"),
                            "Ativo": "Sim" if u.ativo else "Não",
                        }
                        for u in users
                    ]
                )
                selected = st.dataframe(
                    df_users,
                    width="stretch",
                    hide_index=True,
                    height=200,
                    on_select="rerun",
                    selection_mode="single-row",
                )
                if selected.get("selection", {}).get("rows"):
                    selected_email = df_users.iloc[
                        selected["selection"]["rows"][0]
                    ]["Email"]
            else:
                st.warning("Nenhum usuário cadastrado")

            st.divider()
            st.markdown("#### Dados do Usuário")
            user_edit = None
            if selected_email:
                with SessionLocal() as db:
                    user_edit = (
                        db.query(Usuarios)
                        .filter(Usuarios.email == selected_email)
                        .first()
                    )

            with st.form("form_usuario"):
                nome = st.text_input(
                    "Nome", value=user_edit.nome if user_edit else ""
                )
                email = st.text_input(
                    "Email",
                    value=user_edit.email if user_edit else "",
                    disabled=bool(user_edit),
                )
                perfil_val = (
                    getattr(user_edit, "perfil", "Usuario")
                    if user_edit
                    else "Usuario"
                )
                perfil = st.selectbox(
                    "Perfil",
                    ["Usuario", "Admin"],
                    index=0 if perfil_val == "Usuario" else 1,
                )
                ativo = st.checkbox(
                    "Ativo", value=user_edit.ativo if user_edit else True
                )

                col1, col2 = st.columns(2)
                btn_salvar = col1.form_submit_button(
                    "💾 Salvar", width="stretch"
                )
                btn_senha = col2.form_submit_button(
                    "🔑 Nova Senha",
                    width="stretch",
                    disabled=not bool(user_edit),
                )

            if btn_salvar:
                with SessionLocal() as db:
                    if user_edit:
                        u_db = (
                            db.query(Usuarios)
                            .filter(Usuarios.email == user_edit.email)
                            .first()
                        )
                        if u_db:
                            u_db.nome = nome
                            if hasattr(u_db, "perfil"):
                                u_db.perfil = perfil
                            u_db.ativo = ativo
                            msg = "Usuário atualizado!"
                    else:
                        if db.query(Usuarios).filter_by(email=email).first():
                            st.error("Erro: Este email já está cadastrado")
                            st.stop()
                        senha_temp = gerar_senha_aleatoria()
                        novo = Usuarios(
                            nome=nome,
                            email=email,
                            senha_hash=gerar_hash(senha_temp),
                            perfil=perfil,
                            ativo=ativo,
                        )
                        db.add(novo)
                        msg = f"Usuário cadastrado! Senha padrão: `{senha_temp}`"
                    db.commit()
                st.success(msg)
                st.rerun()

    st.sidebar.divider()
    st.sidebar.markdown("### Data de Referência")
    DATA_REF_DATE = st.sidebar.date_input(
        "Selecione a Data",
        value=date.today(),
        format="DD/MM/YYYY",
        key="data_lancamento",
    )
    st.sidebar.markdown("### Filtros")
    empresas_selecionadas = st.sidebar.multiselect(
        "Empresas",
        ["MATRIZ", "WS", "EUSEBIO"],
        default=["MATRIZ", "WS", "EUSEBIO"],
    )
    st.sidebar.divider()
    st.sidebar.markdown("### Resumo do Dia")

    col_m, col_ws, col_e = st.sidebar.columns(3)
    empresas_cards = {"MATRIZ": col_m, "WS": col_ws, "EUSEBIO": col_e}

    @st.cache_data(ttl=10)
    def get_total_empresa(data, empresa):
        with SessionLocal() as db:
            regs = (
                db.query(PosicaoDiaria)
                .filter(
                    PosicaoDiaria.data == data, PosicaoDiaria.empresa == empresa
                )
                .all()
            )

            total = 0.0
            for r in regs:
                valor_float = float(r.valor or 0.0)

                if r.tipo_titulo == "OBRIG. A PAGA":
                    total -= valor_float
                elif r.tipo_titulo not in ["TRANSITORIA", "DIF_TRANS_ADIANT"]:
                    total += valor_float

            return total

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
            st.metric(
                label=f"TOTAL {emp}",
                value=formatar_br(total_hoje),
                delta=f"{variacao:.2f}%",
                delta_color=delta_cor,
            )

st.markdown(
    """
    <style>
       .main-title {
            font-size: 20px;
            font-weight: 600;
            margin-top: -25px;
            margin-bottom: 8px;
        }
        [data-testid="stMetricValue"] {
            font-size: 18px!important;
            display: flex!important;
            align-items: baseline!important;
            gap: 3px!important;
        }
        [data-testid="stMetricValue"]::before {
            content: "R$ ";
            font-size: 11px!important;
            font-weight: 600;
        }
        [data-testid="stMetricLabel"] {
            font-size: 10px!important;
        }
    </style>
    <h1 class="main-title">Dashboard Financeira Diária</h1>
""",
    unsafe_allow_html=True,
)


def get_all_data():
    with engine.connect() as conn:
        return pd.read_sql("SELECT * FROM posicoes_diarias", conn)


tab1, tab2, tab3, tab4 = st.tabs(
    ["Lançamento", "Histórico", "Manutenção", "Graficos"]
)
ITENS = [
    ("CARTEIRA", "CARTEIRA"),
    ("MERCADO PAGO", "MERCADO PAGO"),
    ("VEICULO", "VEICULO"),
    ("SEGURADORA", "SEGURADORA"),
    ("GARANTIA", "GARANTIA"),
    ("BANCOS", "BANCOS"),
    ("CARTOES", "CARTOES"),
    ("NOVOS PAGOS", "NOVOS PAGOS"),
    ("USADOS PAGOS", "USADOS PAGOS"),
    ("FUNDAO NOVOS", "FUNDAO NOVOS"),
    ("FIDIC", "FIDIC"),
    ("H.B.PECAS", "H.B.PECAS"),
    ("ESTOQUE PECAS", "ESTOQUE PECAS"),
    ("OBRIG. A PAGA", "OBRIG. A PAGA"),
    ("ADIANTAMENTOS", "ADIANTAMENTOS"),
    ("TRANSITORIA", "TRANSITORIA"),
    ("DIF_TRANS_ADIANT", "DIF_TRANS_ADIANT"),
]
ITENS_MANUAIS = [
    ("NOVOS PAGOS", "NOVOS PAGOS"),
    ("USADOS PAGOS", "USADOS PAGOS"),
    ("FUNDAO NOVOS", "FUNDAO NOVOS"),
    ("FIDIC", "FIDIC"),
    ("H.B.PECAS", "H.B.PECAS"),
    ("ESTOQUE PECAS", "ESTOQUE PECAS"),
]


def carregar_valores_manuais_do_banco(data_ref):
    if isinstance(data_ref, str):
        data_ref = date.fromisoformat(data_ref)
    valores = {
        "MATRIZ": {k: 0.0 for k, _ in ITENS_MANUAIS},
        "WS": {k: 0.0 for k, _ in ITENS_MANUAIS},
        "EUSEBIO": {k: 0.0 for k, _ in ITENS_MANUAIS},
    }
    valores_qtd = {
        "MATRIZ": {"NOVOS PAGOS": 0, "USADOS PAGOS": 0},
        "WS": {"NOVOS PAGOS": 0, "USADOS PAGOS": 0},
        "EUSEBIO": {"NOVOS PAGOS": 0, "USADOS PAGOS": 0},
    }

    with SessionLocal() as db:
        registros = (
            db.query(PosicaoDiaria)
            .filter(
                PosicaoDiaria.data == data_ref,
                PosicaoDiaria.tipo_titulo.in_([k for k, _ in ITENS_MANUAIS]),
            )
            .all()
        )
        for reg in registros:
            if reg.empresa in valores and reg.tipo_titulo in valores[reg.empresa]:
                valores[reg.empresa][reg.tipo_titulo] = float(reg.valor or 0.0)
            if (
                reg.empresa in valores_qtd
                and reg.tipo_titulo in valores_qtd[reg.empresa]
            ):
                valores_qtd[reg.empresa][reg.tipo_titulo] = int(
                    getattr(reg, "qtd_veiculos", 0) or 0
                )
    return valores, valores_qtd


def salvar_posicao_no_banco(df, data_ref, modo="novo"):
    usuario_logado = st.session_state.get("email", "sistema")
    df = df.copy()

    # 1. Garantir existência das colunas
    if "Qtd" not in df.columns:
        df["Qtd"] = 0
    if "ValorMedio" not in df.columns:
        df["ValorMedio"] = 0.0

    # 2. Converter valores numéricos tratando NaNs e Infs
    df["Saldo"] = pd.to_numeric(df["Saldo"], errors="coerce").fillna(0.0)
    df["Qtd"] = pd.to_numeric(df["Qtd"], errors="coerce").fillna(0).astype(int)
    df["ValorMedio"] = (
        pd.to_numeric(df["ValorMedio"], errors="coerce")
        .fillna(0.0)
        .replace([np.inf, -np.inf], 0.0)
    )

    # 3. Tratar e limpar colunas de texto
    df["Empresa"] = df["Empresa"].astype(str).str.strip()
    df["Tipo de Título"] = df["Tipo de Título"].astype(str).str.strip()

    # Filtra mantendo apenas empresas e títulos válidos
    invalidos = ["nan", "none", "", "null", "<na>"]
    df = df[~df["Empresa"].str.lower().isin(invalidos)]
    df = df[~df["Tipo de Título"].str.lower().isin(invalidos)]

    if isinstance(data_ref, str):
        data_ref = date.fromisoformat(data_ref)

    if df.empty:
        st.warning("Nenhuma linha válida para salvar.")
        return

    db = SessionLocal()
    try:
        if modo == "manutencao":
            for _, row in df.iterrows():
                emp_val = str(row["Empresa"])
                tipo_val = str(row["Tipo de Título"])

                reg = (
                    db.query(PosicaoDiaria)
                    .filter(
                        PosicaoDiaria.data == data_ref,
                        PosicaoDiaria.empresa == emp_val,
                        PosicaoDiaria.tipo_titulo == tipo_val,
                    )
                    .first()
                )

                if reg:
                    reg.valor = float(row["Saldo"])
                    if hasattr(reg, "qtd_veiculos"):
                        reg.qtd_veiculos = int(row["Qtd"])
                    if hasattr(reg, "valor_medio"):
                        reg.valor_medio = float(row["ValorMedio"])
                    if hasattr(reg, "criado_por"):
                        reg.criado_por = usuario_logado
                else:
                    dados = dict(
                        data=data_ref,
                        empresa=emp_val,
                        tipo_titulo=tipo_val,
                        valor=float(row["Saldo"]),
                    )
                    if hasattr(PosicaoDiaria, "qtd_veiculos"):
                        dados["qtd_veiculos"] = int(row["Qtd"])
                    if hasattr(PosicaoDiaria, "valor_medio"):
                        dados["valor_medio"] = float(row["ValorMedio"])
                    if hasattr(PosicaoDiaria, "criado_por"):
                        dados["criado_por"] = usuario_logado
                    db.add(PosicaoDiaria(**dados))
        else:
            db.query(PosicaoDiaria).filter(
                PosicaoDiaria.data == data_ref
            ).delete()
            objs = []
            for _, row in df.iterrows():
                dados = dict(
                    data=data_ref,
                    empresa=str(row["Empresa"]),
                    tipo_titulo=str(row["Tipo de Título"]),
                    valor=float(row["Saldo"]),
                )
                if hasattr(PosicaoDiaria, "qtd_veiculos"):
                    dados["qtd_veiculos"] = int(row["Qtd"])
                if hasattr(PosicaoDiaria, "valor_medio"):
                    dados["valor_medio"] = float(row["ValorMedio"])
                if hasattr(PosicaoDiaria, "criado_por"):
                    dados["criado_por"] = usuario_logado

                objs.append(PosicaoDiaria(**dados))

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
    st.header(f"Lançamento da Posição Diária - {DATA_REF_DATE.strftime('%d/%m/%Y')}")

    valores_existentes, qtd_existente = carregar_valores_manuais_do_banco(
        DATA_REF_DATE
    )

    with st.form("form_lancamento_unico"):
        # --- Importação de Arquivos Selecionando Todos de Uma Vez ---
        st.subheader("📁 Importação de Relatórios Automáticos")
        st.caption(
            "💡 **Dica:** Selecione ou arraste todos os relatórios do dia de uma só vez."
        )

        arquivos_carregados = st.file_uploader(
            "Selecione os relatórios em lote (Excel / CSV)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="up_arquivos_multiplos",
        )

        st.divider()

        # --- Valores Manuais ---
        st.subheader("📝 Lançamentos Manuais por Empresa")
        dados_manuais_form = []

        cols_emp = st.columns(len(empresas_selecionadas))
        for idx, emp in enumerate(empresas_selecionadas):
            with cols_emp[idx]:
                st.markdown(f"### **{emp}**")

                for item_codigo, item_nome in ITENS_MANUAIS:
                    val_padrao = float(
                        valores_existentes.get(emp, {}).get(item_codigo, 0.0)
                    )
                    qtd_padrao = int(
                        qtd_existente.get(emp, {}).get(item_codigo, 0)
                    )

                    v_input = st.number_input(
                        f"{item_nome} (R$)",
                        value=val_padrao,
                        step=100.0,
                        format="%.2f",
                        key=f"{emp}_{item_codigo}_val",
                    )

                    qtd_input = 0
                    if item_codigo in ["NOVOS PAGOS", "USADOS PAGOS"]:
                        qtd_input = st.number_input(
                            f"Qtd Veículos - {item_nome}",
                            value=qtd_padrao,
                            step=1,
                            key=f"{emp}_{item_codigo}_qtd",
                        )

                    val_medio = (
                        (v_input / qtd_input) if qtd_input > 0 else 0.0
                    )

                    dados_manuais_form.append(
                        {
                            "Empresa": emp,
                            "Tipo de Título": item_codigo,
                            "Saldo": float(v_input),
                            "Qtd": int(qtd_input),
                            "ValorMedio": float(val_medio),
                        }
                    )

        st.divider()

        # --- Botão Único de Ação ---
        btn_salvar_tudo = st.form_submit_button(
            "🚀 Salvar e Processar Posição Completa",
            type="primary",
            width="stretch",
        )

        if btn_salvar_tudo:
            dfs_para_salvar = []

            with st.spinner("Classificando e processando relatórios..."):
                arquivos_classificados = {
                    "posicao": [],
                    "obrigacoes": [],
                    "creditos": [],
                    "adiantamentos": [],
                }

                # Classificação automática por palavras-chave
                if arquivos_carregados:
                    for arq in arquivos_carregados:
                        nome_lower = arq.name.lower()

                        if any(k in nome_lower for k in ["posicao", "analitica", "posição"]):
                            arquivos_classificados["posicao"].append(arq)
                        elif any(k in nome_lower for k in ["obrig", "obrigacao", "obricações"]):
                            arquivos_classificados["obrigacoes"].append(arq)
                        elif any(k in nome_lower for k in ["credito", "crédito", "nao_ident"]):
                            arquivos_classificados["creditos"].append(arq)
                        elif any(k in nome_lower for k in ["adiant", "adiantamento"]):
                            arquivos_classificados["adiantamentos"].append(arq)
                        else:
                            arquivos_classificados["posicao"].append(arq)

                # Processar os arquivos identificados
                if arquivos_classificados["posicao"]:
                    df_p = carregar_posicao_analitica(arquivos_classificados["posicao"])
                    if df_p is not None and not df_p.empty:
                        dfs_para_salvar.append(df_p)

                if arquivos_classificados["obrigacoes"]:
                    df_o = carregar_obrigacoes(arquivos_classificados["obrigacoes"])
                    if df_o is not None and not df_o.empty:
                        dfs_para_salvar.append(df_o)

                if arquivos_classificados["creditos"]:
                    df_c = carregar_creditos_nao_identificados(arquivos_classificados["creditos"])
                    if df_c is not None and not df_c.empty:
                        dfs_para_salvar.append(df_c)

                if arquivos_classificados["adiantamentos"]:
                    df_a = carregar_adiantamentos(arquivos_classificados["adiantamentos"])
                    if df_a is not None and not df_a.empty:
                        dfs_para_salvar.append(df_a)

                # Adicionar Lançamentos Manuais
                df_manuais = pd.DataFrame(dados_manuais_form)
                if not df_manuais.empty:
                    dfs_para_salvar.append(df_manuais)

                # Concatenar Tudo e Gravar no Banco
                if dfs_para_salvar:
                    df_final = pd.concat(dfs_para_salvar, ignore_index=True)
                    salvar_posicao_no_banco(
                        df_final, DATA_REF_DATE, modo="manutencao"
                    )
                else:
                    st.warning(
                        "Nenhum arquivo ou dado manual válido foi informado."
                    )


# ========== ABA 2: HISTÓRICO E EXPORTAÇÃO ==========
with tab2:
    st.markdown("### Exportar Posição do Dia")
    col1, col2 = st.columns(2)
    with col1:
        data_export = st.date_input(
            "Selecione a Data", value=DATA_REF_DATE, key="data_export"
        )
    with col2:
        empresas_export = st.multiselect(
            "Empresas",
            ["MATRIZ", "WS", "EUSEBIO"],
            default=empresas_selecionadas,
            key="emp_export",
        )

    st.markdown("#### Preview da Planilha")
    df_preview = montar_df_dashboard(data_export, empresas_export)
    if df_preview is not None:
        st.dataframe(df_preview, width="stretch", hide_index=True)
        st.divider()
        if st.button("📊 Gerar Excel Dashboard", type="primary"):
            with st.spinner("Montando planilha..."):
                excel_data = gerar_excel_dashboard(data_export, empresas_export)
            if excel_data:
                st.download_button(
                    label="📥 Baixar Planilha",
                    data=excel_data,
                    file_name=f"Posicao_{data_export.strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("Planilha pronta para download!")
    else:
        st.warning(
            "Nenhum dado encontrado para a data e empresas selecionadas"
        )

# ========== ABA 3: MANUTENÇÃO ==========
with tab3:
    st.subheader("Manutenção de Dados Carregados")
    st.warning("Edite ou exclua valores já lançados por data e empresa")

    col1, col2 = st.columns(2)
    with col1:
        data_manut = st.date_input(
            "Selecione a Data", format="DD/MM/YYYY", key="data_manut_aba3"
        )
    with col2:
        empresa_manut = st.selectbox(
            "Empresa",
            ["TODAS", "MATRIZ", "WS", "EUSEBIO"],
            key="empresa_manut_aba3",
        )

    if st.button("Carregar Dados da Data", key="btn_carregar_manut"):
        with SessionLocal() as session:
            query = session.query(PosicaoDiaria).filter(
                PosicaoDiaria.data == data_manut
            )
            if empresa_manut != "TODAS":
                query = query.filter(PosicaoDiaria.empresa == empresa_manut)
            registros = query.all()

            if not registros:
                st.info("Nenhum dado encontrado para essa data/empresa")
                st.session_state.pop("df_manut", None)
            else:
                df_manut = pd.DataFrame(
                    [
                        {
                            "id": r.id,
                            "tipo_de_titulo": r.tipo_titulo,
                            "empresa": r.empresa,
                            "saldo": r.valor or 0.0,
                        }
                        for r in registros
                    ]
                )
                df_manut = df_manut.sort_values("tipo_de_titulo")
                st.session_state["df_manut"] = df_manut

    if "df_manut" in st.session_state and not st.session_state["df_manut"].empty:
        st.write("Edite os valores e clique em Salvar")
        df_editado = st.data_editor(
            st.session_state["df_manut"],
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "tipo_de_titulo": st.column_config.TextColumn(
                    "Tipo", disabled=True
                ),
                "empresa": st.column_config.TextColumn(
                    "Empresa", disabled=True
                ),
                "saldo": st.column_config.NumberColumn(
                    "Saldo R$", format="R$ %.2f", step=0.01
                ),
            },
            hide_index=True,
            width="stretch",
        )

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("💾 Salvar Alterações", key="btn_salvar_manut"):
                with SessionLocal() as session:
                    for _, row in df_editado.iterrows():
                        reg = (
                            session.query(PosicaoDiaria)
                            .filter(PosicaoDiaria.id == row["id"])
                            .first()
                        )
                        if reg:
                            reg.valor = float(row["saldo"])
                    session.commit()
                st.cache_data.clear()
                st.success("Alterações salvas com sucesso!")
                del st.session_state["df_manut"]
                st.rerun()

# ========== ABA 4: GRÁFICOS ==========
with tab4:
    st.header("Gráficos")

    df = get_all_data()
    df.columns = df.columns.str.lower()

    if not df.empty:
        df["data"] = pd.to_datetime(df["data"])

        st.subheader("Filtros")
        col1, col2, col3 = st.columns(3)
        with col1:
            data_inicio = st.date_input("Data Início", df["data"].min())
        with col2:
            data_fim = st.date_input("Data Fim", df["data"].max())
        with col3:
            empresas_filtro = st.multiselect(
                "Filtrar Empresas",
                df["empresa"].unique(),
                default=df["empresa"].unique(),
            )

        mask = (
            (df["data"] >= pd.to_datetime(data_inicio))
            & (df["data"] <= pd.to_datetime(data_fim))
            & (df["empresa"].isin(empresas_filtro))
        )
        df_filtrado = df[mask]

        if not df_filtrado.empty:
            df_total_valor = (
                df_filtrado.groupby(["data", "empresa"])["valor"]
                .sum()
                .reset_index()
            )
            if "qtd_veiculos" in df_filtrado.columns:
                df_total_qtd = (
                    df_filtrado.groupby(["data", "empresa"])["qtd_veiculos"]
                    .sum()
                    .reset_index()
                )
            else:
                df_total_qtd = pd.DataFrame()
            data_hoje = df_filtrado["data"].max()
            df_dia = df_filtrado[df_filtrado["data"] == data_hoje]

            st.subheader("KPIs do Período")
            kpi1, kpi2, kpi3 = st.columns(3)
            total_geral = df_filtrado["valor"].sum()
            total_dia = df_dia["valor"].sum() if not df_dia.empty else 0
            media_dia = (
                df_total_valor.groupby("data")["valor"].sum().mean()
                if not df_total_valor.empty
                else 0
            )

            with kpi1:
                st.metric(
                    "Total R$",
                    f"R$ {total_geral:,.2f}".replace(",", "X")
                    .replace(".", ",")
                    .replace("X", "."),
                )
            with kpi2:
                st.metric(
                    f"Dia {data_hoje.strftime('%d/%m')}",
                    f"R$ {total_dia:,.2f}".replace(",", "X")
                    .replace(".", ",")
                    .replace("X", "."),
                )
            with kpi3:
                st.metric(
                    "Média/Dia",
                    f"R$ {media_dia:,.2f}".replace(",", "X")
                    .replace(".", ",")
                    .replace("X", "."),
                )

            st.subheader("1. Evolução do TOTAL R$ por Empresa")
            fig1 = px.line(
                df_total_valor,
                x="data",
                y="valor",
                color="empresa",
                markers=True,
            )
            fig1.update_layout(height=400)
            fig1.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
            st.plotly_chart(fig1, width="stretch")

            if not df_total_qtd.empty:
                st.subheader("2. Evolução da QTD de Veículos por Empresa")
                fig2 = px.line(
                    df_total_qtd,
                    x="data",
                    y="qtd_veiculos",
                    color="empresa",
                    markers=True,
                )
                fig2.update_layout(height=400)
                fig2.update_yaxes(tickformat=",.0f")
                st.plotly_chart(fig2, width="stretch")

            st.subheader(
                f"3. Composição R$ por Conta - {data_hoje.strftime('%d/%m/%Y')}"
            )
            fig3 = px.bar(
                df_dia,
                x="tipo_titulo",
                y="valor",
                color="empresa",
                barmode="group",
            )
            fig3.update_layout(height=400, xaxis_tickangle=-45)
            fig3.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
            st.plotly_chart(fig3, width="stretch")

            st.subheader(
                f"4. Distribuição % por Conta - {data_hoje.strftime('%d/%m/%Y')}"
            )
            df_pizza = (
                df_dia.groupby("tipo_titulo")["valor"].sum().reset_index()
            )
            fig4 = px.pie(
                df_pizza, names="tipo_titulo", values="valor", hole=0.4
            )
            fig4.update_traces(
                texttemplate="%{label}<br>R$ %{value:,.2f}<br>%{percent}"
            )
            fig4.update_layout(height=400)
            st.plotly_chart(fig4, width="stretch")

        else:
            st.warning("Nenhum dado encontrado para os filtros selecionados.")

    else:
        st.warning("Não há dados no banco ainda.")
