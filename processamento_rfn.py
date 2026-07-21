# processamento_rfn.py
import pandas as pd
from utils import detectar_empresa, converter_valor_br, normalizar_tipo

def pegar_arquivo(uploaded_files, nomes_possiveis):
    for f in uploaded_files:
        if f.name in nomes_possiveis:
            return f
    return None

def carregar_posicao_analitica(uploaded_files):
    if not uploaded_files: return pd.DataFrame()
    
    arquivo = pegar_arquivo(uploaded_files, ['RFN003_PosicaoAnaliticoReceber_Excel.xls', 'RFN003_PosicaoAnaliticoReceber_Excel.xlsx'])
    if arquivo is None: return pd.DataFrame()
    
    engine = 'xlrd' if arquivo.name.endswith('.xls') else 'openpyxl'
    df_raw = pd.read_excel(arquivo, header=None, usecols=[0, 5, 19], skiprows=1, engine=engine)
    df_raw.columns = ['empresa', 'AgenteCobrador', 'valor']
    
    df_raw['valor'] = df_raw['valor'].apply(converter_valor_br).astype(float)
    df_raw['empresa'] = df_raw['empresa'].apply(detectar_empresa)
    df_raw['tipo_titulo'] = df_raw['AgenteCobrador'].apply(normalizar_tipo)
    
    df_raw = df_raw[(df_raw['empresa'] != 'OUTROS') & (df_raw['valor'] > 0) & (df_raw['tipo_titulo'] != 'OUTROS')]
    df_agrupado = df_raw.groupby(['tipo_titulo', 'empresa'], as_index=False)['valor'].sum()
    df_agrupado['qtd_veiculos'] = 0
    df_agrupado['valor_medio'] = 0.0
    return df_agrupado

def carregar_obrigacoes(uploaded_files):
    if not uploaded_files: return pd.DataFrame()
    
    arquivo = pegar_arquivo(uploaded_files, ['RFN003_PosicaoAnaliticoPagar.xlsx', 'RFN003_PosicaoAnaliticoPagar.xls'])
    if arquivo is None: return pd.DataFrame()

    engine = 'xlrd' if arquivo.name.endswith('.xls') else 'openpyxl'
    df_obrig_raw = pd.read_excel(arquivo, header=None, engine=engine)
    
    obrig_dict = {'MATRIZ': 0.0, 'WS': 0.0, 'EUSEBIO': 0.0}
    max_cols = df_obrig_raw.shape[1]
    
    for i in range(len(df_obrig_raw)):
        linha_toda = " ".join([str(x) for x in df_obrig_raw.iloc[i] if pd.notna(x)]).upper()
        if 'RESUMO' in linha_toda and 'EMPRESA' in linha_toda:
            for j in range(1, 6):
                if i+j >= len(df_obrig_raw): break
                emp = str(df_obrig_raw.iloc[i + j, 0])
                # Proteção de limite de coluna (index 9)
                col_val = 9 if max_cols > 9 else max_cols - 1
                val = float(converter_valor_br(df_obrig_raw.iloc[i + j, col_val]) or 0.0)
                
                if 'EUSEBIO' in emp.upper(): obrig_dict['EUSEBIO'] = val
                elif 'MATRIZ' in emp.upper(): obrig_dict['MATRIZ'] = val
                elif 'WS' in emp.upper(): obrig_dict['WS'] = val
            break

    linhas = [{
        'tipo_titulo': 'OBRIG. A PAGA',
        'empresa': e,
        'valor': v,
        'qtd_veiculos': 0,
        'valor_medio': 0.0
    } for e, v in obrig_dict.items() if v > 0]
    
    return pd.DataFrame(linhas)

def carregar_creditos_nao_identificados(uploaded_files):
    if not uploaded_files: return pd.DataFrame()

    arquivo = pegar_arquivo(uploaded_files, ['RFN024_SaldoCreditosNaoIdentificados.xlsx', 'RFN024_SaldoCreditosNaoIdentificados.xls'])
    if arquivo is None: return pd.DataFrame()

    engine = 'xlrd' if arquivo.name.endswith('.xls') else 'openpyxl'
    df_cred_raw = pd.read_excel(arquivo, header=None, engine=engine)
    
    credito_dict = {'MATRIZ': 0.0, 'WS': 0.0, 'EUSEBIO': 0.0}
    for i in range(len(df_cred_raw)):
        if not str(df_cred_raw.iloc[i, 0]).isdigit(): continue
        emp = detectar_empresa(str(df_cred_raw.iloc[i, 2]))
        sal = float(converter_valor_br(df_cred_raw.iloc[i, df_cred_raw.shape[1]-1]) or 0.0)
        if emp in credito_dict and sal > 0:
            credito_dict[emp] += sal

    linhas = [{
        'tipo_titulo': 'TRANSITORIA',
        'empresa': e,
        'valor': v,
        'qtd_veiculos': 0,
        'valor_medio': 0.0
    } for e, v in credito_dict.items() if v > 0]
    
    return pd.DataFrame(linhas)

def carregar_adiantamentos(uploaded_files):
    if not uploaded_files: return pd.DataFrame()

    arquivo = pegar_arquivo(uploaded_files, ['RFN013_FichaRazaoSaldo_Excel.xls', 'RFN013_FichaRazaoSaldoExcel.xlsx'])
    if arquivo is None: return pd.DataFrame()

    engine = 'xlrd' if arquivo.name.endswith('.xls') else 'openpyxl'
    df_ad_raw = pd.read_excel(arquivo, header=None, engine=engine)

    adiant_dict = {'MATRIZ': 0.0, 'WS': 0.0, 'EUSEBIO': 0.0}
    for i in range(1, len(df_ad_raw)):
        emp = detectar_empresa(str(df_ad_raw.iloc[i, 3]))
        sal = float(converter_valor_br(str(df_ad_raw.iloc[i, 6])) or 0.0)
        if emp in adiant_dict: adiant_dict[emp] += sal

    linhas = [{
        'tipo_titulo': 'ADIANTAMENTOS',
        'empresa': e,
        'valor': v,
        'qtd_veiculos': 0,
        'valor_medio': 0.0
    } for e, v in adiant_dict.items() if v > 0]
    
    return pd.DataFrame(linhas)
