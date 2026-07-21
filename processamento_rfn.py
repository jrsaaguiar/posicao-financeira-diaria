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
    df_raw.columns = ['Empresa', 'AgenteCobrador', 'Saldo']
    
    df_raw['Saldo'] = df_raw['Saldo'].apply(converter_valor_br)
    df_raw['Empresa'] = df_raw['Empresa'].apply(detectar_empresa)
    df_raw['Tipo'] = df_raw['AgenteCobrador'].apply(normalizar_tipo)
    
    df_raw = df_raw[(df_raw['Empresa']!= 'OUTROS') & (df_raw['Saldo'] > 0) & (df_raw['Tipo']!= 'OUTROS')]
    df_agrupado = df_raw.groupby(['Tipo', 'Empresa'], as_index=False)['Saldo'].sum()
    df_agrupado.rename(columns={'Tipo': 'Tipo de Título'}, inplace=True)
    df_agrupado['Qtd'] = 0
    df_agrupado['ValorMedio'] = 0.0
    return df_agrupado

def carregar_obrigacoes(uploaded_files):
    if not uploaded_files: return pd.DataFrame()
    
    arquivo = pegar_arquivo(uploaded_files, ['RFN003_PosicaoAnaliticoPagar.xlsx', 'RFN003_PosicaoAnaliticoPagar.xls'])
    if arquivo is None: return pd.DataFrame()

    engine = 'xlrd' if arquivo.name.endswith('.xls') else 'openpyxl'
    df_obrig_raw = pd.read_excel(arquivo, header=None, engine=engine)
    
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

    linhas = [{'Tipo de Título': 'OBRIG. A PAGA', 'Empresa': e, 'Saldo': v, 'Qtd': 0, 'ValorMedio': 0.0} for e,v in obrig_dict.items() if v > 0]
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
        sal = converter_valor_br(df_cred_raw.iloc[i, df_cred_raw.shape[1]-1])
        if emp in credito_dict and sal > 0:
            credito_dict[emp] += sal

    linhas = [{'Tipo de Título': 'TRANSITORIA', 'Empresa': e, 'Saldo': v, 'Qtd': 0, 'ValorMedio': 0.0} for e,v in credito_dict.items() if v > 0]
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
        sal = converter_valor_br(str(df_ad_raw.iloc[i, 6]))
        if emp in adiant_dict: adiant_dict[emp] += sal

    linhas = [{'Tipo de Título': 'ADIANTAMENTOS', 'Empresa': e, 'Saldo': v, 'Qtd': 0, 'ValorMedio': 0.0} for e,v in adiant_dict.items() if v > 0]
    return pd.DataFrame(linhas)
