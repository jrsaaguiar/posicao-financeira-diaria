# utils.py
import pandas as pd

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
    
    # Ordem importa: do mais específico pro mais genérico
    if t == 'CARTEIRA': return 'CARTEIRA' # tem que ser exato
    if 'MERCADO PAGO' in t or 'MERCADO.PAGO' in t: return 'MERCADO PAGO'
    if 'VENDA VEIC' in t or 'V-VEICULO' in t or 'VEICULO' in t: return 'VEICULO'
    if 'VENDA GARANTIA' in t or 'GARANTIA' in t: return 'GARANTIA'
    if t.startswith('BANCO ') or t == 'BANCO': return 'BANCOS' # só se começar com BANCO
    if 'CARTAO DE CREDITO' in t or 'CARTAO CREDITO' in t: return 'CARTOES' # tira o CREDITO solto
    if 'CARTAO' in t or 'CARTÕES' in t: return 'CARTOES'
    if 'SEGURADORA' in t: return 'SEGURADORA' # tirei o SEGURO solto
    if 'FUNDAO NOVOS' in t or 'FUNDAO' in t: return 'FUNDAO NOVOS'
    if 'NOVOS PAGOS' in t or 'NOVOS.PAGOS' in t: return 'NOVOS PAGOS'
    if 'USADOS PAGOS' in t or 'USADOS.PAGOS' in t: return 'USADOS PAGOS' # tirei USADOS solto
    if 'FIDIC' in t: return 'FIDIC'
    if 'ESTOQUE PECAS' in t or 'EST.PECAS' in t: return 'ESTOQUE PECAS'
    if 'H.B.PECAS' in t or 'HB PECAS' in t: return 'H.B.PECAS' # PECAS solto vira OUTROS
    if 'OBRIG' in t: return 'OBRIG. A PAGA'
    if 'ADIANTAMENTO' in t: return 'ADIANTAMENTOS'
    if 'TRANSITORIA' in t: return 'TRANSITORIA'
    if 'DIF_TRANS_ADIANT' in t: return 'DIF_TRANS_ADIANT'
    return 'OUTROS'