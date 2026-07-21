# utils.py
import pandas as pd
import unicodedata

def remover_acentos(texto):
    """Remove acentos e caracteres especiais para comparação segura."""
    if not isinstance(texto, str):
        texto = str(texto)
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

def converter_valor_br(valor):
    """Converte valores numéricos ou strings financeiras em float."""
    if pd.isna(valor):
        return 0.0
    
    # Se já for um número (int ou float)
    if isinstance(valor, (int, float)):
        return float(valor)
    
    val = str(valor).strip()
    if not val or val in ['-', '--', 'None', 'nan', 'NaN']:
        return 0.0
    
    # Limpa formatação de moeda
    val = val.replace('R$', '').replace(' ', '').strip()
    
    # Trata padrão brasileiro (ex: 1.500,50 -> 1500.50)
    if ',' in val:
        val = val.replace('.', '').replace(',', '.')
        
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def detectar_empresa(nome):
    """Detecta o nome da empresa padronizado."""
    nome_norm = remover_acentos(str(nome)).upper()
    
    if 'MATRIZ' in nome_norm:
        return 'MATRIZ'
    if 'WS' in nome_norm:
        return 'WS'
    if 'EUSEBIO' in nome_norm:
        return 'EUSEBIO'
        
    return 'OUTROS'

def normalizar_tipo(titulo):
    """Padroniza a categoria/tipo de título."""
    t = remover_acentos(str(titulo)).upper().strip()
    
    # Ordem importa: do mais específico pro mais genérico
    if t == 'CARTEIRA': return 'CARTEIRA'
    if 'MERCADO PAGO' in t or 'MERCADO.PAGO' in t: return 'MERCADO PAGO'
    if 'VENDA VEIC' in t or 'V-VEICULO' in t or 'VEICULO' in t: return 'VEICULO'
    if 'VENDA GARANTIA' in t or 'GARANTIA' in t: return 'GARANTIA'
    if t.startswith('BANCO ') or t == 'BANCO': return 'BANCOS'
    if 'CARTAO DE CREDITO' in t or 'CARTAO CREDITO' in t or 'CARTAO' in t: return 'CARTOES'
    if 'SEGURADORA' in t: return 'SEGURADORA'
    if 'FUNDAO NOVOS' in t or 'FUNDAO' in t: return 'FUNDAO NOVOS'
    if 'NOVOS PAGOS' in t or 'NOVOS.PAGOS' in t: return 'NOVOS PAGOS'
    if 'USADOS PAGOS' in t or 'USADOS.PAGOS' in t: return 'USADOS PAGOS'
    if 'FIDIC' in t: return 'FIDIC'
    if 'ESTOQUE PECAS' in t or 'EST.PECAS' in t: return 'ESTOQUE PECAS'
    if 'H.B.PECAS' in t or 'HB PECAS' in t: return 'H.B.PECAS'
    if 'OBRIG' in t: return 'OBRIG. A PAGA'
    if 'ADIANTAMENTO' in t: return 'ADIANTAMENTOS'
    if 'TRANSITORIA' in t: return 'TRANSITORIA'
    if 'DIF_TRANS_ADIANT' in t: return 'DIF_TRANS_ADIANT'
    
    return 'OUTROS'
