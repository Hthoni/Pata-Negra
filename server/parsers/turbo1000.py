"""
Parser Turbo 1000 — sistema DOJÃO, multi-filial.
Layout: cabeçalho com filial/endereço/pedido, tabela com colunas
Item | Cód | Produto | Embal | Qtde | Cód.EAN | Cód.Fab | Valor N.F. | ... | Valor
Qtde = kg totais (KG) ou nº caixas (CX). Valor N.F. = preço por kg ou por cx.
"""
import re, io
import pdfplumber
from perfil import processar_item

__cliente_nome__ = "Turbo 1000"

def _limpa_float(txt):
    try: return float(str(txt).strip().replace('.','').replace(',','.'))
    except: return 0.0

ITEM_RE = re.compile(
    r'^\s*(\d+)\s+'
    r'(\d+)\s+'
    r'(.+?)\s+'
    r'(KG|CX-\d+)\s+'
    r'(\d+)\s+'
    r'(?:\d+\s+)?'
    r'(?:\d+\s+)?'
    r'([\d.,]+)\s+'
    r'.+?'
    r'([\d.,]+)\s*$',
    re.IGNORECASE
)

def parse(pdf_bytes, produtos):
    texto = ''
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto += (page.extract_text() or '') + '\n'

    linhas = texto.splitlines()

    def fm(pat):
        m = re.search(pat, texto, re.IGNORECASE)
        return m.group(1).strip() if m else ''

    pedidoNum    = fm(r'Pedido\s+N[oº°]?[:\s]*(\d+)')
    data_pedido  = fm(r'Data\s+Pedido[:\s]*(\d{2}/\d{2}/\d{4})')
    data_entrega = fm(r'Entrega[:\s]*(\d{2}/\d{2}/\d{4})')
    cond_pgto    = fm(r'Condi[çc][aã]o\s+de\s+Pagamento\s*[\n\r]+([^\n\r]+)')
    solicitante  = fm(r'Solicitante[:\s]*(.+?)(?:\s{2,}|$)')
    cnpj_fat     = fm(r'Filial[^\n]*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})')
    filial_nome  = fm(r'Filial[:\s]+(.+?)\s*[-–]\s*\d{2}\.\d{3}')

    endereco = ''
    for i, ln in enumerate(linhas):
        if re.search(r'\bENTREGA\b', ln):
            resto = re.sub(r'ENTREGA', '', ln).strip()
            endereco = resto if resto else (linhas[i+1].strip() if i+1 < len(linhas) else '')
            break

    itens = []
    for ln in linhas:
        m = ITEM_RE.match(ln)
        if not m: continue
        cod_cli  = int(m.group(2))
        nome_raw = m.group(3).strip()
        emb      = m.group(4).upper()
        qtde     = int(m.group(5))
        preco    = _limpa_float(m.group(6))
        total    = _limpa_float(m.group(7))
        emb_tipo = 'CX' if emb.startswith('CX') else 'KG'
        it = processar_item(cod_cli, nome_raw, emb_tipo, qtde, qtde, preco, total, produtos)
        itens.append(it)

    if not itens: return []

    return [{
        'filial':      filial_nome,
        'pedidoNum':   pedidoNum,
        'cnpj':        cnpj_fat,
        'dataPedido':  data_pedido,
        'dataEntrega': data_entrega,
        'condPgto':    cond_pgto,
        'solicitante': solicitante,
        'endereco':    endereco,
        'empresa':     2,
        'itens':       itens,
    }]
