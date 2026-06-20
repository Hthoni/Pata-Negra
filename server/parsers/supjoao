"""
Parser Supermercado do João — sistema proprietário DOJÃO.
Layout: cabeçalho com filial/endereço/pedido, tabela com colunas
Item | Cód | Produto | Embal | Qtde | Cód.EAN | Cód.Fab | Valor N.F. | ... | *Unitário | Bon. | Valor
Qtde = número de caixas. Valor N.F. = preço por caixa.
"""

__cliente_nome__ = "Superm. do João"

import re
import io
import pdfplumber
from perfil import processar_item

def _limpa_float(txt):
    txt = str(txt).strip().replace('.', '').replace(',', '.')
    try:
        return float(txt)
    except ValueError:
        return 0.0

def parse(pdf_bytes, produtos):
    texto = ''
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto += (page.extract_text() or '') + '\n'

    linhas = texto.splitlines()

    def fm(pat):
        m = re.search(pat, texto, re.IGNORECASE)
        return m.group(1).strip() if m else ''

    # -- Cabeçalho --
    pedidoNum   = fm(r'Pedido\s+N[oº°]?[:\s]*(\d+)')
    data_pedido = fm(r'Data\s+Pedido[:\s]*(\d{2}/\d{2}/\d{4})')
    data_entrega = fm(r'Entrega[:\s]*(\d{2}/\d{2}/\d{4})')
    cond_pgto   = fm(r'Condi[çc][aã]o\s+de\s+Pagamento\s*[\n\r]+([^\n\r]+)')
    solicitante = fm(r'Solicitante[:\s]*(\w+)')
    cnpj_fat    = fm(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})')
    filial_nome = fm(r'Filial[:\s]+(.+?)\s*[-–]\s*\d{2}\.\d{3}')

    # Endereço — linha após "ENTREGA"
    endereco = ''
    for i, ln in enumerate(linhas):
        if re.search(r'\bENTREGA\b', ln):
            resto = re.sub(r'ENTREGA', '', ln).strip()
            if resto:
                endereco = resto
            elif i + 1 < len(linhas):
                endereco = linhas[i + 1].strip()
            break

    # -- Itens --
    # Item Cód Produto Embal Qtde [EAN] [Fab] ValorNF ... ValorTotal
    ITEM_RE = re.compile(
        r'^\s*(\d+)\s+'           # Item
        r'(\d+)\s+'               # Cód
        r'(.+?)\s+'               # Produto
        r'(CX-\d+)\s+'            # Embalagem
        r'(\d+)\s+'               # Qtde (nº caixas)
        r'(?:\d+\s+)?'            # Cód EAN (opcional)
        r'(?:\d+\s+)?'            # Cód Fab (opcional)
        r'([\d.,]+)\s+'           # Valor N.F. = preço por caixa
        r'.+?'                    # colunas intermediárias
        r'([\d.,]+)\s*$',         # Valor Total (último)
        re.IGNORECASE
    )

    itens = []
    for ln in linhas:
        m = ITEM_RE.match(ln)
        if not m:
            continue

        cod_cli  = int(m.group(2))
        nome_raw = m.group(3).strip()
        qtde_cx  = int(m.group(5))
        preco_cx = _limpa_float(m.group(6))
        total    = _limpa_float(m.group(7))

        it = processar_item(
            cod_cli, nome_raw,
            'CX', qtde_cx, qtde_cx,
            preco_cx, total, produtos
        )
        itens.append(it)

    if not itens:
        return []

    return [{
        'filial':      filial_nome or 'Santa Luzia',
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
