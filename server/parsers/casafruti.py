"""
Parser Casafruti — formato próprio "casa frúti" (PEDIDO DE COMPRA).
Layout de tabela simples: Código | GTIN | Produto | Embalagem | Qtde |
Vlr. Unitário | Vlr. Total. Uma linha por produto (PDF limpo, não escaneado).

Embalagem:
 - 'KG' -> Qtde ja em kg.
 - 'UN' -> Qtde em unidades; converte pelo peso no nome (ex.: 400G -> 0,4 kg/un).

Empresa (faturamento) detectada pelo CNPJ do FORNECEDOR:
 - 10.171.633 -> Industria (empresa 1)
 - 56.423.719 -> Distribuidora (empresa 2)
Cliente identificado pelo CNPJ da LOJA (casa com a tabela de filiais do perfil).
"""

__cliente_nome__ = "Casafruti"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

CNPJ_INDUSTRIA = '10.171.633'

# item: Codigo GTIN Produto Embalagem(KG|UN) Qtde Vlr.Unit Vlr.Total
_ITEM = re.compile(
    r'^(\d+)\s+(\d+)\s+(.+?)\s+(KG|UN)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$',
    re.M
)


def _num(s):
    return float((s or '0').replace('.', '').replace(',', '.'))


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    def fm(pat):
        m = re.search(pat, txt, re.I | re.S)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'PEDIDO DE COMPRA Nº\s*(\d+)')
    dataEntrega = fm(r'ENTREGAR EM\s*([\d/]+)')
    forn = fm(r'FORNECEDOR[\s\S]*?CPF/CNPJ\.?:\s*([\d./\-]+)')
    cnpj = fm(r'LOJA[\s\S]*?CPF/CNPJ\.?:\s*([\d./\-]+)')
    empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in forn.replace('.', '').replace('-', '') else 2

    itens = []
    for m in _ITEM.finditer(txt):
        nome = m.group(3).strip()
        emb = m.group(4).upper()
        qtde = _num(m.group(5))
        preco = _num(m.group(6))
        total = _num(m.group(7))
        if emb == 'KG':
            kg = qtde
        else:  # UN
            mg = re.search(r'(\d+)\s*G\b', nome, re.I)
            gramas = int(mg.group(1)) if mg else 1000
            kg = qtde * gramas / 1000.0
        it = processar_item('', nome, 'KG', 1, kg, preco, total, produtos)
        it['empresa'] = empresa
        itens.append(it)

    if not itens:
        return []
    return [{
        'filial': 'CASAFRUTI', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
        'endereco': '', 'dataPedido': '', 'dataEntrega': dataEntrega,
        'condPgto': '', 'empresa': empresa, 'itens': itens
    }]
