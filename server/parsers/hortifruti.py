"""
Parser Hortifruti (Neogrid) — formato "PEDIDO DE COMPRA" da Neogrid, usado
pelas redes Americanas / Hortifruti (Natural da Terra).

MUITO diferente do SuasVendas: layout de tabela, descrições que quebram em
várias linhas, cliente identificado por CNPJ (sem nome de filial no texto).

Estratégia: usa pdfplumber.extract_tables() — a tabela de itens junta as
linhas quebradas num só campo por item. Cada item vira:
   [desc-A] N Codigo [desc-B] (Quilograma|Unidade) 8-numeros [desc-C]
A descricao completa = desc-A + desc-B + desc-C.

Unidade de medida:
 - 'Quilograma' -> Qtde Pedida ja esta em KG.
 - 'Unidade'    -> Qtde Pedida em unidades; converte p/ kg pelo peso no nome
                   do produto (ex.: '500G' -> 0,5 kg/un; '400G' -> 0,4 kg/un).
"""

__cliente_nome__ = "Hortifruti"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

# linha de item da tabela Neogrid (tudo num campo). Captura desc antes/meio/depois,
# N, codigo, unidade e os numeros; a Qtde Pedida e o 2o numero, o preco liquido o 5o.
_ITEM = re.compile(
    r'^(?P<a>.*?)\s*(?P<n>\d{1,2})\s+(?P<cod>\d{10,})\s+(?P<b>.*?)\s*'
    r'(?P<un>Quilograma|Unidade)\s+'
    r'(?P<qemb>[\d.,]+)\s+(?P<qped>[\d.,]+)\s+(?P<qbon>[\d.,]+)\s+'
    r'(?P<pbruto>[\d.,]+)\s+(?P<pliq>[\d.,]+)\s+(?P<vdesc>[\d.,]+)\s+'
    r'(?P<ipi>[\d.,]+)\s+(?P<vtot>[\d.,]+)\s*(?P<c>.*)$'
)


def _num(s):
    return float((s or '0').replace('.', '').replace(',', '.'))


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        full_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
        tabelas = []
        for pg in pdf.pages:
            tabelas.extend(pg.extract_tables() or [])

    def fm(pat, txt=full_text):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'Número Pedido:\s*(\d+)')
    dataPedido = fm(r'Data de Emissão:\s*([\d/ ]+)').replace(' ', '')
    dataEntrega = fm(r'Data Inicial:\s*([\d/ ]+)').replace(' ', '')
    cnpj = fm(r'CNPJ do Local de Entrega:\s*([\d./ \-]+)').replace(' ', '')
    if not cnpj:
        cnpj = fm(r'Comprador.*?CNPJ:\s*([\d./ \-]+)').replace(' ', '')

    itens = []
    for t in tabelas:
        for row in t:
            for cell in row:
                if not cell:
                    continue
                linha = cell.replace('\n', ' ').strip()
                m = _ITEM.match(linha)
                if not m:
                    continue
                # descricao completa remontada
                nome = ' '.join(x for x in [m.group('a'), m.group('b'), m.group('c')] if x).strip()
                nome = re.sub(r'\s+', ' ', nome)
                un = m.group('un')
                qped = _num(m.group('qped'))
                preco = _num(m.group('pliq'))
                total = _num(m.group('vtot'))

                if un.lower().startswith('quilog'):
                    kg = qped                     # ja esta em kg
                else:
                    # 'Unidade' -> pega peso no nome (ex.: 500G, 400G) e converte
                    mg = re.search(r'(\d+)\s*G\b', nome, re.I)
                    gramas = int(mg.group(1)) if mg else 1000
                    kg = qped * gramas / 1000.0

                it = processar_item(m.group('cod'), nome, 'KG', 1, kg, preco, total, produtos)
                itens.append(it)

    if not itens:
        return []
    return [{
        'filial': 'HORTIFRUTI', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
        'endereco': '', 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
        'condPgto': '', 'empresa': 2, 'itens': itens
    }]

