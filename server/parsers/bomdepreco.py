"""
Parser Bom de Preço / Redeconomia (grupo, multi-filial) — formato TOTVS
"PEDIDO" (rptPedido.rdlc). Cada PDF é o pedido de UMA filial; o parser também
aceita vários pedidos concatenados num PDF só (divide por "Pedido Nº:").

Faturamento em KG: a coluna Embal. é "KG" e a Qtde já vem em quilos
(Qtde = kg, não caixas). Por isso passa emb_tipo='KG' p/ o processar_item
NÃO multiplicar por kgCx. Confere: soma das Qtdes = "Itens Quantidade" do
rodapé e soma dos valores = SUBTOTAL.

Layout do item:
  {Item} {Cód.Cliente} {Nome...} kg KG {Qtde} {Cód.EAN} {ValorNF} {%Desc}
  {CustoNF} {Unitário} {Valor}
Ancorado por " kg KG " (marca fim do nome e início dos números). Preço/kg =
Valor ÷ Qtde (robusto a desconto). Nome casa por NOME no perfil (col C).
CNPJ da filial no cabeçalho ("Filial: NOME - CNPJ"); o main.py enriquece
nome/número/região/lat/lng pelo CNPJ contra a tabela M:T do perfil.
Fornecedor "PATA NEGRA DISTRIBUIDORA" (empresa herda do perfil, col A).
"""

__cliente_nome__ = "Bom de Preço"

import io
import re
import pdfplumber
from perfil import processar_item

_RE_NUM = re.compile(r'[\d.]+,\d+|\d+')
# item: nº, cód cliente, nome, "kg KG", qtde, EAN, resto (preços + valor)
_RE_ITEM = re.compile(r'^\s*(\d+)\s+(\d+)\s+(.+?)\s+kg\s+KG\s+([\d.,]+)\s+\d+\s+(.+?)\s*$', re.M)


def _num(s):
    return float(str(s).replace('.', '').replace(',', '.'))


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    filiais = []
    # divide em blocos por pedido (cada filial = um pedido). O lookahead
    # mantém o cabeçalho "Pedido Nº:" no início de cada bloco.
    blocos = re.split(r'(?=Pedido\s*N[ºo°]:)', full)
    for txt in blocos:
        if 'Filial:' not in txt:
            continue

        def fm(pat):
            m = re.search(pat, txt, re.I)
            return m.group(1).strip() if m else ''

        pedidoNum = fm(r'Pedido\s*N[ºo°]:\s*(\d+)')
        dataPedido = fm(r'Pedido\s*N[ºo°]:\s*\d+\s+([\d/]+)')
        dataEntrega = fm(r'Entrega:\s*([\d/]+)')
        prazo = fm(r'(\d+)\s*-\s*DIAS')
        condPgto = f'{prazo} dias' if prazo else ''
        cnpj = fm(r'Filial:.*?-\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})')
        filial_nome = fm(r'Filial:\s*(.+?)\s*-\s*\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
        end_m = re.search(r'ENTREGA\s+(.+?)\s+RJ\s+\d{8}\s+Brasil', txt, re.S | re.I)
        endereco = re.sub(r'\s+', ' ', end_m.group(1)).strip() if end_m else ''

        itens = []
        for m in _RE_ITEM.finditer(txt):
            nome = m.group(3).strip()
            qtde = _num(m.group(4))           # já em kg
            nums = _RE_NUM.findall(m.group(5))
            if not nums:
                continue
            total = _num(nums[-1])            # coluna Valor
            preco = round(total / qtde, 2) if qtde else 0.0
            it = processar_item(m.group(2), nome, 'KG', 1, qtde, preco, total, produtos)
            itens.append(it)

        if itens:
            filiais.append({
                'filial': filial_nome or 'BOM DE PREÇO',
                'pedidoNum': pedidoNum,
                'cnpj': cnpj,
                'endereco': endereco,
                'dataPedido': dataPedido,
                'dataEntrega': dataEntrega,
                'condPgto': condPgto,
                'empresa': 2,
                'itens': itens,
            })
    return filiais
