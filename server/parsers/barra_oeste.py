"""
Parser Supermercado Barra Oeste — formato SuasVendas (mesmo modelo do Zona Sul / Adonai / Torre).
Princesa). O cliente migrou do formato antigo (ERP próprio) para o SuasVendas.

CNPJ da loja no cabeçalho ("CNPJ/CPF:"); o main.py casa contra a tabela de
filiais (M:T) do Perfil para enriquecer nome/região/lat/lng.

Código do produto é MISTO: a maioria sem dígito verificador (3425, 48315,
044391) e alguns com (6580-3, 5116-0). Regex: (\\d+(?:-\\d+)?).

Unidade por item: MISTURA unidades — a maioria em kg, mas alguns (linguiça,
feijoada) vêm em CAIXAS, apesar de o PDF rotular "Kg". emb_tipo é decidido
item a item pelo Perfil: unidFat='cx' -> 'CX' (qtde = nº de caixas ->
kg = qtde x kgCx); senão 'KG' (qtde já em kg).
"""

__cliente_nome__ = "Supermercado Barra Oeste"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil


def parse(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    def fm(pat):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    # Nº do pedido: usa o do RODAPÉ (Observação -> 'Pedido: NNNNNN'), que é o
    # número que o Henrique controla; cai no Nº do cabeçalho só se faltar.
    pedidoNum = (fm(r'Observaç[ãa]o\s*Pedido:\s*(\d+)')
                 or fm(r'\bPedido:\s*(\d+)')
                 or fm(r'Informações sobre PEDIDO.*?Nº\s*(\d+)'))
    dataPedido = fm(r'Data da Venda:\s*([\d/]+)')
    cnpj = fm(r'CNPJ/CPF:\s*([\d./\-]+)')
    razao = fm(r'Razão Social:\s*(.+?)\s+E-?mail')
    end_m = re.search(r'Endereço:\s*(.+?)CEP', txt)
    endereco = end_m.group(1).strip() if end_m else ''

    # itens: Seq  Cód(-DV opcional)  Nome  Qtde  IPI%  Peso  R$ Preço/Kg  Total
    reItem = re.compile(
        r'(\d+)\s+(\d+(?:-\d+)?)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+([\d.,]+)\s+[\d,.]+\s+[\d,.]+\s+R\$\s*([\d,.]+)\s+([\d,.]+)',
        re.M
    )
    itens = []
    for m in reItem.finditer(txt):
        nome = m.group(3).strip()
        qtde_ped = float(m.group(4).replace('.', '').replace(',', '.'))
        preco = float(m.group(5).replace('.', '').replace(',', '.'))
        total = float(m.group(6).replace('.', '').replace(',', '.'))
        pf = match_perfil(nome, produtos)
        emb_tipo = 'CX' if (pf and str(pf.get('unidFat', '')).lower() == 'cx') else 'KG'
        it = processar_item(m.group(2), nome, emb_tipo, 1, qtde_ped, preco, total, produtos)
        itens.append(it)

    if itens:
        filiais.append({'filial': razao or 'O BOM', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
