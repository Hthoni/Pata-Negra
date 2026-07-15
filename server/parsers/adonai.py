"""
Parser Adonai — formato SuasVendas (sistema de terceiro).
Empresa sempre = 2 (Distribuidora), pois não há CNPJ de fornecedor
identificável neste formato.

IMPORTANTE: SuasVendas sempre fatura na mesma unidade do pedido
original (kg ou pacotes), nunca convertendo para caixas — o cliente
exige a nota fiscal na mesma unidade pedida.
"""

__cliente_nome__ = "Adonai Atacadista"

import io
import re
import pdfplumber
from perfil import processar_item


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
    filial_m = re.search(r'Endereço:\s*(.+?)CEP', txt)
    filial = filial_m.group(1).strip().rstrip(',') if filial_m else 'ADONAI'
    endereco = filial_m.group(1).strip() if filial_m else ''

    # itens: Seq  Cód  Nome  Qtde  IPI%  Peso  Preço/Kg  Total
    reItem = re.compile(
        r'(\d+)\s+(\d+)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+(\d+)\s+[\d,.]+\s+[\d,.]+\s+R\$\s*([\d,.]+)\s+([\d,.]+)',
        re.M
    )
    itens = []
    for m in reItem.finditer(txt):
        nome = m.group(3).strip()
        qtde_ped = float(m.group(4).replace('.', '').replace(',', '.'))
        preco = float(m.group(5).replace('.', '').replace(',', '.'))
        total = float(m.group(6).replace('.', '').replace(',', '.'))
        # qtde e preço direto do PDF, sem conversão de unidade
        it = processar_item(m.group(2), nome, 'KG', 1, qtde_ped, preco, total, produtos)
        itens.append(it)

    if itens:
        filiais.append({'filial': 'PADRE MIGUEL', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                         'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                         'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
