"""
Parser Princesa — formato SuasVendas (mesmo layout do Adonai).
Empresa por item vem do Perfil (coluna A / Fat.). O CNPJ da loja que
pede é lido do cabeçalho ("CNPJ/CPF:") e casado, no main.py, contra a
tabela de filiais (M:T) do Perfil para enriquecer nome/região/lat/lng.

IMPORTANTE: SuasVendas fatura na mesma unidade do pedido (kg), sem
converter para caixas.
"""

__cliente_nome__ = "Princesa"

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

    pedidoNum = fm(r'Informações sobre PEDIDO.*?Nº\s*(\d+)')
    dataPedido = fm(r'Data da Venda:\s*([\d/]+)')
    cnpj = fm(r'CNPJ/CPF:\s*([\d./\-]+)')
    razao = fm(r'Razão Social:\s*(.+?)\s+E-?mail')
    end_m = re.search(r'Endereço:\s*(.+?)CEP', txt)
    endereco = end_m.group(1).strip() if end_m else ''

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
        # 'filial' é fallback: o main.py sobrescreve pelo nome oficial ao
        # casar o CNPJ contra a tabela de filiais do Perfil (M:T).
        filiais.append({'filial': razao or 'PRINCESA', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                         'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                         'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
