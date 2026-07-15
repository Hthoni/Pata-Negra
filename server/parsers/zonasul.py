"""
Parser Zona Sul — formato SuasVendas (mesmo layout do Adonai / Princesa).
Empresa por item vem do Perfil (coluna A / Fat.). O CNPJ da loja que
pede é lido do cabeçalho ("CNPJ/CPF:") e casado, no main.py, contra a
tabela de filiais (M:T) do Perfil para enriquecer nome/região/lat/lng.

Código do produto tem dígito verificador (ex. 2329-9, 110617-1), então a
regex do código é (\\d+-\\d+).

Unidade por item: o Zona Sul MISTURA unidades no mesmo pedido — a maioria
vem em kg, mas alguns produtos (ex. linguiça) vêm em CAIXAS, apesar de o
PDF rotular "Kg". Por isso o emb_tipo é decidido item a item pelo próprio
Perfil: se o produto casado tem unidFat='cx', passamos 'CX' (a qtde do PDF
é nº de caixas → kg = qtde × kgCx); senão 'KG' (qtde já em kg).
"""

__cliente_nome__ = "Zona Sul"

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
    pedidoNum = (fm(r'Observaç[ãa]o\s*Pedido:\s*(\d+(?:-\d+)?)')
                 or fm(r'\bPedido:\s*(\d+(?:-\d+)?)')
                 or fm(r'Informações sobre PEDIDO.*?Nº\s*(\d+)'))
    dataPedido = fm(r'Data da Venda:\s*([\d/]+)')
    cnpj = fm(r'CNPJ/CPF:\s*([\d./\-]+)')
    razao = fm(r'Razão Social:\s*(.+?)\s+E-?mail')
    end_m = re.search(r'Endereço:\s*(.+?)CEP', txt)
    endereco = end_m.group(1).strip() if end_m else ''

    # itens: Seq  Cód(-DV)  Nome  Qtde  IPI%  Peso  Preço/Kg  Total
    reItem = re.compile(
        r'(\d+)\s+(\d+-\d+)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+(\d+)\s+[\d,.]+\s+[\d,.]+\s+R\$\s*([\d,.]+)\s+([\d,.]+)',
        re.M
    )
    itens = []
    for m in reItem.finditer(txt):
        nome = m.group(3).strip()
        qtde_ped = float(m.group(4).replace('.', '').replace(',', '.'))
        preco = float(m.group(5).replace('.', '').replace(',', '.'))
        total = float(m.group(6).replace('.', '').replace(',', '.'))
        # Decide a unidade pelo Perfil: produto cadastrado como caixa -> 'CX'
        pf = match_perfil(nome, produtos)
        emb_tipo = 'CX' if (pf and str(pf.get('unidFat', '')).lower() == 'cx') else 'KG'
        it = processar_item(m.group(2), nome, emb_tipo, 1, qtde_ped, preco, total, produtos)
        itens.append(it)

    if itens:
        # 'filial' é fallback: o main.py sobrescreve pelo nome oficial ao
        # casar o CNPJ contra a tabela de filiais do Perfil (M:T).
        filiais.append({'filial': razao or 'ZONA SUL', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
