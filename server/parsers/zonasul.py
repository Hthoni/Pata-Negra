"""
Parser Zona Sul — formato SuasVendas.

FIX (18/08/2026): o cliente mudou o layout do pedido. O formato antigo tinha
colunas Seq/Cód/Nome/Qtde/IPI%/Peso/R$ Total (sem impostos)/R$ Total c/
impostos. O novo formato é mais enxuto — SEM as colunas Qtde e IPI%
separadas, e sem "R$ Total" sem impostos:

    Seq  Cód(-DV)  Produto  Peso(Kg)  R$  Preço/Kg  R$ Total c/impostos

Ex.: "1 2329-9 BACON DO SEU JEITO KG 600,000 R$ 35,050 21.030,00"

Regex reescrita pra esse layout mais curto. O resto da lógica (match por
nome exato no Perfil, CNPJ casado contra a tabela M:T pelo main.py,
unidade decidida item a item pelo Perfil) continua igual.

Empresa por item vem do Perfil (coluna A / Fat.) — a maioria dos produtos
do Zona Sul é Fat 2 (Distribuidora), mas a Linguiça é Fat 1 (Indústria).

Código do produto tem dígito verificador (ex. 2329-9, 110617-1), então a
regex do código é (\\d+-\\d+).

Unidade por item: o Zona Sul MISTURA unidades no mesmo pedido — a maioria
vem em kg, mas alguns produtos (ex. linguiça, ingredientes de feijoada)
vêm em CAIXAS, com o "Peso (Kg)" na verdade sendo o Nº de caixas e o
"Preço/Kg" sendo o preço POR CAIXA (mesmo padrão já visto no GMAP e na
Vianense) — o Perfil já reflete isso corretamente pra esses dois SKUs
(unidFat='cx', Preço Unit. na mesma base de caixa). emb_tipo é decidido
item a item pelo próprio Perfil: unidFat='cx' -> 'CX' (qtde = nº de
caixas -> kg = qtde x kgCx); senão 'KG' (qtde já em kg).
"""

__cliente_nome__ = "Zona Sul"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

_RE_ITEM = re.compile(
    r'(\d+)\s+(\d+(?:-\d+)?)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+([\d.,]+)\s+R\$\s*([\d,.]+)\s+([\d,.]+)',
    re.M
)


def _num(s):
    return float((s or '0').replace('.', '').replace(',', '.'))


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

    itens = []
    for m in _RE_ITEM.finditer(txt):
        cod = m.group(2)
        nome = re.sub(r'\s+', ' ', m.group(3)).strip()
        qtde_ped = _num(m.group(4))
        preco = _num(m.group(5))
        total = _num(m.group(6))

        pf = match_perfil(nome, produtos)
        emb_tipo = 'CX' if (pf and str(pf.get('unidFat', '')).lower() == 'cx') else 'KG'
        it = processar_item(cod, nome, emb_tipo, 1, qtde_ped, preco, total, produtos)
        itens.append(it)

    if itens:
        # 'filial' é fallback: o main.py sobrescreve pelo nome oficial ao
        # casar o CNPJ contra a tabela de filiais do Perfil (M:T).
        filiais.append({'filial': razao or 'ZONA SUL', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
