"""
Parser Padrão do Fonseca — formato SuasVendas.

FIX (03/09/2026): o cliente mudou o layout do pedido (mesma migração já
vista no Zona Sul em 18/08/2026, no Torre em 20/08/2026, e no GMAP em
03/09/2026 — 4ª ocorrência da mesma mudança do fornecedor SuasVendas). O
formato antigo tinha colunas Seq/Cód/Nome/Qtde/IPI%/Peso/R$ Total (sem
impostos)/R$ Total c/ impostos. O novo formato é mais enxuto — SEM as
colunas Qtde e IPI% separadas, e sem "R$ Total" sem impostos:

    Seq  Cód(-DV)  Produto  Peso(Kg)  R$  Preço/Kg  R$ Total c/impostos

Ex.: "1 48315 BACON DEF PATA NEGRA PORC KG 80,000 R$ 35,200 2.816,00"

Regex reescrita pra esse layout mais curto. O resto da lógica (match por
nome exato no Perfil, CNPJ casado contra a tabela M:T pelo main.py,
unidade decidida item a item pelo Perfil) continua igual.

Código do produto é MISTO: a maioria sem dígito verificador (48315,
61981, 3413) e alguns com (6580-3, 5116-0). Regex: (\\d+(?:-\\d+)?).

Unidade por item: o Padrão do Fonseca MISTURA unidades no mesmo pedido —
a maioria em kg, mas alguns (linguiça, ingrediente de feijoada) vêm em
CAIXAS, com o "Peso (Kg)" na verdade sendo o Nº de caixas e o "Preço/Kg"
sendo o preço POR CAIXA (mesmo padrão já visto em outros clientes
SuasVendas) — o Perfil já reflete isso corretamente pra esses SKUs
(unidFat='cx', Preço Unit. na mesma base de caixa). emb_tipo é decidido
item a item pelo próprio Perfil: unidFat='cx' -> 'CX' (qtde = nº de
caixas -> kg = qtde x kgCx); senão 'KG' (qtde já em kg).
"""

__cliente_nome__ = "Padrão do Fonseca"

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
    pedidoNum = (fm(r'Observaç[ãa]o\s*Pedido:\s*(\d+)')
                 or fm(r'\bPedido:\s*(\d+)')
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
        filiais.append({'filial': razao or 'PADRAO FONSECA', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
