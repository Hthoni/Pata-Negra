"""
Parser Vianense (Supermercados Vianense Ltda / CD Vianense) — formato
SuasVendas (mesmo layout do Zona Sul / Adonai / Princesa / O Bom /
Superprix / Costa Azul / GMAP).
CNPJ da loja no cabeçalho ("CNPJ/CPF:"); o main.py casa contra a tabela de
filiais (M:T) do Perfil para enriquecer nome/região/lat/lng.

Código do produto é MISTO: a maioria sem dígito verificador (12130, 47627)
e alguns com (8841-1, 044391-1 — mesmos códigos "-1" já vistos no GMAP,
indicando uma variante de embalagem específica desses SKUs). Regex:
(\\d+(?:-\\d+)?).

Unidade por item: MISTURA — alguns vêm em KG diretamente (Qtde == Peso(Kg)
no PDF), outros em CAIXAS (Qtde = nº de caixas, Peso(Kg) = Qtde x kgCx).
emb_tipo é decidido item a item pelo Perfil: unidFat='cx' -> 'CX' (qtde =
nº de caixas -> kg = qtde x kgCx); senão 'KG' (qtde já em kg).

Usa a mesma regex de item já corrigida no O Bom/Costa Azul (quantifier
{3,4}, cobrindo 4 ou 5 números na cauda — com ou sem Preço/Kg preenchido).
"""

__cliente_nome__ = "Vianense"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

_RE_ITEM = re.compile(
    r'(\d+)\s+(\d+(?:-\d+)?)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+([\d.,]+)\s+'
    r'((?:R\$\s*)?[\d,.]+(?:\s+(?:R\$\s*)?[\d,.]+){3,4})',
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
        nome = m.group(3).strip()
        qtde_ped = _num(m.group(4))

        nums = [_num(n) for n in re.findall(r'[\d.,]+', m.group(5))]
        if not nums:
            continue
        total = nums[-1]  # R$ Total c/ impostos - sempre a última, autoritativa

        if len(nums) >= 5:
            # IPI%, Peso, Preço/Kg, R$Total, R$Total c/ impostos -> preço veio preenchido
            preco = nums[2]
        else:
            # Preço/Kg ausente (Tabela de Preço "Não definida") -> deriva do total
            preco = round(total / qtde_ped, 4) if qtde_ped else 0.0

        pf = match_perfil(nome, produtos)
        emb_tipo = 'CX' if (pf and str(pf.get('unidFat', '')).lower() == 'cx') else 'KG'
        it = processar_item(cod, nome, emb_tipo, 1, qtde_ped, preco, total, produtos)
        itens.append(it)

    if itens:
        filiais.append({'filial': razao or 'VIANENSE', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
