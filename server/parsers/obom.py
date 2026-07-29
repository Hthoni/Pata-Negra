"""
Parser O Bom (ex-Grupo Emanuel) — formato SuasVendas (mesmo layout do Zona Sul / Adonai /
Princesa). O cliente migrou do formato antigo (ERP próprio) para o SuasVendas.
CNPJ da loja no cabeçalho ("CNPJ/CPF:"); o main.py casa contra a tabela de
filiais (M:T) do Perfil para enriquecer nome/região/lat/lng.
Código do produto é MISTO: a maioria sem dígito verificador (3425, 48315,
044391) e alguns com (6580-3, 5116-0). Regex: (\\d+(?:-\\d+)?).
Unidade por item: MISTURA unidades — a maioria em kg, mas alguns (linguiça,
feijoada) vêm em CAIXAS, apesar de o PDF rotular "Kg". emb_tipo é decidido
item a item pelo Perfil: unidFat='cx' -> 'CX' (qtde = nº de caixas ->
kg = qtde x kgCx); senão 'KG' (qtde já em kg).

FIX (29/07/2026): pedido nº 1231 (Emanuel/CD-OBOM) veio com "Tabela de
Preço: Não definida" -> a coluna "Preço/Kg" fica em branco na linha do item,
e o "R$" que normalmente aparece colado nela também some. A regex antiga
exigia literalmente "R$" na linha ('...R\\$\\s*([\\d,.]+)...'), então a linha
inteira deixava de casar -> 0 itens extraídos -> parse() retornava [] ->
"Nenhuma filial encontrada no pedido" (o erro não tinha nada a ver com CNPJ
ou perfil, era a extração de item falhando silenciosamente antes disso).
Reescrito para: (1) não exigir mais "R$" na linha, e (2) quando a coluna
Preço/Kg vier ausente (4 números na cauda em vez de 5), derivar o preço
unitário por total ÷ qtde, em vez de capturar por engano o R$ Total como se
fosse o preço/kg.
"""

__cliente_nome__ = "O Bom"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

# seq  cod(-dv opcional)  nome  qtde  + cauda numérica (4 ou 5 números,
# "R$" opcional colado em qualquer um deles, ex.: "R$ 25,74" ou "9.900,00")
_RE_ITEM = re.compile(
    r'(\d+)\s+(\d+(?:-\d+)?)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+([\d.,]+)\s+'
    r'((?:R\$\s*)?[\d,.]+(?:\s+(?:R\$\s*)?[\d,.]+){2,3})',
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
        filiais.append({'filial': razao or 'O BOM', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
