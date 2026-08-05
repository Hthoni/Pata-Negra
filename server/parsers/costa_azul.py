"""
Parser Costa Azul (Costazul Alimentos Ltda) — formato SuasVendas (mesmo
layout do Zona Sul / Adonai / Princesa / O Bom / Superprix).
CNPJ da loja no cabeçalho ("CNPJ/CPF:"); o main.py casa contra a tabela de
filiais (M:T) do Perfil para enriquecer nome/região/lat/lng — a Rede Costa
Azul tem 7 lojas, cada uma com seu próprio CNPJ (mesma raiz 17.493.338,
sufixo de filial diferente).

Todos os itens vêm em KG diretamente (Qtde == Peso(Kg) em todos os pedidos
de amostra) — sem conversão de caixa/pacote necessária, diferente de outros
clientes SuasVendas que às vezes misturam unidades.

FIX (04/08/2026): a regex de item herdada do parser do O Bom tinha um bug
latente no quantifier — só capturava 4 números no máximo (`{2,3}`), mas a
linha real tem 5 números quando o Preço/Kg vem preenchido (IPI%, Peso,
Preço/Kg, R$ Total, R$ Total c/ impostos). Isso cortava o último número
(R$ Total c/ impostos) do texto capturado; não dava sintoma visível porque
nos pedidos vistos até agora o IPI é sempre 0,00 (os dois totais coincidem),
mas pegaria o total ERRADO (sem imposto) assim que aparecesse IPI != 0.
Corrigido para `{3,4}` (4 a 5 números), cobrindo os dois casos: com e sem
Preço/Kg preenchido.
"""

__cliente_nome__ = "Costa Azul"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

# seq  cod(-dv opcional)  nome  qtde  + cauda numérica (4 ou 5 números,
# "R$" opcional colado em qualquer um deles, ex.: "R$ 25,74" ou "9.900,00")
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
        filiais.append({'filial': razao or 'COSTA AZUL', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
