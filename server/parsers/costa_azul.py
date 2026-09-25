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

FIX (04/08/2026): quantifier da cauda numérica ajustado (histórico).

FIX (25/09/2026): o cliente migrou pro layout SuasVendas mais enxuto
(mesma mudança já vista em Zona Sul, Torre, GMAP, Padrão do Fonseca). A
linha do item agora tem só 3 números — Peso(Kg), R$ Preço/Kg, R$ Total —
sem as colunas IPI%/Total-sem-imposto que existiam antes. A regex antiga
exigia de 4 a 5 números na cauda; com só 3, ela falhava e, pior, a busca
gulosa engolia a linha seguinte tentando completar a cauda — resultado
real: 3 itens capturados em vez de 6, com dados de duas linhas misturados
(pesos e preços trocados, preços absurdos tipo R$ 4.762/kg no PDF gerado).
Regex reescrita pro layout novo:
    Seq  Cód(-dv)  Nome  Peso(Kg)  R$ Preço/Kg  R$ Total
O total é a última coluna (não há mais "com impostos" separado). Preço/Kg
vem sempre preenchido nesse layout; mantido o fallback (deriva do total)
por segurança caso volte a faltar.
"""

__cliente_nome__ = "Costa Azul"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

# seq  cod(-dv opcional)  nome  qtde  + cauda numérica (4 ou 5 números,
# "R$" opcional colado em qualquer um deles, ex.: "R$ 25,74" ou "9.900,00")
_RE_ITEM = re.compile(
    r'(\d+)\s+(\d+(?:-\d+)?)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+'
    r'([\d.,]+)\s+R\$\s*([\d,.]+)\s+([\d,.]+)',
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
        # fallback: se por acaso o Preço/Kg vier vazio/zerado, deriva do total
        if not preco and qtde_ped:
            preco = round(total / qtde_ped, 4)

        pf = match_perfil(nome, produtos)
        emb_tipo = 'CX' if (pf and str(pf.get('unidFat', '')).lower() == 'cx') else 'KG'
        it = processar_item(cod, nome, emb_tipo, 1, qtde_ped, preco, total, produtos)
        itens.append(it)

    if itens:
        filiais.append({'filial': razao or 'COSTA AZUL', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                        'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': '',
                        'condPgto': '', 'empresa': 2, 'itens': itens})
    return filiais
