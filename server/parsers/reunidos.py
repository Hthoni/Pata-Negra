"""
Parser Reunidos (Reunidos Supermercados) — formato TOTVS
"Pedido com Verba Comercial" (rptPedidoVerbaComercial).

Pedido de 1 filial, pode ter 2 páginas (itens continuam na pág. 2, totais
no fim). Junta todas as páginas num texto só.

Layout da linha de item (a coluna Embal. traz "CX-N"):
  {Item} {Produto [unidade kg/400g]} CX-N [{Fabr}] {Qtde} {ValorNF} {%Desc}
  {CustoNF} {Unitário} {Verba} {Final} {Valor}

Robustez:
 - âncora = "CX-N" (marca a coluna Embal.); antes vem Item+Produto, depois
   os números da linha;
 - a coluna Fabr. às vezes traz um código (ex.: 3220) e às vezes vem vazia —
   por isso os campos são ancorados pela DIREITA (as colunas finais estão
   sempre preenchidas): Valor=-1, Final=-2, Verba=-3, Unitário=-4,
   CustoNF=-5, %Desc=-6, ValorNF=-7, Qtde=-8;
 - Qtde = nº de CAIXAS -> emb_tipo='CX' -> kg = Qtde x kgCx (do Perfil).
   Confere: soma dos itens = SUBTOTAL do pedido e soma das Qtdes = "Itens
   Quantidade" do rodapé.

Empresa (Indústria x Distribuidora) vem por produto do Perfil (coluna A);
o fornecedor do pedido é "PATA NEGRA DISTRIBUIDORA" (default 2).
CNPJ da filial no cabeçalho ("Filial:... - NN.NNN.NNN/NNNN-NN"); o main.py
casa contra a tabela de filiais (M:T) do Perfil p/ enriquecer
nome/número/região/lat/lng.
"""

__cliente_nome__ = "Reunidos"

import io
import re
import pdfplumber
from perfil import processar_item


def _num(s):
    return float(str(s).replace('.', '').replace(',', '.'))


# número BR: "15.382,40" | "704,00" | "5" | "3220"
_RE_NUM = re.compile(r'[\d.]+,\d+|\d+')
_RE_ITEM = re.compile(r'^\s*(\d{1,3})\s+(.+?)\s+(CX-\d+)\s+(.+?)\s*$', re.M)


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    def fm(pat):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'Pedido\s*N[ºo°:]*\s*(\d+)')
    dataPedido = fm(r'Pedido\s*N[ºo°:]*\s*\d+\s+([\d/]+)')
    dataEntrega = fm(r'Entrega:\s*([\d/]+)')
    prazo = fm(r'Pagamento:\s*(\d+)')
    condPgto = f'{prazo} dias' if prazo else ''

    # CNPJ da filial: "Filial:RN COLUBANDE - 32.352.751/0001-63"
    cnpj = fm(r'Filial:.*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})')
    filial_nome = fm(r'Filial:\s*(.+?)\s*-\s*\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')

    end_m = re.search(r'ENTREGA\s+(.+?)\s+Brasil', txt, re.S | re.I)
    endereco = re.sub(r'\s+', ' ', end_m.group(1)).strip() if end_m else ''

    itens = []
    for m in _RE_ITEM.finditer(txt):
        nome = m.group(2).strip()
        tail = m.group(4)
        nums = _RE_NUM.findall(tail)
        if len(nums) < 8:
            continue  # não é linha de item (cabeçalho / rodapé)
        qtde = _num(nums[-8])          # nº de caixas
        preco = _num(nums[-4])         # Unitário (R$/kg; p/ LING é R$/pct)
        total = _num(nums[-1])         # Valor total da linha
        # remove a unidade que fecha o nome (" kg"), preserva pesos ("400g")
        nome = re.sub(r'\s+kg$', '', nome, flags=re.I).strip()
        it = processar_item('', nome, 'CX', 1, qtde, preco, total, produtos)
        itens.append(it)

    if not itens:
        return []
    return [{
        'filial': filial_nome or 'REUNIDOS',
        'pedidoNum': pedidoNum,
        'cnpj': cnpj,
        'endereco': endereco,
        'dataPedido': dataPedido,
        'dataEntrega': dataEntrega,
        'condPgto': condPgto,
        'empresa': 2,
        'itens': itens,
    }]
