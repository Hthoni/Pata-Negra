"""
Parser Assaí — formato Consinco/TOTVS, multi-filial (várias lojas no mesmo
PDF, uma seção por loja) ou uma página por filial.
Código do produto vem colado ao nome (ex: '1156510BACON LOMBO...'), exceto
em algumas linhas onde vem com espaço — o item-regex já tolera os dois
casos via backtracking.

FIX (11/08/2026): nome da filial não era extraído pra lojas cujo código
"LJXXX" vem com espaço no meio ("LJ 219" em vez de "LJ219", visto na loja
D Caxias Pq Fluminense) — mesmo tipo de artefato de kerning do pdfplumber
que o CNPJ_RE já tolerava (\\s* ao redor de / e -), só faltava aplicar a
mesma tolerância aqui. Sem o nome batendo, a filial caía no fallback
genérico 'ASSAÍ' — os itens em si processavam normal, só o nome da loja
ficava errado no romaneio/PDF/mapa.

FIX (03/09/2026): removida a limpeza de sufixo "FRAC KG"/"KG" do nome
do produto (re.sub que existia aqui antes). O perfil atual do Assaí já
padronizou os nomes COM esse sufixo incluído (ex.: "BACON PATA NEGRA
FRAC KG" é o nome de verdade, não "BACON PATA NEGRA" + unidade solta)
— a limpeza antiga, que fazia sentido numa versão anterior do perfil
sem esse sufixo, virou o oposto do que precisa agora e quebrava 134 de
150 itens de um pedido real (22322611/L, 03/09/2026) com "Produto não
cadastrado", mesmo o produto estando cadastrado — só que sob o nome
COM "FRAC KG", que a limpeza arrancava antes de comparar. Sem a
limpeza, os 150 itens batem 100%.
"""

__cliente_nome__ = "Assaí"

import re
import pdfplumber
from perfil import processar_item

CNPJ_INDUSTRIA = '10.171.633'
CNPJ_RE = r'\d{2}\.\d{3}\.\d{3}\s*/\s*\d{4}\s*-\s*\d{2}'  # tolera espaço ao redor de / e - (artefato do pdfplumber)


def parse(pdf_bytes, produtos):
    import io
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    # Divide por "PEDIDO DE COMPRAS" no texto TODO (não por página isolada).
    # FIX (26/08/2026): "1 página por filial" era só coincidência do layout
    # curto -- pedido grande que estourasse de página cortava itens da loja
    # de verdade E criava uma filial fantasma 'ASSAÍ' (sem CNPJ/endereço,
    # nome genérico) com os itens que sobraram na página seguinte.
    blocos = re.split(r'(?=PEDIDO DE COMPRAS\s+\S+)', full)
    for txt in blocos:
        if 'PEDIDO DE COMPRAS' not in txt:
            continue

        def fm(pat, txt=txt):
            m = re.search(pat, txt, re.I)
            return m.group(1).strip() if m else ''

        pedidoNum = fm(r'PEDIDO DE COMPRAS\s+(\S+)')
        cnpj_m = re.search(rf'CNPJ\s+({CNPJ_RE})\s+Cidade.*?CNPJ\s+({CNPJ_RE})', txt, re.S)
        cnpj_forn = cnpj_m.group(1) if cnpj_m else ''
        cnpj_loja = cnpj_m.group(2) if cnpj_m else ''
        empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj_forn.replace('.', '') else 2

        # FIX: LJ\s*\d+ tolera "LJ219" (comum) e "LJ 219" (kerning) igual
        filial_m = re.search(r'R\. Social SENDAS.*?LJ\s*\d+\s+\d+\s+(.+?)$', txt, re.M)
        filial = filial_m.group(1).strip() if filial_m else 'ASSAÍ'

        end_m = re.search(r'ENDEREÇO PARA ENTREGA.*?Endereço\s+(.+?)\s+Endereço', txt, re.S)
        endereco = end_m.group(1).strip() if end_m else ''
        dataPedido = fm(r'Data da emiss[aã]o\s+([\d/]+)')
        dataEntrega = fm(r'Previs[aã]o de entrega\s+([\d/]+)')
        cond_m2 = re.search(r'pagamento\s+(\d+)\s*\(', txt)
        condPgto = cond_m2.group(1) + ' dias' if cond_m2 else ''

        reItem = re.compile(
            r'^(\d{7})([A-Z][^\n]+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)',
            re.M)
        itens = []
        for m in reItem.finditer(txt):
            nome_raw = re.sub(r'\s+', ' ', m.group(2)).strip()
            qtde_ped = float(m.group(5).replace('.', '').replace(',', '.'))
            preco = float(m.group(6).replace('.', '').replace(',', '.'))
            total = float(m.group(7).replace('.', '').replace(',', '.'))
            it = processar_item(m.group(1), nome_raw, m.group(3),
                                 int(m.group(4)), qtde_ped, preco, total, produtos)
            it['empresa'] = empresa
            itens.append(it)

        if itens:
            filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj_loja,
                             'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                             'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
