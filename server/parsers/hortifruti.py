"""
Parser Hortifruti (Neogrid) — formato "PEDIDO DE COMPRA" da Neogrid, usado
pelas redes Americanas / Hortifruti (Natural da Terra).

Layout de tabela; descrições quebram em varias linhas. Usa extract_tables():
cada item vem como um campo (possivelmente multi-linha) do tipo
   [desc-A]\\n[N?] [codigo] [desc-B] (Quilograma|Unidade) [8 numeros]\\n[desc-C]
As vezes o N e o codigo se fundem (ex.: '1,007898611040613'); por isso a
descricao e reconstruida por TOKENS (descartando os puramente numericos) e a
quantidade/preco sao lidos DEPOIS da unidade de medida.

Unidade:
 - 'Quilograma' -> Qtde Pedida ja em KG.
 - 'Unidade'    -> Qtde Pedida em unidades; converte pelo peso no nome
                   (ex.: 500G -> 0,5 kg/un, 400G -> 0,4 kg/un).

Empresa (faturamento): detectada pelo CNPJ do Fornecedor:
 - 10.171.633 -> Industria  (empresa 1)
 - 56.423.719 -> Distribuidora (empresa 2)
"""

__cliente_nome__ = "Hortifruti"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

CNPJ_INDUSTRIA = '10.171.633'
CNPJ_DISTRIBUIDORA = '56.423.719'
_UNID = re.compile(r'\b(Quilograma|Unidade)\b', re.I)
_NUM = re.compile(r'\d[\d.,]*')


def _num(s):
    return float((s or '0').replace('.', '').replace(',', '.'))


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
        tabelas = []
        for pg in pdf.pages:
            tabelas.extend(pg.extract_tables() or [])

    def fm(pat, txt=full_text):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'Número Pedido:\s*(\d+)')
    dataPedido = fm(r'Data de Emissão:\s*([\d/ ]+)').replace(' ', '')
    dataEntrega = fm(r'Data Inicial:\s*([\d/ ]+)').replace(' ', '')
    cnpj = fm(r'CNPJ do Local de Entrega:\s*([\d./ \-]+)').replace(' ', '')

    # Empresa pela origem de faturamento (2o CNPJ da linha Comprador/Fornecedor).
    cnpjs_linha = re.search(r'CNPJ:\s*([\d./ \-]+)\s+CNPJ:\s*([\d./ \-]+)', full_text)
    forn = (cnpjs_linha.group(2) if cnpjs_linha else '').replace(' ', '')
    empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in forn.replace('.', '').replace('-', '') else 2

    itens = []
    for t in tabelas:
        for row in t:
            for cell in row:
                if not cell:
                    continue
                mu = _UNID.search(cell)
                if not mu:
                    continue
                linha = cell.replace('\n', ' ').strip()
                mu = _UNID.search(linha)
                if not mu:
                    continue
                un = mu.group(1)
                antes = linha[:mu.start()].strip()
                depois = linha[mu.end():].strip()

                # separa as COLUNAS numéricas (logo após a unidade) da desc-C:
                # os tokens numéricos iniciais são as colunas; o resto é descrição.
                dtoks = depois.split()
                nums, resto = [], []
                for i, w in enumerate(dtoks):
                    if not resto and re.match(r'^[\d.,]+$', w):
                        nums.append(w)
                    else:
                        resto.append(w)
                if len(nums) < 6:
                    continue
                qped = _num(nums[1])                       # Qtde Pedida
                preco = _num(nums[4]) if len(nums) >= 5 else 0.0   # Preço Líquido
                total = _num(nums[7]) if len(nums) >= 8 else _num(nums[-1])  # Valor Total
                descC = ' '.join(resto).strip()

                # desc-A/B = tokens não-puramente-numéricos do "antes"
                toks = [w for w in antes.split() if not re.match(r'^[\d.,]+$', w)]
                nome = ' '.join(toks + ([descC] if descC else [])).strip()
                nome = re.sub(r'\s+', ' ', nome)

                if un.lower().startswith('quilog'):
                    kg = qped
                else:
                    mg = re.search(r'(\d+)\s*G\b', nome, re.I)
                    gramas = int(mg.group(1)) if mg else 1000
                    kg = qped * gramas / 1000.0

                it = processar_item('', nome, 'KG', 1, kg, preco, total, produtos)
                it['empresa'] = empresa
                itens.append(it)

    if not itens:
        return []
    return [{
        'filial': 'HORTIFRUTI', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
        'endereco': '', 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
        'condPgto': '', 'empresa': empresa, 'itens': itens
    }]
