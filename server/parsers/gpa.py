"""
Parser GPA (Grupo Pão de Açúcar / Cia Brasileira de Distribuição) — formato
TOTVS "Pedido Mercadoria" (EDI GPA). Cada PDF é o pedido de UM fornecedor
Pata Negra para o CD 0409 (Duque de Caxias).

Particularidades do layout:
 - A descrição do produto quebra em várias linhas e o extract_text() a embaralha
   com a coluna do código. Por isso o NOME é reconstruído por POSIÇÃO das
   palavras: cada palavra da coluna de descrição é atribuída ao código
   verticalmente mais próximo (extract_words + vizinho mais próximo em 'top').
 - A linha do código carrega Emb., Qtde. e Custo Merc.:
       {codEAN} [frag. nome] {Emb} {Qtde} {Custo} 0,00 ...
   Emb=1 -> item vendido por KG (Qtde já em kg).
   Emb>1 -> item em caixa (Qtde = nº de caixas); kg = Qtde * Emb * peso_un,
            com peso_un lido do nome (500G, 400GR).
   O total da linha = Qtde * Custo (confere com "Tot. Pedido" do rodapé).

 - EMPRESA (faturamento) vem do FORNECEDOR do pedido (confiável, um por PDF):
       CNPJ 10.171.633 (39518 IND ALIMENTOS PATA NEGRA) -> empresa 1 (Indústria)
       CNPJ 56.423.719 (61891 PATA NEGRA DIST)          -> empresa 2 (Distribuidora)
   Bate com a coluna Fat do perfil (linguiça=1, demais=2).

 - Filial = local de entrega (CD 0409, CNPJ 47.508.411/2519-06). O main.py
   enriquece nome/região/lat/lng pelo CNPJ contra a tabela M:T do perfil.

Segue o mesmo contrato dos outros parsers: parse(pdf_bytes, produtos) ->
lista de pedidos; itens montados via perfil.processar_item, que casa por NOME
no perfil (col C) e devolve o Cód. Interno (col B). Converte tudo para KG e
passa emb_tipo='KG' para o processar_item NÃO remultiplicar por kgCx.
"""

__cliente_nome__ = "GPA"

import io
import re
import pdfplumber
from perfil import processar_item

CNPJ_INDUSTRIA = '10171633'
CNPJ_DISTRIBUIDORA = '56423719'

_COD = re.compile(r'^\d{13,14}$')
_PESO_UN = re.compile(r'(\d+)\s*GR?\b', re.I)   # 500G, 400GR -> gramas por unidade
_X_NOME_MAX = 200                                # coluna de descrição fica à esquerda


def _num(s):
    return float(str(s).replace('.', '').replace(',', '.'))


def _digitos(s):
    return re.sub(r'\D', '', s or '')


def _reconstruir_nomes(page):
    """Devolve {top_do_codigo: nome} reconstruindo a descrição por posição:
    cada palavra alfabética da coluna de descrição vai para o código mais
    próximo verticalmente."""
    words = page.extract_words()
    codes = [w for w in words if _COD.match(w['text'])]
    if not codes:
        return {}, []
    code_tops = sorted(c['top'] for c in codes)
    lo, hi = code_tops[0] - 25, code_tops[-1] + 35

    buckets = {t: [] for t in code_tops}
    for w in words:
        if w['x0'] >= _X_NOME_MAX:
            continue
        if not re.search(r'[A-Za-z]', w['text']):
            continue
        if not (lo <= w['top'] <= hi):
            continue
        # código mais próximo em 'top'
        alvo = min(code_tops, key=lambda t: abs(t - w['top']))
        buckets[alvo].append((w['top'], w['x0'], w['text']))

    nomes = {}
    for t, ws in buckets.items():
        ws.sort(key=lambda x: (round(x[0]), x[1]))
        nomes[t] = re.sub(r'\s+', ' ', ' '.join(x[2] for x in ws)).strip()
    return nomes, codes


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full = '\n'.join(p.extract_text() or '' for p in pdf.pages)
        pages = pdf.pages
        nomes, codes = {}, []
        for pg in pages:
            n, c = _reconstruir_nomes(pg)
            nomes.update(n)
            codes.extend(c)

    def fm(pat, txt=full):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'\n\s*(\d{6,})\s+\d+\s*DD')                      # 32856010409
    if not pedidoNum:
        pedidoNum = fm(r'Pedido\s*N[ºo°]?\s*\n?\s*(\d{6,})')
    dataPedido = fm(r'Emiss[aã]o:\s*([\d/]+)')
    dataEntrega = fm(r'\d+\s*DD[^\n]*?([\d]{2}/[\d]{2}/[\d]{4})')     # data no fim da linha do cabeçalho
    prazo = fm(r'(\d+)\s*DD')
    condPgto = f'{prazo} dias' if prazo else ''
    frete = 'CIF' if re.search(r'\bCIF\b', full) else ('FOB' if re.search(r'\bFOB\b', full) else '')

    # fornecedor -> empresa
    forn = fm(r'Fornecedor:\s*(\d+)')
    cnpjs = re.findall(r'CGC:\s*([\d./\-]+)', full)
    forn_cnpj = _digitos(cnpjs[0]) if cnpjs else ''
    empresa = 1 if forn_cnpj.startswith(CNPJ_INDUSTRIA) else 2 if forn_cnpj.startswith(CNPJ_DISTRIBUIDORA) else 2

    # filial = local de entrega
    filial_nome = fm(r'Local de entrega:\s*([^\n]+)')
    entrega_cnpj_raw = cnpjs[1] if len(cnpjs) > 1 else ''
    entrega_cnpj = _digitos(entrega_cnpj_raw)
    endereco = fm(r'(ROD\s+WASHINGTON[^\n]+)')

    itens = []
    for code in sorted(codes, key=lambda c: c['top']):
        # linha do código no texto -> Emb, Qtde, Custo
        linha = next((ln for ln in full.splitlines() if code['text'] in ln), '')
        depois = linha.split(code['text'], 1)[1] if code['text'] in linha else ''
        nums = re.findall(r'\d+(?:,\d+)?', depois)
        if len(nums) < 3:
            continue
        emb = int(_num(nums[0]))
        qtde = _num(nums[1])
        custo = _num(nums[2])
        nome = nomes.get(code['top'], '').strip()

        total = round(qtde * custo, 2)
        if emb <= 1:
            kg = qtde                          # já em kg
        else:
            mg = _PESO_UN.search(nome)
            peso_un = int(mg.group(1)) / 1000.0 if mg else 0.0
            kg = qtde * emb * peso_un if peso_un else qtde
        preco_kg = round(total / kg, 4) if kg else 0.0

        it = processar_item(code['text'], nome, 'KG', 1, kg, preco_kg, total, produtos)
        it['empresa'] = empresa
        it['codEAN'] = code['text']
        it['qtdeCaixas'] = qtde if emb > 1 else None
        itens.append(it)

    if not itens:
        return []

    return [{
        'filial': filial_nome or 'GPA CD DUQUE',
        'pedidoNum': pedidoNum,
        'cnpj': entrega_cnpj_raw or entrega_cnpj,
        'endereco': endereco,
        'dataPedido': dataPedido,
        'dataEntrega': dataEntrega,
        'condPgto': condPgto,
        'frete': frete,
        'fornecedor': forn,
        'empresa': empresa,
        'itens': itens,
    }]


def debug_layout(pdf_bytes, n=60):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pg in enumerate(pdf.pages, 1):
            print(f'--- pagina {i} ---')
            print('\n'.join((pg.extract_text() or '').splitlines()[:n]))
