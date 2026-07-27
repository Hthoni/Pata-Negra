"""
Parser Imbatível (Imbatível Supermercados — Niterói/RJ). Formato ERP Superus
"Pedido de Compra". Um pedido por PDF (cliente com uma filial).

Layout do item (uma linha por produto):
  {CodImbativel} {RefForn} KG {Descrição...} {Qtd} {TipoEmb}{TamCx} KG
  {Decto} {Bonif} {Pr.Unit} {Pr.Emb} {Vlr.Total}
Observações do layout (extração via pdfplumber cola alguns tokens):
  - Embalagem sai como "CX10 KG" / "KG10 KG": Tipo (CX|KG) + tamanho da caixa
    (kg/cx) grudados. TamCx = 10 nestes pedidos.
  - Decto e Bonif podem vir grudados ("0,000,000"); por isso lê-se o VALOR pela
    ÚLTIMA coluna numérica da linha (Vlr. Total), que é a autoritativa.

Conversão p/ KG (faturamento em KG):
  A coluna Qtd é em CAIXAS; kg = Qtd * TamCx (10). O "Pr. Unit" impresso é o
  preço de TABELA (~2% acima); o preço líquido real = Vlr.Total ÷ kg. Ex.:
  CHISPE 5 cx * 10 = 50 kg, Vlr 530,00 -> 10,60/kg. A soma dos Vlr. Total
  confere com "Valor Total Produtos" do rodapé. Passa emb_tipo='KG' p/ o
  processar_item NÃO remultiplicar por kgCx.

EMPRESA: pelo fornecedor do pedido.
  CNPJ 10.171.633 (INDUSTRIA DE ALIMENTOS PATA NEGRA) -> empresa 1 (Indústria)
  CNPJ 56.423.719 (PATA NEGRA DISTRIBUIDORA)          -> empresa 2 (Distribuidora)
Bate com o Fat do perfil (produtos Imbatível são Fat 1 / Indústria).

Filial: cliente de uma unidade (CNPJ 28.480.886/0001-37, Niterói). O main.py
enriquece região/lat/lng pelo CNPJ contra a tabela M:T do perfil. Match de
produto por NOME no perfil (col C); devolve o Cód. Interno (col B).
"""

__cliente_nome__ = "Imbativel"

import io
import re
import pdfplumber
from perfil import processar_item

CNPJ_INDUSTRIA = '10171633'
CNPJ_DISTRIBUIDORA = '56423719'

_RE_NUM = re.compile(r'\d[\d.]*,\d+|\d+')
# item: Cod, RefForn, KG, Desc, Qtd, (CX|KG)+TamCx, KG, resto(nºs; Vlr=último)
_RE_ITEM = re.compile(
    r'^\s*(\d+)\s+(\d+)\s+KG\s+(.+?)\s+(\d+)\s+(CX|KG)\s*(\d+)\s+KG\s+(.+)$', re.M)


def _num(s):
    return float(str(s).replace('.', '').replace(',', '.'))


def _digitos(s):
    return re.sub(r'\D', '', s or '')


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    def fm(pat):
        m = re.search(pat, full, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'N[ºo°]\s*Pedido:\s*(\d+)')
    dataPedido = fm(r'N[ºo°]\s*Pedido:.*?Data:\s*([\d/]+)') or fm(r'Emiss[aã]o:\s*([\d/]+)')
    dataEntrega = fm(r'Data\s*(?:de\s*)?[Ee]ntrega:\s*([\d/]+)')
    prazo = fm(r'Prazo pagamento:\s*(\d+)\s*DIAS') or fm(r'(\d+)\s*DIAS')
    condPgto = f'{prazo} dias' if prazo else ''
    frete = 'CIF' if re.search(r'\bCIF\b', full) else ('FOB' if re.search(r'\bFOB\b', full) else '')

    # empresa pelo fornecedor
    empresa = 1
    if CNPJ_DISTRIBUIDORA in _digitos(full) and 'DISTRIBUIDORA' in full.upper():
        empresa = 2
    elif re.search(r'DISTRIBUIDORA', full, re.I) and CNPJ_INDUSTRIA not in _digitos(full):
        empresa = 2

    # cliente = CNPJ que não é da Pata Negra
    cnpj = ''
    for c in re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', full):
        if _digitos(c)[:8] not in (CNPJ_INDUSTRIA, CNPJ_DISTRIBUIDORA):
            cnpj = c
            break
    filial_nome = fm(r'Raz[aã]o Social:\s*([A-Z].+?)\s+CNPJ:') or 'IMBATIVEL'
    endereco = fm(r'Endere[çc]o:\s*(EST[^\n]+?)\s*(?:Telefone|Fax|$)') or fm(r'(EST FRANCISCO[^\n]+)')

    itens = []
    for m in _RE_ITEM.finditer(full):
        refforn = m.group(2)
        nome = re.sub(r'\s+', ' ', m.group(3)).strip()
        qtde_cx = _num(m.group(4))
        tam_cx = _num(m.group(6))
        nums = _RE_NUM.findall(m.group(7))
        if not nums:
            continue
        total = _num(nums[-1])                 # Vlr. Total (coluna final)
        kg = qtde_cx * tam_cx if tam_cx else qtde_cx
        preco = round(total / kg, 4) if kg else 0.0
        it = processar_item(refforn, nome, 'KG', 1, kg, preco, total, produtos)
        it['empresa'] = empresa
        it['qtdeCaixas'] = qtde_cx
        itens.append(it)

    if not itens:
        return []

    return [{
        'filial': filial_nome,
        'pedidoNum': pedidoNum,
        'cnpj': cnpj,
        'endereco': endereco,
        'dataPedido': dataPedido,
        'dataEntrega': dataEntrega,
        'condPgto': condPgto,
        'frete': frete,
        'empresa': empresa,
        'itens': itens,
    }]


def debug_layout(pdf_bytes, n=80):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pg in enumerate(pdf.pages, 1):
            print(f'--- pagina {i} ---')
            print('\n'.join((pg.extract_text() or '').splitlines()[:n]))
