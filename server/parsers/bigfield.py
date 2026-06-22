"""
Parser Supermercado Big Field — formato de pedido de compra proprietário.
Layout: cabeçalho com dados do cliente (CNPJ, endereço), tabela com colunas
SEQ | REFER/CÓDIGO | QTDE | DESCRIÇÃO | EMB KG | QT.EMB | VLR.EMB | DESC% | IPI% | ICMS% | VL.LQ.UNIT | VLR.TOTAL
QTDE = nº de caixas. QT.EMB = kg/cx. VLR.EMB = preço por caixa.
"""
import re, io
import pdfplumber
from perfil import processar_item

__cliente_nome__ = "Big Field"

def _limpa_float(txt):
    try: return float(str(txt).strip().replace('.','').replace(',','.'))
    except: return 0.0

def parse(pdf_bytes, produtos):
    texto = ''
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto += (page.extract_text() or '') + '\n'

    linhas = texto.splitlines()

    def fm(pat):
        m = re.search(pat, texto, re.IGNORECASE)
        return m.group(1).strip() if m else ''

    pedidoNum   = fm(r'Pedido\s+num[:\s]*(\d+)')
    data_pedido = fm(r'Data\s+Pedido[:\s]*(\d{2}/\d{2}/\d{4})')
    data_entrega= fm(r'Data\s+Entrega[:\s]*(\d{2}/\d{2}/\d{4})')
    cond_pgto   = fm(r'Cond\.?Pagto[:\s]*(\S+)')
    cnpj_fat    = fm(r'CNPJ[:\s]*([\d./-]+)')

    # Endereço — primeira linha do cabeçalho
    endereco = ''
    for ln in linhas:
        if re.search(r'AV\.|RUA|EST\.|ESTRADA', ln, re.IGNORECASE):
            endereco = ln.strip()
            break

    ITEM_RE = re.compile(
        r'^\s*(\d{3})\s+'
        r'([\d-]+)\s+'
        r'(\d+)\s+'
        r'(.+?)\s+'
        r'EMB\s+KG\s+'
        r'(\d+)\s+'
        r'([\d.,]+)\s+'
        r'[\d.,]+\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+\s+'
        r'([\d.,]+)\s*$',
        re.IGNORECASE
    )

    itens = []
    for ln in linhas:
        m = ITEM_RE.match(ln)
        if not m: continue
        cod_cli  = int(re.sub(r'[^\d]', '', m.group(2)) or 0)
        nome_raw = m.group(4).strip()
        qtde_cx  = int(m.group(3))
        preco_cx = _limpa_float(m.group(6))
        total    = _limpa_float(m.group(7))
        it = processar_item(cod_cli, nome_raw, 'CX', qtde_cx, qtde_cx, preco_cx, total, produtos)
        itens.append(it)

    if not itens: return []

    return [{
        'filial':      'Central',
        'pedidoNum':   pedidoNum,
        'cnpj':        cnpj_fat,
        'dataPedido':  data_pedido,
        'dataEntrega': data_entrega,
        'condPgto':    cond_pgto,
        'endereco':    endereco,
        'empresa':     2,
        'itens':       itens,
    }]
