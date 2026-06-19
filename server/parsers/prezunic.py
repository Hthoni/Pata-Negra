"""
Parser Prezunic (Cencosud Brasil) — formato de pedido de compra proprietário.
Layout: cabeçalho com dados do fornecedor e faturamento, tabela de itens com
colunas fixas separadas por espaços. Unidade KG (Qt=1) = venda por kg;
unidade CX (Qt=N) = venda por caixa.
"""
import re
import pdfplumber
from perfil import processar_item

CNPJ_RE = re.compile(r'\d{2}\.?\d{3}\.?\d{3}\s*/\s*\d{4}\s*-\s*\d{2}')

def _limpa_float(txt):
    """Converte '3.590,00' -> 3590.0"""
    txt = str(txt).strip().replace('.', '').replace(',', '.')
    try:
        return float(txt)
    except ValueError:
        return 0.0


def parse(pdf_bytes, produtos):
    texto = ''
    with pdfplumber.open(__import__('io').BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto += (page.extract_text() or '') + '\n'

    linhas = texto.splitlines()

    def fm(pat, txt=texto):
        m = re.search(pat, txt, re.IGNORECASE)
        return m.group(1).strip() if m else ''

    # -- Cabeçalho --
    pedidoNum = fm(r'N[uú]mero\s+(\d+)')
    data_emis = fm(r'Data\s+Emiss[aã]o:\s*(\d{2}[./]\d{2}[./]\d{2,4})')
    cond_pgto = fm(r'Cond\.\s*de\s*Pagamento:\s*(.+?)(?:\s+Enc\.|\s{2,}|$)')

    # CNPJ de faturamento (bloco EMITIR NF)
    cnpj_fat = ''
    for ln in linhas:
        if 'CNPJ:' in ln and '39.346' in ln:
            m = CNPJ_RE.search(ln)
            if m:
                cnpj_fat = m.group(0)
                break

    # Endereço de entrega (segundo bloco Endereço:, dentro de EMITIR NF)
    endereco = ''
    count_end = 0
    for ln in linhas:
        if re.search(r'Endere[çc]o:', ln, re.IGNORECASE):
            count_end += 1
            if count_end == 2:
                resto = re.sub(r'Endere[çc]o:\s*', '', ln, flags=re.IGNORECASE).strip()
                if resto:
                    endereco = resto
                break

    # -- Itens --
    ITEM_RE = re.compile(
        r'^\s*(\d{7})\s+'
        r'(.+?)\s+'
        r'([\d.,]+)\s+'
        r'(KG|CX)\s+'
        r'(\d+)\s+'
        r'([\d.,]+)\s+'
        r'([\d.,]+)',
        re.IGNORECASE
    )

    itens = []
    for ln in linhas:
        m = ITEM_RE.match(ln)
        if not m:
            continue

        cod_cli  = int(m.group(1))
        nome_raw = m.group(2).strip()
        quant    = _limpa_float(m.group(3))
        unidade  = m.group(4).upper()
        qt       = int(m.group(5))
        preco    = _limpa_float(m.group(6))
        total    = _limpa_float(m.group(7))

        if unidade == 'CX':
            emb_tipo = 'CX'
            qtde_emb = qt
            qtde     = quant
        else:
            emb_tipo = 'KG'
            qtde_emb = 1
            qtde     = quant

        it = processar_item(
            cod_cli, nome_raw,
            emb_tipo, qtde_emb, qtde,
            preco, total, produtos
        )
        itens.append(it)

    if not itens:
        return []

    return [{
        'filial':      'Central',
        'pedidoNum':   pedidoNum,
        'cnpj':        cnpj_fat,
        'dataEmissao': data_emis,
        'condPgto':    cond_pgto,
        'endereco':    endereco,
        'empresa':     2,
        'itens':       itens,
    }]
