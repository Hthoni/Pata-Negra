"""
Parser Mercado Summer — formato proprietário (rptPedido.rdlc).
Layout: 1 filial por PDF, cabeçalho no topo, linhas de itens com colunas
separadas por múltiplos espaços. Unidade sempre KG no PDF, exceto quando
a embalagem começa com KG-N (indica venda por caixa de N kg).
"""
import re
import pdfplumber
from perfil import processar_item

CNPJ_RE = re.compile(r'\d{2}\.?\d{3}\.?\d{3}\s*/\s*\d{4}\s*-\s*\d{2}')


_SINONIMOS_SUMMER = [
    (re.compile(r'\bRABO\b', re.IGNORECASE), 'RABINHO'),
]

_SUFIXOS_GENERICOS = re.compile(
    r'\b(?:PATA\s+NEGRA|PCT|SUINO|SALG|SALGADO|SALGADA|DEFUMADO|DEFUMADA)\b',
    re.IGNORECASE
)

def _nome_limpo(nome):
    """Remove termos genéricos e expande sinônimos Summer-específicos
    pra melhorar o fuzzy matching com o Perfil."""
    s = nome
    for pat, sub in _SINONIMOS_SUMMER:
        s = pat.sub(sub, s)
    s = _SUFIXOS_GENERICOS.sub(' ', s)
    return re.sub(r'\s{2,}', ' ', s).strip()


def _limpa_float(txt):
    """Converte '24.408,00' → 24408.0"""
    txt = txt.strip().replace('.', '').replace(',', '.')
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
        m = re.search(pat, txt)
        return m.group(1).strip() if m else ''

    # ── Cabeçalho ──────────────────────────────────────────────────────────
    pedidoNum = fm(r'Pedido\s+N[oº°]:\s*(\d+)')
    data_pedido = fm(r'Data\s+Pedido:\s*(\d{2}/\d{2}/\d{4})')

    # CNPJ da filial vem na linha "Filial: Fil XX CD - 12.968.606/0004-47"
    cnpj_raw = ''
    filial_nome = ''
    for ln in linhas:
        if ln.startswith('Filial:'):
            m = CNPJ_RE.search(ln)
            if m:
                cnpj_raw = m.group(0)
            # nome da filial: texto entre "Fil XX CD -" e o CNPJ
            mf = re.search(r'Filial:\s*(.*?)\s*(?:' + re.escape(cnpj_raw) + r')', ln)
            if mf:
                filial_nome = mf.group(1).strip().rstrip('-').strip()
            break

    solicitante = fm(r'Solicitante:\s*(.+)')
    cond_pgto   = fm(r'(\d+\s+dias)')

    # ── Itens ──────────────────────────────────────────────────────────────
    # Cada linha de item começa com um número de sequência, depois o cód,
    # depois o nome (com sufixo " kg"), depois a embalagem (KG ou KG-N),
    # depois a qtde, depois o cód EAN, depois o cód fab, depois valores.
    #
    # Exemplo real:
    # "1 36205 BACON DEFUMADO PATA NEGRA PCT kg KG-20 36 1213 678,00 0 678,00 43,46 24.408,00"
    # "3 34515 COSTELA SUINO SALG PATA NEGRA kg KG 140 34516 26,90 0 26,90 26,90 3.766,00"
    #
    # Regex: seq  cod  nome(+kg)  embalagem  qtde  codEAN  (codFab)  valorNF  %desc  *IPI  %ST  custoNF  unitario  (bon)  valor
    ITEM_RE = re.compile(
        r'^\s*(\d+)\s+'           # seq
        r'(\d{4,6})\s+'           # cód produto
        r'(.+?)\s+kg\s+'          # nome produto (termina com " kg")
        r'(KG-?\d*)\s+'           # embalagem: KG, KG-10, KG-20 etc
        r'(\d+)\s+'               # qtde
        r'\d+\s+'                 # cód EAN
        r'(?:\d+\s+)?'            # cód fab (opcional)
        r'([\d.,]+)\s+'           # valor NF
        r'(?:\d+\s+)'             # % desc
        r'([\d.,]+)',              # custo NF / unitário
        re.IGNORECASE
    )

    itens = []
    for ln in linhas:
        m = ITEM_RE.match(ln)
        if not m:
            continue
        cod_cli  = int(m.group(2))
        nome_raw = m.group(3).strip()
        emb_str  = m.group(4).upper()   # KG, KG-20, KG-10 …
        qtde     = float(m.group(5))
        preco    = _limpa_float(m.group(7))
        total    = _limpa_float(m.group(7)) * qtde  # recalculado — PDF pode ter arredondamento

        # Se embalagem é KG-N (ex: KG-20), é venda por caixa de N kg
        # Se embalagem é simplesmente KG, é venda por kg
        if re.match(r'KG-\d+', emb_str):
            emb_tipo = 'CX'
            n = int(re.search(r'\d+', emb_str).group())
            qtde_emb = n    # kg por caixa
        else:
            emb_tipo = 'KG'
            qtde_emb = 1

        it = processar_item(cod_cli, _nome_limpo(nome_raw), emb_tipo, qtde_emb, qtde, preco, total, produtos)
        itens.append(it)

    if not itens:
        return []

    return [{
        'filial':     filial_nome,
        'pedidoNum':  pedidoNum,
        'cnpj':       cnpj_raw,
        'dataPedido': data_pedido,
        'condPgto':   cond_pgto,
        'solicitante': solicitante,
        'empresa':    2,   # Pata Negra Distribuidora (CNPJ ...0001-90 no PDF)
        'itens':      itens,
    }]
