"""
Parser Mercado Summer — formato proprietário (rptPedido.rdlc).
Layout: 1 filial por PDF, cabeçalho no topo, linhas de itens com colunas
separadas por múltiplos espaços. Unidade sempre KG no PDF, exceto quando
a embalagem começa com KG-N (indica venda por caixa de N kg).

FIX (10/08/2026): matching de produto quebrado desde que o sistema passou
a exigir nome EXATO (perfil.py, match_perfil) em vez de aproximado. Duas
causas, as duas neste parser:
  1) A regex cortava o nome ANTES do "kg" (ex.: "CHISPE SUINO SALG PATA
     NEGRA"), mas o perfil cadastra o nome COM "kg" no final (ex.:
     "CHISPE SUINO SALG PATA NEGRA kg") — nunca batia exato. Corrigido
     capturando o nome JUNTO com o "kg" minúsculo (que é parte do nome
     de exibição do produto, diferente do "KG"/"KG-N" maiúsculo da coluna
     Embalagem que vem logo depois).
  2) _nome_limpo() removia palavras genéricas (SUINO, SALG, PATA NEGRA)
     pra "ajudar" o matching aproximado antigo — mas o perfil MANTÉM essas
     palavras no nome cadastrado, então a limpeza afastava ainda mais do
     nome exato esperado. Removida (junto com o sinônimo RABO->RABINHO,
     que também nunca bateu: o perfil cadastra "RABO", não "RABINHO").
"""

__cliente_nome__ = "Mercado Summer"

import re
import pdfplumber
from perfil import processar_item

CNPJ_RE = re.compile(r'\d{2}\.?\d{3}\.?\d{3}\s*/\s*\d{4}\s*-\s*\d{2}')


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

    # Endereço de entrega: "ENTREGA  Av. Carlos Marighella, 7974..."
    endereco = ''
    capturando = False
    for ln in linhas:
        if ln.strip().startswith('ENTREGA'):
            resto = re.sub(r'^ENTREGA\s*', '', ln).strip()
            if resto:
                endereco = resto
            else:
                capturando = True
        elif capturando:
            endereco = ln.strip()
            capturando = False
        if endereco:
            break

    # CNPJ da filial vem na linha "Filial: Fil XX CD - 12.968.606/0004-47"
    cnpj_raw = ''
    filial_nome = ''
    for ln in linhas:
        if ln.startswith('Filial:'):
            m = CNPJ_RE.search(ln)
            if m:
                cnpj_raw = m.group(0)
            mf = re.search(r'Filial:\s*(.*?)\s*(?:' + re.escape(cnpj_raw) + r')', ln)
            if mf:
                filial_nome = mf.group(1).strip().rstrip('-').strip()
            break

    solicitante = fm(r'Solicitante:\s*(.+)')
    cond_pgto   = fm(r'(\d+\s+dias)')

    # ── Itens ──────────────────────────────────────────────────────────────
    # FIX: nome agora capturado JUNTO com o "kg" minúsculo final (parte do
    # nome de exibição), separado da coluna Embalagem (KG/KG-N, maiúsculo)
    # que vem logo depois.
    ITEM_RE = re.compile(
        r'^\s*(\d+)\s+'           # seq
        r'(\d{4,6})\s+'           # cód produto
        r'(.+?\s+kg)\s+'          # nome produto, incluindo o "kg" final
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

        if re.match(r'KG-\d+', emb_str):
            emb_tipo = 'CX'
            n = int(re.search(r'\d+', emb_str).group())
            qtde_emb = n    # kg por caixa
        else:
            emb_tipo = 'KG'
            qtde_emb = 1

        it = processar_item(cod_cli, nome_raw, emb_tipo, qtde_emb, qtde, preco, total, produtos)
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
        'endereco':   endereco,
        'empresa':    2,   # Pata Negra Distribuidora (CNPJ ...0001-90 no PDF)
        'itens':      itens,
    }]
