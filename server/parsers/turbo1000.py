"""
Parser Turbo 1000 — sistema DOJÃO, multi-filial.
Layout: cabeçalho com filial/endereço/pedido, tabela com colunas
Item | Cód | Produto | Embal | Qtde | Cód.EAN | Cód.Fab | Valor N.F. | ... | Valor
Qtde = kg totais (KG) ou nº caixas (CX). Valor N.F. = preço por kg ou por cx.

FIX (24/08/2026):
1) Embalagem só aceitava (KG|CX-N) — o pedido nº 154316 trouxe um item
   ("INGREDIENTES P/ FEIJOADA...") com embalagem "UN" (unidade), tipo
   nunca visto nesse parser. Adicionado como terceira opção.
2) Esse mesmo item tem o nome quebrado em 2 linhas no PDF (o
   pdfplumber extrai "INGREDIENTES P/ FEIJOADA PATA NEGRA" numa linha e
   "PCT 500g" na linha seguinte, junto com o resquício de um código EAN
   que também quebrou). O perfil cadastra o nome COMPLETO, incluindo
   "PCT 500g" — sem mesclar, o match falhava.
   Em vez de mesclar sempre (arriscado: no fim da página, a linha
   seguinte pode ser lixo de rodapé, tipo "Condição de Pagamento..."),
   só mescla quando isso resulta num match de VERDADE no perfil — tenta
   o nome puro primeiro; só tenta com a linha seguinte colada (removendo
   um resquício numérico solto no final, resto de EAN quebrado) se o
   nome puro não bateu com nada. Nunca inventa nome, só confirma contra
   o que já está cadastrado.

FIX (25/08/2026):
3) Item UN caía no 'else KG' do emb_tipo — pacotes viravam kg 1:1 (40
   pacotes de "PCT 500g" = 40kg, errado). Corrigido pra repassar
   emb_tipo='UN' de verdade; processar_item converte puxando o peso do
   nome do produto (500g -> 0,5kg/un, fallback 1000g). 40 x 0,5kg = 20kg,
   confere com o pedido real (nº 154316).
"""
import re, io
import pdfplumber
from perfil import processar_item, match_perfil

__cliente_nome__ = "Turbo 1000"

def _limpa_float(txt):
    try: return float(str(txt).strip().replace('.','').replace(',','.'))
    except: return 0.0

ITEM_RE = re.compile(
    r'^\s*(\d+)\s+'
    r'(\d+)\s+'
    r'(.+?)\s+'
    r'(KG|CX-\d+|UN)\s+'
    r'(\d+)\s+'
    r'(?:\d+\s+)?'
    r'(?:\d+\s+)?'
    r'([\d.,]+)\s+'
    r'.+?'
    r'([\d.,]+)\s*$',
    re.IGNORECASE
)


def _mescla_continuacao(nome_atual, prox_linha):
    """Cola a linha seguinte no nome, removendo um resquício numérico
    solto no final (sobra de EAN quebrado em 2 linhas) — só chamada
    quando o nome puro não bateu com nada no perfil (ver parse())."""
    prox_linha = (prox_linha or '').strip()
    if not prox_linha or re.match(r'^\d+\s+\d+\s', prox_linha):
        return None  # linha vazia ou já é o próximo item -> não é continuação
    partes = prox_linha.split()
    while partes and partes[-1].isdigit():
        partes.pop()
    complemento = ' '.join(partes)
    return f'{nome_atual} {complemento}'.strip() if complemento else None


def parse(pdf_bytes, produtos):
    texto = ''
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto += (page.extract_text() or '') + '\n'

    linhas = texto.splitlines()

    def fm(pat):
        m = re.search(pat, texto, re.IGNORECASE)
        return m.group(1).strip() if m else ''

    pedidoNum    = fm(r'Pedido\s+N[oº°]?[:\s]*(\d+)')
    data_pedido  = fm(r'Data\s+Pedido[:\s]*(\d{2}/\d{2}/\d{4})')
    data_entrega = fm(r'Entrega[:\s]*(\d{2}/\d{2}/\d{4})')
    cond_pgto    = fm(r'Condi[çc][aã]o\s+de\s+Pagamento\s*[\n\r]+([^\n\r]+)')
    solicitante  = fm(r'Solicitante[:\s]*(.+?)(?:\s{2,}|$)')
    cnpj_fat     = fm(r'Filial[^\n]*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})')
    filial_nome  = fm(r'Filial[:\s]+(.+?)\s*[-–]\s*\d{2}\.\d{3}')

    endereco = ''
    for i, ln in enumerate(linhas):
        if re.search(r'\bENTREGA\b', ln):
            resto = re.sub(r'ENTREGA', '', ln).strip()
            endereco = resto if resto else (linhas[i+1].strip() if i+1 < len(linhas) else '')
            break

    itens = []
    for i, ln in enumerate(linhas):
        m = ITEM_RE.match(ln)
        if not m: continue
        cod_cli  = int(m.group(2))
        nome_raw = m.group(3).strip()
        emb      = m.group(4).upper()
        qtde     = int(m.group(5))
        preco    = _limpa_float(m.group(6))
        total    = _limpa_float(m.group(7))
        # emb_tipo: CX = qtde de caixas (kg = qtde x kgCx); UN = qtde de
        # pacotes/unidades (processar_item extrai o peso do nome, ex. "500g",
        # fallback 1000g); senão KG (qtde já é kg). BUG (25/08/2026): UN
        # estava caindo no 'else KG', tratando pacotes como se já fossem kg
        # (40 pacotes de 500g virava 40kg em vez de 20kg).
        emb_tipo = 'CX' if emb.startswith('CX') else ('UN' if emb == 'UN' else 'KG')

        # nome puro não bateu -> tenta mesclar com a linha seguinte (nome
        # quebrado em 2 linhas), só usa se isso resultar num match real
        if not match_perfil(nome_raw, produtos):
            prox = linhas[i + 1] if i + 1 < len(linhas) else ''
            candidato = _mescla_continuacao(nome_raw, prox)
            if candidato and match_perfil(candidato, produtos):
                nome_raw = candidato

        it = processar_item(cod_cli, nome_raw, emb_tipo, qtde, qtde, preco, total, produtos)
        itens.append(it)

    if not itens: return []

    return [{
        'filial':      filial_nome,
        'pedidoNum':   pedidoNum,
        'cnpj':        cnpj_fat,
        'dataPedido':  data_pedido,
        'dataEntrega': data_entrega,
        'condPgto':    cond_pgto,
        'solicitante': solicitante,
        'endereco':    endereco,
        'empresa':     2,
        'itens':       itens,
    }]
