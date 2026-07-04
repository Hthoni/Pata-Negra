"""
Tabela mestra de produtos: de-para código interno -> nome MASTER (nome
oficial Pata Negra). Uma planilha (MASTER.xlsx) no bucket de perfis,
editável pelo usuário.

Layout: cabeçalho na linha 1; a partir da linha 2, coluna A = nome master,
colunas B em diante = TODOS os códigos internos que apontam para esse
produto (resolve códigos duplicados do sistema legado).

Usado em duas frentes: troca do nome nos PDF/Excel de embarque e o
somatório por produto no simulador. O mapa é cacheado em memória e
recarregado no upload.
"""

import io
import openpyxl
from storage import master_existe, carregar_master_bytes

_cache = None   # {codigo_str: nome_master}
_ordem = None   # [nome_master, ...] na ordem da planilha


def _norm(v):
    """Normaliza um valor de célula de código para string (163.0 -> '163')."""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _carregar():
    mapa = {}
    ordem = []
    if not master_existe():
        return mapa, ordem
    wb = openpyxl.load_workbook(io.BytesIO(carregar_master_bytes()), data_only=True)
    ws = wb['Produtos MASTER'] if 'Produtos MASTER' in wb.sheetnames else wb[wb.sheetnames[0]]
    for r in range(2, ws.max_row + 1):
        nome = ws.cell(r, 1).value
        if not nome or not str(nome).strip():
            continue
        nome = str(nome).strip()
        ordem.append(nome)
        for c in range(2, ws.max_column + 1):
            s = _norm(ws.cell(r, c).value)
            if s and s.replace('.', '', 1).isdigit():
                mapa[s] = nome
    return mapa, ordem


def _garantir():
    global _cache, _ordem
    if _cache is None:
        _cache, _ordem = _carregar()


def get_mapa():
    _garantir()
    return _cache


def get_ordem():
    """Lista de nomes master na ordem da planilha."""
    _garantir()
    return _ordem


def recarregar():
    global _cache, _ordem
    _cache, _ordem = _carregar()
    return _cache


def nome_master(codigo, fallback=''):
    """Nome master de um código interno; fallback (ex.: nome do cliente) se não mapeado."""
    return get_mapa().get(_norm(codigo), fallback)


# Sufixos de embalagem que definem a variação — removê-los revela a "raiz"
# (o produto físico, somado para planejamento de produção).
_SUFIXOS = ['GRANEL', 'PORCIONADO', 'PORCIONADA', 'DO SEU JEITO 400G',
            'DO SEU JEITO', 'PACT. 400G', '400G', '500G']


def raiz(nome):
    """Raiz (produto base) de um nome master, tirando o sufixo de embalagem.
    Ex.: 'BACON PORCIONADO' -> 'BACON'."""
    import re
    n = str(nome or '').strip()
    for s in sorted(_SUFIXOS, key=len, reverse=True):
        r = re.sub(r'\s+' + re.escape(s) + r'\s*$', '', n)
        if r != n:
            return r.strip()
    return n
