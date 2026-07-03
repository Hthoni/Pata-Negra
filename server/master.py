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

_cache = None  # {codigo_str: nome_master}


def _norm(v):
    """Normaliza um valor de célula de código para string (163.0 -> '163')."""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _carregar():
    mapa = {}
    if not master_existe():
        return mapa
    wb = openpyxl.load_workbook(io.BytesIO(carregar_master_bytes()), data_only=True)
    ws = wb['Produtos MASTER'] if 'Produtos MASTER' in wb.sheetnames else wb[wb.sheetnames[0]]
    for r in range(2, ws.max_row + 1):
        nome = ws.cell(r, 1).value
        if not nome or not str(nome).strip():
            continue
        nome = str(nome).strip()
        for c in range(2, ws.max_column + 1):
            s = _norm(ws.cell(r, c).value)
            if s and s.replace('.', '', 1).isdigit():
                mapa[s] = nome
    return mapa


def get_mapa():
    global _cache
    if _cache is None:
        _cache = _carregar()
    return _cache


def recarregar():
    global _cache
    _cache = _carregar()
    return _cache


def nome_master(codigo, fallback=''):
    """Nome master de um código interno; fallback (ex.: nome do cliente) se não mapeado."""
    return get_mapa().get(_norm(codigo), fallback)
