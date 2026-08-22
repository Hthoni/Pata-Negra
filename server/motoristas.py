"""
Motoristas — planilha simples (nome, telefone) usada em dois lugares:
  1. Dropdown de seleção de motorista ao criar uma entrega (Plano de
     Entrega), pra gravar motorista/telefoneMotorista na entrega.
  2. Canal WhatsApp (whatsapp.py) — busca reversa por telefone pra
     identificar se quem mandou mensagem é motorista.

Planilha (aba única, cabeçalho na linha 1, dados a partir da linha 2):
  Col A: Nome
  Col B: Telefone
"""
import io
import re
import openpyxl


def _normaliza_telefone(tel):
    return re.sub(r'\D', '', str(tel or ''))


def ler_motoristas(xlsx_bytes):
    """Devolve lista [{'nome':, 'telefone':}], na ordem da planilha."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        nome = str(row[0]).strip()
        telefone = str(row[1] or '').strip() if len(row) > 1 else ''
        if nome:
            out.append({'nome': nome, 'telefone': telefone})
    return out


def buscar_motorista_por_telefone(xlsx_bytes, telefone):
    """Confere se um telefone bate com algum motorista cadastrado."""
    alvo = _normaliza_telefone(telefone)
    for m in ler_motoristas(xlsx_bytes):
        if _normaliza_telefone(m['telefone']) == alvo:
            return m
    return None
