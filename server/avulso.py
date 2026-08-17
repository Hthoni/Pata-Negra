"""
Cliente Avulso — pedidos pontuais sem parser/perfil dedicado.

Em vez de um Perfil por cliente (produtos + filiais cadastrados), usa uma
única planilha de FATURAMENTO com os dados fixos por CNPJ (razão social,
condição de pagamento, vendedor, endereço de entrega). O operador digita o
CNPJ no popup, o sistema completa esses dados automaticamente, e monta o
pedido linha a linha escolhendo o produto da tabela MASTER (não existe
perfil próprio pra esse "cliente", então não há preço de referência nem
alerta de divergência de preço pra esses pedidos).

Planilha de faturamento (aba única, uma linha por CNPJ):
  Col A: CNPJ              Col D: Vendedor
  Col B: Razão Social       Col E: Endereço de Entrega
  Col C: Condição de Pagamento
Cabeçalho na linha 1; dados a partir da linha 2.
"""
import io
import re
import openpyxl


def _normaliza_cnpj(cnpj):
    return re.sub(r'\D', '', str(cnpj or ''))


def ler_faturamento_avulso(xlsx_bytes):
    """Lê a planilha de faturamento avulso e devolve
    {cnpj_normalizado: {cnpj, razaoSocial, condPgto, vendedor, endereco}}."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        cnpj_raw = str(row[0]).strip()
        cnpj_norm = _normaliza_cnpj(cnpj_raw)
        if not cnpj_norm:
            continue
        out[cnpj_norm] = {
            'cnpj': cnpj_raw,
            'razaoSocial': str(row[1] or '').strip(),
            'condPgto': str(row[2] or '').strip(),
            'vendedor': str(row[3] or '').strip(),
            'endereco': str(row[4] or '').strip(),
        }
    return out


def calc_item_avulso(qtde, unidade, peso_unit_kg, preco_unit):
    """kg e valor de um item avulso, a partir da unidade escolhida pelo
    operador (kg -> qtde já é kg; pct/cx -> qtde x peso da unidade em kg).
    Preço sempre na MESMA unidade escolhida (preço por kg se unidade=kg,
    preço por caixa se unidade=cx etc) — valor = qtde x preço, sempre."""
    qtde = float(qtde or 0)
    preco_unit = float(preco_unit or 0)
    if unidade == 'kg':
        kg = qtde
    else:
        kg = qtde * float(peso_unit_kg or 0)
    valor = round(qtde * preco_unit, 2)
    return round(kg, 3), valor
