"""
Leitura do Perfil Excel de cada cliente e algoritmo de matching
de produtos do pedido (PDF) com produtos cadastrados no perfil.
"""
import io
import openpyxl

CNPJ_DISTRIBUIDORA = '56.423.719'
CNPJ_INDUSTRIA = '10.171.633'


def match_perfil(nome, produtos):
    """Encontra o produto do perfil que melhor corresponde ao nome do PDF.
    NUNCA usar o código do cliente no PDF — sempre casar por nome."""
    n = (nome or '').lower().strip()
    best, score = None, 0
    for p in produtos:
        a = (p.get('nomePerfil', '') or '').lower().strip()
        if not a:
            continue
        if a == n:
            s = 100
        elif n in a:
            s = 50 + len(a)
        elif a in n:
            s = 40 + len(n)
        else:
            common = len(set(a.split()) & set(n.split()))
            s = common * 10 if common >= 2 else 0
        if s > score:
            score, best = s, p
    return best


def ler_perfil(perfil_bytes):
    """Lê o Perfil Excel e retorna (meta, produtos).
    meta: dados do cabeçalho (vendedor, telefone, códigos).
    produtos: lista de dicts com os produtos cadastrados, incluindo
    a coluna A (Fat.) que define a empresa de faturamento (1=Indústria, 2=Distribuidora)."""
    wb_p = openpyxl.load_workbook(io.BytesIO(perfil_bytes), data_only=True)
    pws = wb_p[wb_p.sheetnames[0]]
    pdata = list(pws.iter_rows(values_only=True))
    meta = {
        'empresa': pdata[0][9] if pdata[0][9] else 2,
        'codVend': str(pdata[2][6]) if pdata[2][6] else '',
        'codCond': str(pdata[3][6]) if pdata[3][6] else '',
        'vendedor': pdata[2][8] or '',
        'telefone': pdata[2][9] or '',
    }
    produtos = []
    for r in pdata[7:]:
        if not r or not r[2]:
            break
        produtos.append({
            'empresa': int(r[0]) if r[0] in (1, 2) else None,  # coluna A: Fat.
            'codInterno': r[1],
            'nomePerfil': str(r[2]).strip(),
            'formato': str(r[3] or '').strip(),
            'embalagem': str(r[4]).strip(),
            'kgCx': float(r[6] or 20),
            'unidFat': str(r[7] or 'kg').strip(),
            'precoUnit': float(r[8] or 0),
            'obs': str(r[9] or '').strip(),
        })
    return meta, produtos


def processar_item(cod_cli, nome_raw, emb_tipo, qtde_emb, qtde_ped, preco, total, produtos):
    """Normaliza um item do pedido casando com o perfil e calculando
    kg planejados, número de caixas e demais campos derivados."""
    import re
    nome_raw = re.sub(r'\s+', ' ', nome_raw).strip()
    pf = match_perfil(nome_raw, produtos)
    kgCx = pf['kgCx'] if pf else 20
    embalagem = pf['embalagem'] if pf else ('CX-' + str(qtde_emb) if emb_tipo in ['CX', 'CXA'] else 'CX-20')
    if emb_tipo in ['CX', 'CXA']:
        kgPlan = qtde_ped * kgCx
        nrCx = qtde_ped
        qtdeMult = qtde_ped
        unidFat = 'cx'
    else:
        kgPlan = qtde_ped
        nrCx = round(kgPlan / kgCx, 1) if kgCx else 0
        qtdeMult = kgPlan
        unidFat = 'kg'
    return {
        'empresa': pf.get('empresa') if pf else None,  # herda do perfil (coluna A)
        'codInterno': pf['codInterno'] if pf else cod_cli,
        'nomeProduto': pf['nomePerfil'] if pf else nome_raw,
        'formato': pf.get('formato', '') if pf else '',
        'embalagem': embalagem,
        'kgCx': kgCx,
        'kgPlanejados': kgPlan,
        'nrCaixas': nrCx,
        'obs': pf.get('obs', '') if pf else '',
        'qtdeMultipl': qtdeMult,
        'precoUnit': preco,
        'valorPedido': total,
        'precoSistema': pf['precoUnit'] if pf else 0,
        'unidFat': unidFat,
    }
