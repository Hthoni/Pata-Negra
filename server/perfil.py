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
        # Campos de cabeçalho (linhas 3-6, coluna C) — usados pelos clientes
        # 'Central' (sem tabela de filiais M:N:O): CNPJ/Filial/Endereço únicos,
        # preenchidos direto aqui em vez de vir de uma seleção de loja.
        'clienteNomePerfil': str(pdata[2][2]).strip() if pdata[2][2] else '',
        'cnpjPerfil': str(pdata[3][2]).strip() if pdata[3][2] else '',
        'filialPerfil': str(pdata[4][2]).strip() if pdata[4][2] else '',
        'enderecoPerfil': str(pdata[5][2]).strip() if pdata[5][2] else '',
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


def _normaliza_cnpj(cnpj):
    """Remove pontuação do CNPJ para comparação confiável. Aceita tanto
    texto ('31.698.759/0001-13') quanto número puro (a célula do Excel pode
    vir como int quando não há formatação de texto aplicada)."""
    return ''.join(c for c in str(cnpj or '') if c.isdigit())


def ler_filiais(perfil_bytes):
    """Lê a tabela de filiais do Perfil Excel (colunas M:N:O, a partir da
    linha 9): CNPJ | Nome Filial | Número Filial. Opcionalmente também lê
    Endereço (col P), Cidade (col Q) e Região (col R), usadas pelo fluxo de
    pedido manual e pelo card de Região no PDF de expedição — para clientes
    que não têm essas colunas, ficam como string vazia.
    Retorna dict {cnpj_normalizado: {'nome', 'numero', 'endereco', 'cidade', 'regiao', 'lat', 'lng'}}.
    Usado para enriquecer pedidos que só trazem CNPJ (Atacadão) ou
    CNPJ+nome (DOM) com o número de filial cadastrado uma única vez no perfil,
    e para alimentar o dropdown de filiais no fluxo manual."""
    wb_p = openpyxl.load_workbook(io.BytesIO(perfil_bytes), data_only=True)
    pws = wb_p[wb_p.sheetnames[0]]
    pdata = list(pws.iter_rows(values_only=True))
    filiais = {}
    for r in pdata[8:]:  # a partir da linha 9 (índice 8)
        if not r or len(r) < 15:
            continue
        cnpj_raw, nome, numero = r[12], r[13], r[14]  # colunas M, N, O
        endereco = r[15] if len(r) > 15 else None  # coluna P (opcional)
        cidade = r[16] if len(r) > 16 else None  # coluna Q (opcional)
        regiao = r[17] if len(r) > 17 else None  # coluna R = região (opcional)
        lat = r[18] if len(r) > 18 else None  # coluna S
        lng = r[19] if len(r) > 19 else None  # coluna T
        if not cnpj_raw:
            continue
        cnpj_norm = _normaliza_cnpj(cnpj_raw)
        if cnpj_norm:
            filiais[cnpj_norm] = {
                'nome': str(nome or '').strip(),
                'numero': numero,
                'endereco': str(endereco or '').strip(),
                'cidade': str(cidade or '').strip(),
                'regiao': str(regiao or '').strip(),
                'lat': float(lat) if lat is not None else None,
                'lng': float(lng) if lng is not None else None,
            }
    return filiais


def buscar_filial(cnpj, filiais_map):
    """Busca nome e número de filial pelo CNPJ extraído do pedido.
    Retorna (nome, numero) ou (None, None) se não encontrado."""
    cnpj_norm = _normaliza_cnpj(cnpj)
    info = filiais_map.get(cnpj_norm)
    if info:
        return info['nome'], info['numero']
    return None, None


def ler_operadores(perfil_bytes):
    """Lê a lista de operadores (coluna L, a partir da linha 9) do Perfil.
    Usado nos clientes de pedido manual (sem PDF, ex: Guanabara Lojas), para
    alimentar o dropdown de quem está lançando o pedido no popup. Cresce
    livremente, sem depender do tamanho da tabela de produtos ou de filiais
    — cada lista (produtos, operadores, filiais) avança na sua própria
    coluna, independente das outras."""
    wb_p = openpyxl.load_workbook(io.BytesIO(perfil_bytes), data_only=True)
    pws = wb_p[wb_p.sheetnames[0]]
    pdata = list(pws.iter_rows(values_only=True))
    operadores = []
    for r in pdata[8:]:  # a partir da linha 9 (mesmo início da tabela de filiais)
        if not r or len(r) < 12:
            continue
        nome = r[11]  # coluna L
        if nome:
            operadores.append(str(nome).strip())
    return operadores


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
