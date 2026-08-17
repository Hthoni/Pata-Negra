"""
NOVO (17/08/2026): Cliente Avulso.

Cole este bloco no main.py:
1. O import abaixo, junto dos outros imports do topo do arquivo.
2. As 3 rotas novas, em qualquer lugar depois de _persistir_arquivos_romaneio
   (elas usam _gerar_arquivos_por_empresa, _persistir_arquivos_romaneio e
   _empresas_da_filial, que já existem no arquivo).

Reaproveita a infraestrutura de perfil já existente (salvar_perfil /
carregar_perfil_bytes / perfil_existe) pra guardar a planilha de
faturamento avulso, usando uma chave de cliente reservada
('_avulso_faturamento') — não precisa mexer no storage.py.
"""

# --- import (colar junto aos outros imports do topo) ---
from avulso import ler_faturamento_avulso, calc_item_avulso

_CLIENTE_AVULSO_KEY = '_avulso_faturamento'


# --- rotas novas ---

@app.route('/faturamento-avulso', methods=['POST'])
def upload_faturamento_avulso():
    """Sobe/atualiza a planilha de faturamento de clientes avulsos
    (CNPJ | Razão Social | Condição Pagto | Vendedor | Endereço)."""
    f = request.files.get('faturamento') or request.files.get('file')
    if not f:
        return jsonify({'erro': 'Envie o arquivo de faturamento'}), 400
    try:
        file_bytes = f.read()
        dados = ler_faturamento_avulso(file_bytes)  # valida o formato antes de salvar
        salvar_perfil(_CLIENTE_AVULSO_KEY, file_bytes, f.filename or 'faturamento_avulso.xlsx')
        return jsonify({'ok': True, 'cnpjs': len(dados), 'mensagem': f'{len(dados)} CNPJs carregados'})
    except Exception as e:
        return jsonify({'erro': f'Não consegui ler a planilha: {e}'}), 400


@app.route('/faturamento-avulso/<cnpj>')
def buscar_faturamento_avulso(cnpj):
    """Busca os dados de faturamento de um CNPJ na planilha avulsa."""
    if not perfil_existe(_CLIENTE_AVULSO_KEY):
        return jsonify({'erro': 'Planilha de faturamento avulso ainda não foi cadastrada'}), 404
    try:
        dados = ler_faturamento_avulso(carregar_perfil_bytes(_CLIENTE_AVULSO_KEY))
        cnpj_norm = _normaliza_cnpj(cnpj)
        info = dados.get(cnpj_norm)
        if not info:
            return jsonify({'encontrado': False}), 404
        return jsonify({'encontrado': True, **info})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/master/produtos')
def master_produtos_lista():
    """Lista {codigo, nome} da tabela MASTER, pro dropdown de produto do
    popup de Cliente Avulso (não existe perfil próprio nesse fluxo)."""
    try:
        mapa = master.get_mapa()  # {codigo_normalizado: nome}
        lista = [{'codigo': c, 'nome': n} for c, n in mapa.items()]
        lista.sort(key=lambda p: p['nome'])
        return jsonify({'produtos': lista})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/processar-avulso', methods=['POST'])
def processar_avulso():
    """Gera Excel+PDF de um pedido de Cliente Avulso. Body:
    {cnpj, pedidoNum, operador, razaoSocial, condPgto, vendedor, endereco,
     itens: [{codigoMaster, nomeProduto, quantidade, unidade (kg|pct|cx),
               pesoUnitKg, embalagem, precoUnit, empresa (1|2)}]}
    razaoSocial/condPgto/vendedor/endereco vêm da planilha (auto-preenchido
    no popup) mas podem chegar sobrescritos se o CNPJ não foi encontrado e
    o operador preencheu manualmente — o backend não busca de novo, confia
    no que o front mandou."""
    try:
        body = request.get_json(force=True) or {}
        cnpj = (body.get('cnpj') or '').strip()
        pedido_num = (body.get('pedidoNum') or '').strip()
        operador = (body.get('operador') or '').strip()
        razao = (body.get('razaoSocial') or '').strip()
        cond = (body.get('condPgto') or '').strip()
        vendedor = (body.get('vendedor') or '').strip()
        endereco = (body.get('endereco') or '').strip()
        itens_form = body.get('itens', [])

        if not cnpj:
            return jsonify({'erro': 'CNPJ obrigatório'}), 400
        if not pedido_num:
            return jsonify({'erro': 'Nº do pedido obrigatório'}), 400
        if not razao:
            return jsonify({'erro': 'Razão social obrigatória (CNPJ não encontrado — preencha manualmente)'}), 400
        if not itens_form:
            return jsonify({'erro': 'Adicione pelo menos um item'}), 400

        itens = []
        for it in itens_form:
            qtde = it.get('quantidade')
            if not qtde:
                continue
            unidade = (it.get('unidade') or 'kg').strip().lower()
            peso_unit_kg = it.get('pesoUnitKg')
            preco = it.get('precoUnit')
            kg, valor = calc_item_avulso(qtde, unidade, peso_unit_kg, preco)
            if kg <= 0:
                continue
            empresa_item = int(it.get('empresa') or 2)
            itens.append({
                'empresa': empresa_item,
                'codInterno': it.get('codigoMaster') or '',
                'nomeProduto': it.get('nomeProduto') or '',
                'formato': '',
                'embalagem': it.get('embalagem') or '',
                'kgCx': float(peso_unit_kg or 0) if unidade in ('cx', 'pct') else 0,
                'kgPlanejados': kg,
                'nrCaixas': round(kg / peso_unit_kg, 1) if unidade in ('cx', 'pct') and peso_unit_kg else 0,
                'obs': '',
                'qtdeMultipl': float(qtde),
                'precoUnit': float(preco or 0),
                'valorPedido': valor,
                'precoSistema': float(preco or 0),  # sem perfil de referência -> nunca dispara alerta de preço
                'unidFat': 'cx' if unidade in ('cx', 'pct') else 'kg',
            })

        if not itens:
            return jsonify({'erro': 'Nenhum item com quantidade preenchida'}), 400

        agora = datetime.datetime.now()

        filial_dict = {
            'filial': razao,
            'numFilial': None,
            'pedidoNum': pedido_num,
            'cnpj': cnpj,
            'endereco': endereco,
            'dataPedido': agora.strftime('%d/%m/%Y'),
            'dataEntrega': '',
            'condPgto': cond,
            'solicitante': operador,
            'empresa': itens[0]['empresa'],
            'regiao': '',
            'itens': itens,
            'lat': None,
            'lng': None,
        }

        dados = {'filiais': [filial_dict], 'clienteNome': razao, 'vendedor': vendedor, 'telefone': ''}

        logo_bytes = None  # cliente avulso não tem logo própria no PDF

        ts = datetime.datetime.utcnow().strftime('%Y%m%d_%H%M%S_%f')
        rid = f"avulso_{ts}"
        emps = _empresas_da_filial(filial_dict, dados)

        salvar_romaneio(rid, {
            'id': rid,
            'cliente': 'avulso',
            'clienteNome': razao,
            'filial': razao,
            'numero': None,
            'regiao': '',
            'cnpj': cnpj,
            'lat': None,
            'lng': None,
            'dataPedido': agora.strftime('%d/%m/%Y'),
            'dataGeracao': datetime.datetime.utcnow().isoformat(),
            'kgPlanejados': round(sum(i['kgPlanejados'] for i in itens), 1),
            'itens': [{'cod': str(i.get('codInterno') or '').strip(), 'nome': str(i.get('nomeProduto') or ''), 'kg': round(_kg_pdf(i), 3)} for i in itens],
            'pedidoNum': pedido_num,
            'empresas': emps,
        })
        _persistir_arquivos_romaneio(rid, dados, filial_dict, emps, logo_bytes)

        arquivos, eb_simples, pb_simples, split = _gerar_arquivos_por_empresa(dados, [filial_dict], logo_bytes=logo_bytes)

        return jsonify({
            'ok': True,
            'split': split,
            'filiais': 1,
            'itens': len(itens),
            'totalKg': round(sum(i['kgPlanejados'] for i in itens), 1),
            'totalValor': round(sum(i['valorPedido'] for i in itens), 2),
            'pedidoNum': pedido_num,
            # FIX: 'resumo' é exigido pelo mostrarResultados() do front (já usado
            # por /processar e /processar-manual) — sem isso, data.resumo.map()
            # quebra na tela de resultado. Reaproveita o mesmo formato deles.
            'resumo': [{'filial': razao, 'pedidoNum': pedido_num,
                        'itens': len(itens),
                        'kg': round(sum(i['kgPlanejados'] for i in itens), 1),
                        'valor': round(sum(i['valorPedido'] for i in itens), 2)}],
            'arquivos': arquivos,
            'excel': base64.b64encode(eb_simples).decode() if eb_simples is not None else '',
            'pdf': base64.b64encode(pb_simples).decode() if pb_simples is not None else '',
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'erro': str(e)}), 500
