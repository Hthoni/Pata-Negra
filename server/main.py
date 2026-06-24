"""
Pata Negra — Sistema de Processamento de Pedidos
# rebuild trigger summer
Servidor Flask: recebe pedidos de clientes em PDF, faz o parsing,
e gera Excel (upload interno) + PDF (expedição) padronizados.

Arquitetura modular:
  storage.py    -> persistência de perfis no Google Cloud Storage
  perfil.py     -> leitura do Perfil Excel + matching de produtos
  excel_gen.py  -> geração do Excel de upload
  pdf_gen.py    -> geração do PDF de expedição
  parsers/      -> um parser por cliente (isolados entre si)
"""
from flask import Flask, request, jsonify, send_file
import openpyxl
import io
import base64
import datetime
import traceback
from flask_cors import CORS

from storage import perfil_existe, salvar_perfil, carregar_perfil_bytes, perfil_filename, salvar_romaneio, listar_romaneios, deletar_romaneio
from perfil import ler_perfil, ler_filiais, buscar_filial, ler_operadores
from excel_gen import gerar_excel
from pdf_gen import gerar_pdf

import re as _re
def _normaliza_cnpj(cnpj):
    """Remove formatação do CNPJ, deixando só dígitos. Definida aqui
    localmente pra não depender da versão do perfil.py no servidor."""
    return _re.sub(r'\D', '', str(cnpj or ''))
import importlib
import pkgutil
import parsers

app = Flask(__name__)
CORS(app)

# Descoberta automática de parsers: qualquer módulo em parsers/ que
# exponha uma função parse() é registrado automaticamente.
# O nome de exibição pode ser definido como __cliente_nome__ no módulo;
# caso contrário usa o nome do módulo com capitalização automática.
CLIENTES = {}
for _info in pkgutil.iter_modules(parsers.__path__):
    if _info.name.startswith('_'):
        continue
    try:
        _mod = importlib.import_module(f'parsers.{_info.name}')
        if not hasattr(_mod, 'parse'):
            continue
        _nome = getattr(_mod, '__cliente_nome__', _info.name.replace('_', ' ').title())
        CLIENTES[_info.name] = {'nome': _nome, 'parse': _mod.parse}
    except Exception as _e:
        print(f'[WARN] parser {_info.name} não carregado: {_e}')

# Registro de clientes sem PDF (pedido lançado manualmente no popup, ex:
# pedidos por telefone/WhatsApp). Mesma chave usada em /perfil/<cliente> e
# nos endpoints /operadores, /filiais, /produtos e /processar-manual abaixo.
# multiFilial=True: cliente tem várias lojas (tabela M:N:O:P:Q do Perfil),
# o popup pede pra escolher a filial pelo nome.
# multiFilial=False: cliente tem CNPJ/endereço único — vem direto do
# cabeçalho do Perfil (linhas 3-6), sem seleção nenhuma no popup.
CLIENTES_MANUAIS = {
    'guanabara_lojas': {'nome': 'Guanabara Lojas', 'multiFilial': True},
    'mundial_lojas': {'nome': 'Mundial Lojas', 'multiFilial': True},
    'guanabara_central': {'nome': 'Guanabara Central', 'multiFilial': False},
    'mundial_central': {'nome': 'Mundial Central', 'multiFilial': False},
}


def _gerar_arquivos_por_empresa(dados, filiais, logo_bytes=None):
    """Detecta split por empresa (produtos de Indústria e Distribuidora no
    mesmo pedido) e gera os pares Excel+PDF correspondentes. Compartilhado
    entre /processar (pedidos parseados de PDF) e /processar-manual
    (pedidos lançados na tela, sem PDF) — a lógica de split não depende de
    como os itens chegaram, só do campo 'empresa' de cada item."""
    empresas_nos_itens = set(
        it.get('empresa') or dados.get('empresa', 2)
        for f in filiais for it in f['itens']
    )
    empresas_nos_itens.discard(None)
    if not empresas_nos_itens:
        empresas_nos_itens = {dados.get('empresa', 2)}

    arquivos = []
    eb_simples = pb_simples = None  # guarda o único par gerado quando não há split, p/ reaproveitar
    for emp_split in sorted(empresas_nos_itens):
        override = emp_split if len(empresas_nos_itens) > 1 else None
        eb = gerar_excel(dados, empresa_override=override)
        pb = gerar_pdf(dados, empresa_override=override, logo_bytes=logo_bytes)
        if len(empresas_nos_itens) == 1:
            eb_simples, pb_simples = eb, pb
        label = ('Indústria' if emp_split == 1 else 'Distribuidora') if len(empresas_nos_itens) > 1 else ''
        arquivos.append({
            'empresa': emp_split,
            'label': label,
            'excel': base64.b64encode(eb).decode(),
            'pdf': base64.b64encode(pb).decode(),
        })
    return arquivos, eb_simples, pb_simples, len(empresas_nos_itens) > 1


@app.route('/romaneios')
def get_romaneios():
    """Lista todos os romaneios pendentes para o mapa."""
    try:
        return jsonify(listar_romaneios())
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/romaneio/<romaneio_id>', methods=['DELETE'])
def delete_romaneio(romaneio_id):
    """Marca pedido como entregue deletando o pin do mapa."""
    try:
        ok = deletar_romaneio(romaneio_id)
        if ok:
            return jsonify({'ok': True})
        return jsonify({'erro': 'Romaneio não encontrado'}), 404
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/health')
def health():
    # DIAGNÓSTICO TEMPORÁRIO (18/06): captura erro por cliente em vez de
    # deixar uma exceção silenciosa esconder o motivo de um perfil não
    # aparecer. Reverter pra versão simples depois de identificar a causa.
    perfis = {}
    erros = {}
    for c in {**CLIENTES, **CLIENTES_MANUAIS}:
        try:
            if perfil_existe(c):
                perfis[c] = perfil_filename(c)
        except Exception as e:
            erros[c] = f'{type(e).__name__}: {e}'
    # Lista de clientes disponíveis (PDF + manuais) para o frontend
    clientes_info = {}
    for cid, cdata in CLIENTES.items():
        clientes_info[cid] = {
            'nome': cdata['nome'],
            'tipo': 'proprio',
            'manual': False,
        }
    for cid, cdata in CLIENTES_MANUAIS.items():
        clientes_info[cid] = {
            'nome': cdata['nome'],
            'tipo': 'manual',
            'manual': True,
            'multiFilial': cdata.get('multiFilial', True),
        }

    resp = {'status': 'ok', 'perfis': perfis, 'clientes': clientes_info}
    if erros:
        resp['erros'] = erros
    return jsonify(resp)


@app.route('/perfil/<cliente>', methods=['POST'])
def upload_perfil(cliente):
    """Salva ou atualiza o perfil de um cliente no servidor."""
    if cliente not in CLIENTES and cliente not in CLIENTES_MANUAIS:
        return jsonify({'erro': f'Cliente inválido: {cliente}'}), 400
    f = request.files.get('perfil')
    if not f:
        return jsonify({'erro': 'Envie o arquivo perfil'}), 400
    filename = f.filename or ''
    salvar_perfil(cliente, f.read(), filename)
    return jsonify({'ok': True, 'cliente': cliente, 'filename': filename, 'mensagem': 'Perfil salvo com sucesso'})


@app.route('/logo/<cliente>')
def logo(cliente):
    """Retorna a logo extraída do perfil Excel do cliente."""
    if not perfil_existe(cliente):
        return jsonify({'erro': 'Perfil não encontrado'}), 404
    try:
        perfil_bytes = carregar_perfil_bytes(cliente)
        wb = openpyxl.load_workbook(io.BytesIO(perfil_bytes))
        ws = wb[wb.sheetnames[0]]
        if not ws._images:
            return jsonify({'erro': 'Sem imagem no perfil'}), 404
        img = ws._images[0]
        img.ref.seek(0)
        data = img.ref.read()
        return send_file(io.BytesIO(data), mimetype='image/png')
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/processar', methods=['POST'])
def processar():
    try:
        perfil_file = request.files.get('perfil')
        pedido_file = request.files.get('pedido')
        cliente = request.form.get('cliente', 'dom_atacarejo')

        if not pedido_file:
            return jsonify({'erro': 'Envie o pedido'}), 400

        if cliente not in CLIENTES:
            return jsonify({'erro': f'Cliente {cliente} não implementado'}), 400

        # Perfil: usa o enviado agora (e salva) ou o salvo no servidor
        if perfil_file:
            perfil_bytes = perfil_file.read()
            salvar_perfil(cliente, perfil_bytes, perfil_file.filename)
        elif perfil_existe(cliente):
            perfil_bytes = carregar_perfil_bytes(cliente)
        else:
            return jsonify({'erro': f'Nenhum perfil disponível para {cliente}. Faça upload do perfil primeiro.'}), 400

        meta, produtos = ler_perfil(perfil_bytes)
        filiais_map = ler_filiais(perfil_bytes)
        pdf_bytes = pedido_file.read()

        parse_fn = CLIENTES[cliente]['parse']
        filiais = parse_fn(pdf_bytes, produtos)
        if cliente == 'assai' and filiais:
            meta['empresa'] = filiais[0]['empresa']

        if not filiais:
            return jsonify({'erro': 'Nenhuma filial encontrada no pedido'}), 400

        # Enriquecer cada filial com nome oficial e número, buscando pelo CNPJ
        # (regra única para todos os clientes: CNPJ é o dado mais confiável)
        for fd in filiais:
            nome_oficial, num_filial = buscar_filial(fd.get('cnpj', ''), filiais_map)
            if nome_oficial:
                fd['filial'] = nome_oficial
            if num_filial is not None:
                fd['numFilial'] = num_filial

        dados = {**meta, 'filiais': filiais, 'clienteNome': CLIENTES[cliente]['nome']}

        # Extrair logo do perfil para o PDF
        logo_bytes = None
        try:
            wb_logo = openpyxl.load_workbook(io.BytesIO(perfil_bytes))
            ws_logo = wb_logo[wb_logo.sheetnames[0]]
            if ws_logo._images:
                ws_logo._images[0].ref.seek(0)
                logo_bytes = ws_logo._images[0].ref.read()
        except Exception:
            pass

        arquivos, eb_simples, pb_simples, split = _gerar_arquivos_por_empresa(dados, filiais, logo_bytes=logo_bytes)

        # Salvar pin de mapa para cada filial com coordenadas
        from datetime import datetime
        for fd in filiais:
            lat = fd.get('lat')
            lng = fd.get('lng')
            if lat is None or lng is None:
                continue
            its = fd.get('itens', [])
            if not its:
                continue
            ts = datetime.utcnow().strftime('%Y%m%d_%H%M%S')
            filial_slug = re.sub(r'[^a-z0-9]', '_', fd['filial'].lower())
            rid = f"{cliente}_{filial_slug}_{ts}"
            salvar_romaneio(rid, {
                'id': rid,
                'cliente': cliente,
                'clienteNome': CLIENTES[cliente]['nome'],
                'filial': fd['filial'],
                'cnpj': fd.get('cnpj', ''),
                'lat': lat,
                'lng': lng,
                'dataPedido': fd.get('dataPedido', fd.get('dataEmissao', '')),
                'dataGeracao': datetime.utcnow().isoformat(),
                'kgPlanejados': round(sum(float(i.get('kgPlanejados', 0)) for i in its), 1),
                'pedidoNum': fd.get('pedidoNum', ''),
            })

        todos_itens = [i for f in filiais for i in f['itens']]
        return jsonify({
            'ok': True,
            'split': split,
            'filiais': len(filiais),
            'itens': len(todos_itens),
            'totalKg': round(sum(i['kgPlanejados'] for i in todos_itens), 1),
            'totalValor': round(sum(i['valorPedido'] for i in todos_itens), 2),
            'resumo': [{'filial': f['filial'], 'pedidoNum': f.get('pedidoNum', ''),
                        'itens': len(f['itens']),
                        'kg': round(sum(i['kgPlanejados'] for i in f['itens']), 1),
                        'valor': round(sum(i['valorPedido'] for i in f['itens']), 2)}
                       for f in filiais],
            'arquivos': arquivos,
            # compatibilidade retroativa (caso simples) — reaproveita o que já foi gerado acima
            'excel': base64.b64encode(eb_simples).decode() if eb_simples is not None else '',
            'pdf': base64.b64encode(pb_simples).decode() if pb_simples is not None else '',
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'erro': str(e)}), 500


@app.route('/operadores/<cliente>')
def operadores_cliente(cliente):
    """Lista os operadores cadastrados no Perfil de um cliente do fluxo
    manual, pra alimentar o dropdown de quem está lançando o pedido."""
    if cliente not in CLIENTES_MANUAIS:
        return jsonify({'erro': f'Cliente {cliente} não usa fluxo manual'}), 400
    if not perfil_existe(cliente):
        return jsonify({'erro': 'Perfil não encontrado'}), 404
    perfil_bytes = carregar_perfil_bytes(cliente)
    return jsonify({'operadores': ler_operadores(perfil_bytes)})


@app.route('/filiais/<cliente>')
def filiais_cliente(cliente):
    """Lista as filiais disponíveis pro popup de pedido manual.
    Clientes multiFilial=True: lê a tabela M:N:O:P:Q do Perfil (várias lojas).
    Clientes multiFilial=False (ex: Guanabara Central): não existe tabela —
    devolve uma única filial sintetizada a partir do cabeçalho do Perfil
    (CNPJ/Filial/Endereço únicos), pra manter o mesmo formato de resposta
    e o frontend não precisar de um caminho de código totalmente separado."""
    if cliente not in CLIENTES_MANUAIS:
        return jsonify({'erro': f'Cliente {cliente} não usa fluxo manual'}), 400
    if not perfil_existe(cliente):
        return jsonify({'erro': 'Perfil não encontrado'}), 404
    perfil_bytes = carregar_perfil_bytes(cliente)
    if CLIENTES_MANUAIS[cliente].get('multiFilial', True):
        filiais_map = ler_filiais(perfil_bytes)
        lista = [{'cnpj': cnpj, 'nome': info['nome'], 'numero': info['numero'],
                  'endereco': info['endereco'], 'cidade': info['cidade']}
                 for cnpj, info in filiais_map.items()]
    else:
        meta, _ = ler_perfil(perfil_bytes)
        lista = [{
            'cnpj': _normaliza_cnpj(meta.get('cnpjPerfil', '')),
            'nome': meta.get('filialPerfil') or meta.get('clienteNomePerfil') or CLIENTES_MANUAIS[cliente]['nome'],
            'numero': None,
            'endereco': meta.get('enderecoPerfil', ''),
            'cidade': '',
        }]
    lista.sort(key=lambda f: f['nome'])
    return jsonify({'filiais': lista, 'multiFilial': CLIENTES_MANUAIS[cliente].get('multiFilial', True)})


@app.route('/produtos/<cliente>')
def produtos_cliente(cliente):
    """Lista os produtos cadastrados no Perfil de um cliente do fluxo
    manual, pra pré-popular as linhas do popup (nome, formato, embalagem
    e a unidade em que a quantidade deve ser digitada — cx ou kg)."""
    if cliente not in CLIENTES_MANUAIS:
        return jsonify({'erro': f'Cliente {cliente} não usa fluxo manual'}), 400
    if not perfil_existe(cliente):
        return jsonify({'erro': 'Perfil não encontrado'}), 404
    perfil_bytes = carregar_perfil_bytes(cliente)
    _, produtos = ler_perfil(perfil_bytes)
    lista = [{
        'index': i,
        'codInterno': p['codInterno'],
        'nomePerfil': p['nomePerfil'],
        'formato': p['formato'],
        'embalagem': p['embalagem'],
        'unidFat': p['unidFat'],
        'kgCx': p['kgCx'],
        'precoUnit': p['precoUnit'],
    } for i, p in enumerate(produtos)]
    return jsonify({'produtos': lista})


@app.route('/processar-manual', methods=['POST'])
def processar_manual():
    """Gera Excel+PDF a partir de um pedido lançado manualmente no popup
    (clientes sem PDF de pedido — pedidos por telefone/WhatsApp em formato
    livre). O operador escolhe a filial pelo nome e digita a quantidade de
    cada produto já pré-carregado do Perfil; nenhum parsing de texto livre
    é necessário, já que o produto e a filial vêm de seleção direta."""
    try:
        body = request.get_json(force=True) or {}
        cliente = body.get('cliente', '')
        operador = body.get('operador', '')
        filial_nome = body.get('filialNome', '')
        itens_form = body.get('itens', [])

        if cliente not in CLIENTES_MANUAIS:
            return jsonify({'erro': f'Cliente {cliente} não usa fluxo manual'}), 400
        if not perfil_existe(cliente):
            return jsonify({'erro': f'Nenhum perfil disponível para {cliente}'}), 400
        if not operador:
            return jsonify({'erro': 'Selecione o operador'}), 400

        multi_filial = CLIENTES_MANUAIS[cliente].get('multiFilial', True)
        if multi_filial and not filial_nome:
            return jsonify({'erro': 'Selecione a filial'}), 400

        perfil_bytes = carregar_perfil_bytes(cliente)
        meta, produtos = ler_perfil(perfil_bytes)
        operadores_validos = ler_operadores(perfil_bytes)

        if operador not in operadores_validos:
            return jsonify({'erro': f'Operador "{operador}" não cadastrado no perfil'}), 400

        if multi_filial:
            # encontra a filial selecionada pelo nome (a seleção é direta, não por CNPJ extraído de PDF)
            filiais_map = ler_filiais(perfil_bytes)
            cnpj_sel, filial_info = None, None
            for cnpj, info in filiais_map.items():
                if info['nome'] == filial_nome:
                    cnpj_sel, filial_info = cnpj, info
                    break
            if not filial_info:
                return jsonify({'erro': f'Filial "{filial_nome}" não encontrada no perfil'}), 400
        else:
            # cliente 'Central': CNPJ/Filial/Endereço únicos, vêm do cabeçalho do Perfil
            cnpj_sel = _normaliza_cnpj(meta.get('cnpjPerfil', ''))
            filial_info = {
                'nome': meta.get('filialPerfil') or meta.get('clienteNomePerfil') or CLIENTES_MANUAIS[cliente]['nome'],
                'numero': None,
                'endereco': meta.get('enderecoPerfil', ''),
            }

        itens = []
        for it in itens_form:
            qtde = it.get('quantidade')
            if not qtde:
                continue  # linha em branco, produto não pedido nessa ligação
            qtde = float(qtde)
            if qtde <= 0:
                continue
            idx = it.get('index')
            produto = produtos[idx] if isinstance(idx, int) and 0 <= idx < len(produtos) else None
            if not produto:
                continue
            kgCx = produto['kgCx']
            unidFat = produto['unidFat']
            kgPlan = qtde * kgCx if unidFat == 'cx' else qtde
            preco = produto['precoUnit']
            itens.append({
                'empresa': produto.get('empresa'),
                'codInterno': produto['codInterno'],
                'nomeProduto': produto['nomePerfil'],
                'formato': produto.get('formato', ''),
                'embalagem': produto.get('embalagem', ''),
                'kgCx': kgCx,
                'kgPlanejados': kgPlan,
                'nrCaixas': round(kgPlan / kgCx, 1) if kgCx else 0,
                'obs': produto.get('obs', ''),
                'qtdeMultipl': qtde if unidFat == 'cx' else kgPlan,
                'precoUnit': preco,
                'valorPedido': round(kgPlan * preco, 2),
                'precoSistema': preco,
                'unidFat': unidFat,
            })

        if not itens:
            return jsonify({'erro': 'Nenhum item com quantidade preenchida'}), 400

        agora = datetime.datetime.now()
        pedido_num = f"MANUAL-{agora.strftime('%Y%m%d-%H%M%S')}"

        lat = filial_info.get('lat')
        lng = filial_info.get('lng')

        filiais = [{
            'filial': filial_info['nome'],
            'numFilial': filial_info['numero'],
            'pedidoNum': pedido_num,
            'cnpj': cnpj_sel,
            'endereco': filial_info['endereco'],
            'dataPedido': agora.strftime('%d/%m/%Y'),
            'dataEntrega': '',
            'condPgto': '',
            'solicitante': operador,
            'empresa': meta.get('empresa', 2),
            'itens': itens,
            'lat': lat,
            'lng': lng,
        }]

        dados = {**meta, 'filiais': filiais, 'clienteNome': CLIENTES_MANUAIS[cliente]['nome']}

        # Salvar pin de mapa se tiver coordenadas
        if lat is not None and lng is not None:
            try:
                ts = datetime.datetime.utcnow().strftime('%Y%m%d_%H%M%S')
                filial_slug = re.sub(r'[^a-z0-9]', '_', filial_info['nome'].lower())
                rid = f"{cliente}_{filial_slug}_{ts}"
                salvar_romaneio(rid, {
                    'id': rid,
                    'cliente': cliente,
                    'clienteNome': CLIENTES_MANUAIS[cliente]['nome'],
                    'filial': filial_info['nome'],
                    'cnpj': cnpj_sel,
                    'lat': lat,
                    'lng': lng,
                    'dataPedido': agora.strftime('%d/%m/%Y'),
                    'dataGeracao': datetime.datetime.utcnow().isoformat(),
                    'kgPlanejados': round(sum(float(i.get('kgPlanejados', 0)) for i in itens), 1),
                    'pedidoNum': pedido_num,
                })
            except Exception as e:
                print(f'[WARN] romaneio manual não salvo: {e}')

        arquivos, eb_simples, pb_simples, split = _gerar_arquivos_por_empresa(dados, filiais, logo_bytes=logo_bytes)

        return jsonify({
            'ok': True,
            'split': split,
            'filiais': 1,
            'itens': len(itens),
            'totalKg': round(sum(i['kgPlanejados'] for i in itens), 1),
            'totalValor': round(sum(i['valorPedido'] for i in itens), 2),
            'pedidoNum': pedido_num,
            'resumo': [{'filial': filial_info['nome'], 'pedidoNum': pedido_num,
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


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
