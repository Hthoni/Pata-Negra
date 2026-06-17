"""
Pata Negra — Sistema de Processamento de Pedidos
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
import traceback
from flask_cors import CORS

from storage import perfil_existe, salvar_perfil, carregar_perfil_bytes, perfil_filename
from perfil import ler_perfil, ler_filiais, buscar_filial
from excel_gen import gerar_excel
from pdf_gen import gerar_pdf
from parsers import (
    dom_atacarejo, atacadao, assai, torre_central,
    torre_barra, germans, superprix, adonai,
)

app = Flask(__name__)
CORS(app)

# Registro central de clientes: nome de exibição + função de parsing
CLIENTES = {
    'dom_atacarejo': {'nome': 'DOM Atacarejo', 'parse': dom_atacarejo.parse},
    'atacadao': {'nome': 'Atacadão', 'parse': atacadao.parse},
    'assai': {'nome': 'Assaí', 'parse': assai.parse},
    'torre_central': {'nome': 'Torre Central', 'parse': torre_central.parse},
    'torre_barra': {'nome': 'Torre Barra', 'parse': torre_barra.parse},
    'germans': {'nome': 'Germans', 'parse': germans.parse},
    'superprix': {'nome': 'Superprix', 'parse': superprix.parse},
    'adonai': {'nome': 'Adonai Atacadista', 'parse': adonai.parse},
}


@app.route('/health')
def health():
    perfis = {}
    for c in CLIENTES:
        if perfil_existe(c):
            perfis[c] = perfil_filename(c)
    return jsonify({'status': 'ok', 'perfis': perfis})


@app.route('/perfil/<cliente>', methods=['POST'])
def upload_perfil(cliente):
    """Salva ou atualiza o perfil de um cliente no servidor."""
    if cliente not in CLIENTES:
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

        # detectar se há itens de empresas diferentes (split)
        empresas_nos_itens = set(
            it.get('empresa') or dados.get('empresa', 2)
            for f in filiais for it in f['itens']
        )
        empresas_nos_itens.discard(None)
        if not empresas_nos_itens:
            empresas_nos_itens = {dados.get('empresa', 2)}

        arquivos = []
        for emp_split in sorted(empresas_nos_itens):
            override = emp_split if len(empresas_nos_itens) > 1 else None
            eb = gerar_excel(dados, empresa_override=override)
            pb = gerar_pdf(dados, empresa_override=override)
            label = ('Indústria' if emp_split == 1 else 'Distribuidora') if len(empresas_nos_itens) > 1 else ''
            arquivos.append({
                'empresa': emp_split,
                'label': label,
                'excel': base64.b64encode(eb).decode(),
                'pdf': base64.b64encode(pb).decode(),
            })

        todos_itens = [i for f in filiais for i in f['itens']]
        return jsonify({
            'ok': True,
            'split': len(empresas_nos_itens) > 1,
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
            # compatibilidade retroativa (caso simples)
            'excel': base64.b64encode(gerar_excel(dados)).decode() if len(empresas_nos_itens) == 1 else '',
            'pdf': base64.b64encode(gerar_pdf(dados)).decode() if len(empresas_nos_itens) == 1 else '',
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'erro': str(e)}), 500


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
