"""
Controle de validade — lojas Guanabara (NOVO, 13/11/2026).

Promotores terceirizados (Lider) visitam as lojas Guanabara uma vez por
semana e informam, pelo celular, os produtos com MENOS DE 30 DIAS até o
vencimento (caixas + validade do pior lote) e as rupturas. Coordenadores
da Lider (e admin/operador do BELLOTAS, sem novo login) veem o painel,
exportam Excel e cadastram promotores/lojas.

Módulo SEPARADO do resto do BELLOTAS de propósito:
- login PRÓPRIO (base de usuários e segredo de token próprios, no bucket
  pata-negra-validade). Um token da validade NUNCA vale no BELLOTAS (o
  segredo é outro); o contrário vale só pro painel admin (admin/operador).
- dados no bucket pata-negra-validade:
    usuarios.json                       -> promotores e coordenadores
    secret                              -> segredo dos tokens da validade
    visitas/<segunda AAAA-MM-DD>/<cnpj>.json -> 1 visita por loja por semana
                                           (nova visita na semana SUBSTITUI)
- produtos e lojas NÃO são cadastrados aqui: vêm dos perfis
  guanabara_lojas (5 produtos + 28 lojas) e guanabara_central (9 produtos)
  do bucket de perfis. Mudou o perfil -> o site acompanha sozinho.

Registrado no main.py com: app.register_blueprint(validade.bp)
"""
import os
import io
import re
import json
import hmac
import base64
import hashlib
import secrets
import datetime
import time

from flask import Blueprint, request, jsonify, send_file
from google.cloud import storage as gcs
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from storage import (hash_senha, verifica_senha, validar_sessao,
                     carregar_perfil_bytes, perfil_existe)
from perfil import ler_perfil, ler_filiais

bp = Blueprint('validade', __name__)

BUCKET = os.environ.get('GCS_VALIDADE_BUCKET', 'pata-negra-validade')
PERFIL_LOJAS = 'guanabara_lojas'
PERFIL_CENTRAL = 'guanabara_central'
DIAS_ALERTA = 30            # promotor reporta só o que vence em menos disso
DIAS_TOKEN = {'promotor': 30, 'coordenador': 7}
FUSO = datetime.timezone(datetime.timedelta(hours=-3))   # Brasília (sem horário de verão)
PAPEIS_BELLOTAS_ADMIN = ('admin', 'operador')

_client = None


def _bucket():
    global _client
    if _client is None:
        _client = gcs.Client()
    return _client.bucket(BUCKET)


def _ler_json(nome, padrao=None):
    b = _bucket().blob(nome)
    if not b.exists():
        return padrao
    return json.loads(b.download_as_bytes().decode('utf-8'))


def _gravar_json(nome, dados):
    _bucket().blob(nome).upload_from_string(
        json.dumps(dados, ensure_ascii=False, indent=1), content_type='application/json')


def _so_digitos(s):
    return re.sub(r'\D', '', str(s or ''))


def _agora_brt():
    return datetime.datetime.now(FUSO)


def _segunda(d):
    """Segunda-feira (date) da semana de levantamento (seg a dom) de d."""
    return d - datetime.timedelta(days=d.weekday())


# ─────────────────────────────────────────────────────────────────────
# Usuários e tokens (próprios — não são os do BELLOTAS)
# ─────────────────────────────────────────────────────────────────────
_cache_usuarios = {'t': 0, 'lista': None}


def _usuarios(fresco=False):
    if fresco or _cache_usuarios['lista'] is None or time.time() - _cache_usuarios['t'] > 30:
        _cache_usuarios['lista'] = _ler_json('usuarios.json', []) or []
        _cache_usuarios['t'] = time.time()
    return _cache_usuarios['lista']


def _salvar_usuarios(lista):
    _gravar_json('usuarios.json', lista)
    _cache_usuarios['lista'] = lista
    _cache_usuarios['t'] = time.time()


_secret_cache = None


def _secret():
    global _secret_cache
    if _secret_cache is None:
        b = _bucket().blob('secret')
        if b.exists():
            _secret_cache = b.download_as_bytes()
        else:
            _secret_cache = secrets.token_bytes(32)
            b.upload_from_string(_secret_cache, content_type='application/octet-stream')
    return _secret_cache


def _b64e(b):
    return base64.urlsafe_b64encode(b).decode('ascii').rstrip('=')


def _b64d(s):
    return base64.urlsafe_b64decode(s + '=' * (-len(s) % 4))


def _assina(p):
    return _b64e(hmac.new(_secret(), p.encode('ascii'), hashlib.sha256).digest())


def _criar_token(u):
    dias = DIAS_TOKEN.get(u['papel'], 7)
    exp = datetime.datetime.utcnow() + datetime.timedelta(days=dias)
    p = _b64e(json.dumps({'u': u['usuario'], 'p': u['papel'], 'exp': exp.isoformat()}).encode())
    return p + '.' + _assina(p)


def _validar_token(token):
    try:
        p, sig = str(token).split('.', 1)
        if not secrets.compare_digest(sig, _assina(p)):
            return None
        d = json.loads(_b64d(p).decode('utf-8'))
        if datetime.datetime.utcnow() > datetime.datetime.fromisoformat(d['exp']):
            return None
        return d
    except Exception:
        return None


def _quem():
    """Quem está chamando. Devolve dict {usuario, nome, papel, lojas, origem}
    ou None. Aceita (1) token da validade de usuário ATIVO — desativar um
    promotor corta o acesso na hora, mesmo com token ainda válido; ou (2)
    token do BELLOTAS de admin/operador (entra como 'admin')."""
    auth = request.headers.get('Authorization', '')
    token = auth[7:].strip() if auth.startswith('Bearer ') else ''
    if not token:
        return None
    d = _validar_token(token)
    if d:
        u = next((x for x in _usuarios() if x['usuario'].lower() == d['u'].lower()), None)
        if u and u.get('ativo', True):
            return {'usuario': u['usuario'], 'nome': u.get('nome') or u['usuario'],
                    'papel': u['papel'], 'lojas': u.get('lojas', []), 'origem': 'validade'}
        return None
    s = validar_sessao(token)
    if s and s.get('papel') in PAPEIS_BELLOTAS_ADMIN:
        return {'usuario': s['usuario'], 'nome': s['usuario'], 'papel': 'admin',
                'lojas': [], 'origem': 'bellotas'}
    return None


def _eh_gestor(q):
    return bool(q) and q['papel'] in ('coordenador', 'admin')


# ─────────────────────────────────────────────────────────────────────
# Catálogo: produtos e lojas vêm dos perfis Guanabara
# ─────────────────────────────────────────────────────────────────────
_cache_cat = {'t': 0, 'cat': None}


def _catalogo():
    if _cache_cat['cat'] is not None and time.time() - _cache_cat['t'] < 120:
        return _cache_cat['cat']
    produtos, lojas = [], []
    cnpj_central = ''
    vistos = set()
    for chave in (PERFIL_LOJAS, PERFIL_CENTRAL):
        if not perfil_existe(chave):
            continue
        pb = carregar_perfil_bytes(chave)
        meta, prods = ler_perfil(pb)
        if chave == PERFIL_CENTRAL:
            cnpj_central = _so_digitos(meta.get('cnpjPerfil'))
        for p in prods:
            cod = str(p['codInterno']).strip()
            if cod in vistos:
                continue
            vistos.add(cod)
            produtos.append({'cod': cod, 'nome': p['nomePerfil'], 'kgCx': p['kgCx']})
        if chave == PERFIL_LOJAS:
            for cnpj, f in ler_filiais(pb).items():
                lojas.append({'cnpj': cnpj, 'nome': f['nome'], 'numero': f['numero'],
                              'regiao': f['regiao'], 'cidade': f['cidade']})
    # a Central é o CD, não é loja (aparece na tabela de filiais com nº 0)
    lojas = [l for l in lojas if l['cnpj'] != cnpj_central and str(l['numero']).strip() not in ('0', '')]
    lojas.sort(key=lambda l: (int(l['numero']) if str(l['numero']).isdigit() else 999, l['nome']))
    for l in lojas:
        l['rotulo'] = f"{int(l['numero']):02d} · {l['nome']}" if str(l['numero']).isdigit() else l['nome']
    cat = {'produtos': produtos, 'lojas': lojas}
    _cache_cat['cat'], _cache_cat['t'] = cat, time.time()
    return cat


# ─────────────────────────────────────────────────────────────────────
# Rotas — login
# ─────────────────────────────────────────────────────────────────────
@bp.route('/validade/login', methods=['POST'])
def v_login():
    body = request.get_json(silent=True) or {}
    usuario = (body.get('usuario') or '').strip()
    senha = body.get('senha') or ''
    u = next((x for x in _usuarios(fresco=True) if x['usuario'].lower() == usuario.lower()), None)
    if not (u and u.get('ativo', True) and verifica_senha(senha, u.get('senhaHash', ''))):
        print(f'[INFO] validade login RECUSADO: "{usuario}"')
        return jsonify({'erro': 'Usuário ou senha inválidos, ou usuário desativado'}), 401
    print(f'[INFO] validade login ok: {u["usuario"]} ({u["papel"]})')
    return jsonify({'token': _criar_token(u), 'usuario': u['usuario'],
                    'nome': u.get('nome') or u['usuario'], 'papel': u['papel']})


@bp.route('/validade/me')
def v_me():
    q = _quem()
    if not q:
        return jsonify({'erro': 'sessão inválida'}), 401
    return jsonify({k: q[k] for k in ('usuario', 'nome', 'papel', 'origem')})


# ─────────────────────────────────────────────────────────────────────
# Rotas — promotor
# ─────────────────────────────────────────────────────────────────────
@bp.route('/validade/minhas')
def v_minhas():
    """Lojas do promotor (com 'visitada nesta semana') + lista de produtos.
    Gestor recebe todas as lojas (útil pra testar a tela do promotor)."""
    q = _quem()
    if not q:
        return jsonify({'erro': 'sessão inválida'}), 401
    cat = _catalogo()
    lojas = cat['lojas'] if _eh_gestor(q) else [l for l in cat['lojas'] if l['cnpj'] in q['lojas']]
    seg = _segunda(_agora_brt().date()).isoformat()
    feitas = {}
    for b in _bucket().list_blobs(prefix=f'visitas/{seg}/'):
        try:
            v = json.loads(b.download_as_bytes().decode('utf-8'))
            feitas[v['loja']] = v.get('coletadoEm') or v.get('recebidoEm')
        except Exception:
            pass
    out = [{**l, 'visitadaEm': feitas.get(l['cnpj'])} for l in lojas]
    return jsonify({'semana': seg, 'lojas': out, 'produtos': cat['produtos'],
                    'diasAlerta': DIAS_ALERTA, 'nome': q['nome']})


@bp.route('/validade/visita', methods=['POST'])
def v_visita():
    """Recebe uma visita. Pode chegar horas depois do preenchimento (fila
    offline no celular): a SEMANA é a do momento do preenchimento
    (coletadoEm), não a do envio. Nova visita da mesma loja na mesma
    semana substitui a anterior."""
    q = _quem()
    if not q:
        return jsonify({'erro': 'sessão inválida — entre de novo'}), 401
    body = request.get_json(silent=True) or {}
    cat = _catalogo()
    loja = next((l for l in cat['lojas'] if l['cnpj'] == _so_digitos(body.get('loja'))), None)
    if not loja:
        return jsonify({'erro': 'Loja desconhecida'}), 400
    if not _eh_gestor(q) and loja['cnpj'] not in q['lojas']:
        return jsonify({'erro': f'A loja {loja["rotulo"]} não está atribuída a você'}), 403

    agora = _agora_brt()
    try:
        coletado = datetime.datetime.fromisoformat(str(body.get('coletadoEm')).replace('Z', '+00:00')).astimezone(FUSO)
        if abs((agora - coletado).days) > 7:   # relógio do celular muito errado
            coletado = agora
    except Exception:
        coletado = agora

    prods = {p['cod']: p for p in cat['produtos']}
    itens, erros = [], []
    for it in body.get('itens') or []:
        p = prods.get(str(it.get('cod')))
        if not p:
            continue
        if it.get('ruptura'):
            itens.append({'cod': p['cod'], 'nome': p['nome'], 'ruptura': True})
            continue
        cx_txt, val_txt = str(it.get('caixas') or '').strip(), str(it.get('validade') or '').strip()
        if not cx_txt and not val_txt:
            continue
        try:
            cx = float(cx_txt.replace(',', '.'))
            assert cx > 0
        except Exception:
            erros.append(f'{p["nome"]}: informe a quantidade de caixas'); continue
        try:
            val = datetime.date.fromisoformat(val_txt)
        except Exception:
            erros.append(f'{p["nome"]}: informe a data de validade'); continue
        itens.append({'cod': p['cod'], 'nome': p['nome'], 'caixas': cx,
                      'kg': round(cx * p['kgCx'], 1), 'validade': val.isoformat()})
    if erros:
        return jsonify({'erro': 'Corrija: ' + '; '.join(erros)}), 400

    seg = _segunda(coletado.date()).isoformat()
    visita = {
        'loja': loja['cnpj'], 'lojaNome': loja['rotulo'],
        'promotor': q['usuario'], 'promotorNome': q['nome'],
        'coletadoEm': coletado.isoformat(timespec='seconds'),
        'recebidoEm': agora.isoformat(timespec='seconds'),
        'lat': body.get('lat'), 'lng': body.get('lng'), 'precisao': body.get('precisao'),
        'idLocal': str(body.get('idLocal') or '')[:60],
        'semana': seg, 'itens': itens,
    }
    _gravar_json(f'visitas/{seg}/{loja["cnpj"]}.json', visita)
    print(f'[INFO] validade visita: {loja["rotulo"]} por {q["usuario"]} semana {seg} '
          f'itens={len(itens)} gps={"sim" if visita["lat"] is not None else "não"}')
    return jsonify({'ok': True, 'semana': seg, 'itens': len(itens)})


# ─────────────────────────────────────────────────────────────────────
# Rotas — painel (coordenador Lider / admin e operador do BELLOTAS)
# ─────────────────────────────────────────────────────────────────────
def _dados_painel(semana_txt):
    hoje = _agora_brt().date()
    try:
        seg = _segunda(datetime.date.fromisoformat(semana_txt))
    except Exception:
        seg = _segunda(hoje)
    cat = _catalogo()
    visitas = {}
    for b in _bucket().list_blobs(prefix=f'visitas/{seg.isoformat()}/'):
        try:
            v = json.loads(b.download_as_bytes().decode('utf-8'))
            visitas[v['loja']] = v
        except Exception:
            pass
    dono = {}
    for u in _usuarios():
        if u.get('papel') == 'promotor' and u.get('ativo', True):
            for c in u.get('lojas', []):
                dono[c] = u.get('nome') or u['usuario']
    lojas = []
    for l in cat['lojas']:
        v = visitas.get(l['cnpj'])
        if v:
            for it in v['itens']:
                if it.get('validade'):
                    it['dias'] = (datetime.date.fromisoformat(it['validade']) - hoje).days
        lojas.append({**l, 'promotor': (v or {}).get('promotorNome') or dono.get(l['cnpj'], ''),
                      'visita': v})
    semanas = [(_segunda(hoje) - datetime.timedelta(weeks=i)).isoformat() for i in range(12)]
    return {'semana': seg.isoformat(), 'semanaFim': (seg + datetime.timedelta(days=6)).isoformat(),
            'semanas': semanas, 'hoje': hoje.isoformat(), 'produtos': cat['produtos'], 'lojas': lojas}


@bp.route('/validade/painel')
def v_painel():
    q = _quem()
    if not _eh_gestor(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    return jsonify(_dados_painel(request.args.get('semana', '')))


def _br(iso):
    try:
        return datetime.date.fromisoformat(iso[:10]).strftime('%d/%m/%Y')
    except Exception:
        return ''


@bp.route('/validade/painel.xlsx')
def v_painel_xlsx():
    """Mesma matriz da tela: coluna A = produtos; linhas 1-3 = loja,
    promotor, data do envio; da linha 4 em diante o que foi informado."""
    q = _quem()
    if not _eh_gestor(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    d = _dados_painel(request.args.get('semana', ''))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Validade'
    cinza = PatternFill('solid', fgColor='E9E5DF')
    verm = PatternFill('solid', fgColor='F4C9C5')
    amar = PatternFill('solid', fgColor='FBE7C2')
    preto = PatternFill('solid', fgColor='2B2622')
    cab = PatternFill('solid', fgColor='EFE9E1')
    fino = Side(style='thin', color='D0C8BD')
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)
    centro = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for r, rot in enumerate(('Loja', 'Promotor', 'Envio'), start=1):
        c = ws.cell(r, 1, rot); c.font = Font(bold=True); c.fill = cab; c.border = borda
    for i, p in enumerate(d['produtos'], start=4):
        c = ws.cell(i, 1, p['nome']); c.font = Font(bold=True); c.border = borda
    for j, l in enumerate(d['lojas'], start=2):
        v = l['visita']
        ws.cell(1, j, l['rotulo']).font = Font(bold=True)
        ws.cell(2, j, l['promotor'])
        if v:
            dt = datetime.datetime.fromisoformat(v['coletadoEm'])
            ws.cell(3, j, dt.strftime('%d/%m/%Y %H:%M'))
        else:
            ws.cell(3, j, 'sem visita').font = Font(bold=True, color='7A1F1C')
        for r in (1, 2, 3):
            c = ws.cell(r, j); c.alignment = centro; c.border = borda
            if not v:
                c.fill = cinza
        porcod = {it['cod']: it for it in (v or {}).get('itens', [])}
        for i, p in enumerate(d['produtos'], start=4):
            c = ws.cell(i, j); c.alignment = centro; c.border = borda
            if not v:
                c.fill = cinza; continue
            it = porcod.get(p['cod'])
            if not it:
                continue
            if it.get('ruptura'):
                c.value = 'RUPTURA'; c.fill = preto; c.font = Font(bold=True, color='FFFFFF')
            else:
                cx = it['caixas']; cx = int(cx) if float(cx).is_integer() else cx
                c.value = f"{cx} cx ({it['kg']:g} kg)\nval. {_br(it['validade'])}\n{it['dias']} dias"
                c.fill = verm if it['dias'] < 15 else amar
                c.font = Font(bold=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 18
    ws.column_dimensions['A'].width = 30
    for i in range(4, 4 + len(d['produtos'])):
        ws.row_dimensions[i].height = 46
    ws.freeze_panes = 'B4'
    n = len(d['produtos']) + 6
    ws.cell(n, 1, f"Semana {_br(d['semana'])} a {_br(d['semanaFim'])} · dias contados até {_br(d['hoje'])}").font = Font(italic=True)
    ws.cell(n + 1, 1, 'Vermelho: vence em menos de 15 dias · Amarelo: 15 a 29 dias · Cinza: loja sem visita · Vazio: nada a reportar').font = Font(italic=True)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f'validade_semana_{d["semana"]}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ─────────────────────────────────────────────────────────────────────
# Rotas — cadastro de promotores / coordenadores
# ─────────────────────────────────────────────────────────────────────
def _publico(u):
    return {k: u.get(k) for k in ('usuario', 'nome', 'papel', 'ativo', 'lojas')}


@bp.route('/validade/usuarios', methods=['GET'])
def v_usuarios():
    q = _quem()
    if not _eh_gestor(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    return jsonify({'usuarios': [_publico(u) for u in _usuarios(fresco=True)],
                    'lojas': _catalogo()['lojas']})


@bp.route('/validade/usuarios', methods=['POST'])
def v_usuarios_salvar():
    """Cria ou atualiza (pelo 'usuario'). Senha só muda se vier preenchida.
    Uma loja tem UM promotor: atribuir a loja a alguém tira dos outros."""
    q = _quem()
    if not _eh_gestor(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    b = request.get_json(silent=True) or {}
    usuario = re.sub(r'\s+', '', str(b.get('usuario') or '')).lower()
    if not re.fullmatch(r'[a-z0-9._-]{3,40}', usuario):
        return jsonify({'erro': 'Usuário: 3 a 40 letras/números, sem espaço (pode usar . _ -)'}), 400
    papel = b.get('papel') if b.get('papel') in ('promotor', 'coordenador') else 'promotor'
    nome = str(b.get('nome') or '').strip()
    if not nome:
        return jsonify({'erro': 'Informe o nome'}), 400
    senha = str(b.get('senha') or '')
    validas = {l['cnpj'] for l in _catalogo()['lojas']}
    lojas = [c for c in dict.fromkeys(_so_digitos(x) for x in (b.get('lojas') or [])) if c in validas]
    if papel != 'promotor':
        lojas = []

    lista = [dict(u) for u in _usuarios(fresco=True)]
    u = next((x for x in lista if x['usuario'] == usuario), None)
    if u is None:
        if len(senha) < 6:
            return jsonify({'erro': 'Senha inicial: mínimo 6 caracteres'}), 400
        u = {'usuario': usuario, 'criadoEm': _agora_brt().isoformat(timespec='seconds'),
             'criadoPor': q['usuario']}
        lista.append(u)
    elif senha and len(senha) < 6:
        return jsonify({'erro': 'Nova senha: mínimo 6 caracteres'}), 400
    u.update({'nome': nome, 'papel': papel, 'ativo': bool(b.get('ativo', True)), 'lojas': lojas})
    if senha:
        u['senhaHash'] = hash_senha(senha)
    for outro in lista:
        if outro is not u and outro.get('lojas'):
            outro['lojas'] = [c for c in outro['lojas'] if c not in lojas]
    _salvar_usuarios(lista)
    print(f'[INFO] validade cadastro: {usuario} ({papel}) lojas={len(lojas)} por {q["usuario"]}')
    return jsonify({'ok': True, 'usuario': _publico(u)})
