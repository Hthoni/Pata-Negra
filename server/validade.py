"""
Controle de validade — lojas Guanabara e Atacadão (NOVO, 13/11/2026;
REDES 16/11/2026).

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
- produtos e lojas NÃO são cadastrados aqui: vêm dos perfis de cada REDE
  (ver REDES abaixo). Mudou o perfil -> o site acompanha sozinho.
  Guanabara: guanabara_lojas (5 produtos + 28 lojas) + guanabara_central
  (9 produtos). Atacadão: atacadao (12 produtos + 30 lojas).
- REDES (16/11/2026): um promotor pode ter lojas de redes diferentes; o
  formulário de cada loja mostra os produtos da rede dela. CNPJ é único
  entre redes, então visitas/<semana>/<cnpj>.json não muda.
- RUPTURA EM 2 BOTÕES (24/11/2026): item['ruptura'] = 'geladeira' (não
  tem mais na geladeira) ou 'total' (não tem nada na loja). Ruptura NÃO
  leva caixas/validade. Visitas antigas com ruptura=True contam como
  'total' (o botão antigo era "em falta na câmara").
- ALERTA PRA REPRESENTANTE (24/11/2026): ao receber uma visita COM algo a
  reportar, manda um WhatsApp pra representante da rede (telefone da
  vendedora, célula J3 do perfil de lojas da rede), pelo mesmo envio do
  BELLOTAS (whatsapp._enviar_whatsapp). Só representantes — encarregados
  de loja NÃO recebem. O texto não cita promotor nem a empresa dele. O
  envio nunca derruba a visita: falhou, loga e segue.
- VISITA NA LOJA? (25/11/2026): compara a posição do celular com a
  lat/lng da loja no perfil. Até RAIO_LOJA_M -> 'naLoja' True (data em
  verde no painel); mais longe -> False (vermelho); sem posição ou loja
  sem coordenada -> None (neutro). Usa a posição lida no ENVIO; se
  faltar (GPS falha dentro da câmara), a lida na ABERTURA da loja. É
  calculado ao montar o painel, então corrigir a coordenada da loja no
  perfil corrige também as visitas já gravadas.
- SUPERMERCADO (19/11/2026): papel 'cliente', preso a UMA rede. Vê o
  painel e o Excel só da rede dele (todas as lojas), sem localização nem
  horário de recebimento, sem cadastro. Só admin/operador do BELLOTAS
  cria, edita ou apaga esses logins (coordenador da Lider nem os vê).
- Nome de produto EXIBIDO é simplificado (sem "RF." e sem "PATA NEGRA")
  — a marca não aparece pro promotor e o nome fica curto no celular.

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
# Cada rede: perfis de onde vêm os produtos (na ordem), perfil com a tabela
# de lojas, cores (cor = texto/botão, contraste ok com branco; cor2 = faixa),
# perfil cujo CNPJ de cabeçalho NÃO é loja (CD), e CNPJs que são a MESMA
# loja física (alias -> principal; o perfil de pedidos precisa dos dois).
REDES = {
    'guanabara': {'nome': 'Guanabara', 'perfis': ['guanabara_lojas', 'guanabara_central'],
                  'perfilLojas': 'guanabara_lojas', 'naoLoja': 'guanabara_central',
                  'cor': '#C8102E', 'cor2': '#FFD100', 'claro': '#FDECEE', 'unir': {}},
    'atacadao': {'nome': 'Atacadão', 'perfis': ['atacadao'], 'perfilLojas': 'atacadao',
                 'naoLoja': None,
                 'cor': '#00773A', 'cor2': '#F47B20', 'claro': '#FEF1E6',
                 # PILARES: mesma loja, dois CNPJs (confirmado 16/11/2026)
                 'unir': {'93209765069150': '00063960002659'}},
}
ORDEM_REDES = ['guanabara', 'atacadao']
DIAS_ALERTA = 30            # promotor reporta só o que vence em menos disso
ALERTA_REPRESENTANTE = True  # False desliga o WhatsApp pras representantes
# 27/11/2026: o mesmo alerta vai também pros encarregados DA LOJA visitada
# (colunas U-Z da linha da filial no perfil). False volta a mandar só
# pra representante.
ALERTA_ENCARREGADOS = True
RAIO_LOJA_M = 300            # até essa distância da loja a visita conta como "na loja"
DIAS_TOKEN = {'promotor': 30, 'coordenador': 7, 'cliente': 7}
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
                    'papel': u['papel'], 'lojas': u.get('lojas', []), 'rede': u.get('rede'),
                    'origem': 'validade'}
        return None
    s = validar_sessao(token)
    if s and s.get('papel') in PAPEIS_BELLOTAS_ADMIN:
        return {'usuario': s['usuario'], 'nome': s['usuario'], 'papel': 'admin',
                'lojas': [], 'origem': 'bellotas'}
    return None


def _eh_gestor(q):
    return bool(q) and q['papel'] in ('coordenador', 'admin')


def _eh_bellotas(q):
    return bool(q) and q.get('origem') == 'bellotas'


def _ve_painel(q):
    """Gestor, ou supermercado (cliente) — este só da própria rede."""
    return _eh_gestor(q) or (bool(q) and q['papel'] == 'cliente' and bool(q.get('rede')))


# ─────────────────────────────────────────────────────────────────────
# Catálogo: produtos e lojas vêm dos perfis Guanabara
# ─────────────────────────────────────────────────────────────────────
_cache_cat = {'t': 0, 'cat': None}


def _nome_curto(nome):
    """Nome exibido: sem 'RF.' no início, sem 'PATA NEGRA', espaços simples,
    sem ponto sobrando no fim ('RF.RABO SUINO SAL.PATA NEGRA' -> 'RABO SUINO SAL')."""
    n = re.sub(r'^\s*RF\.?\s*', '', str(nome or ''), flags=re.I)
    n = re.sub(r'PATA\s*NEGRA', ' ', n, flags=re.I)
    n = re.sub(r'\s+', ' ', n).strip(' .-')
    return n or str(nome or '').strip()


def _rede_meta(r):
    c = REDES[r]
    return {'id': r, 'nome': c['nome'], 'cor': c['cor'], 'cor2': c['cor2'], 'claro': c['claro']}


def _catalogo():
    """{'redes': [meta...], 'produtos': {rede: [...]}, 'lojas': [... com 'rede'],
    'alias': {cnpj_alias: cnpj_principal}}. Cache de 2 min."""
    if _cache_cat['cat'] is not None and time.time() - _cache_cat['t'] < 120:
        return _cache_cat['cat']
    produtos, lojas, alias, redes, reps, encs = {}, [], {}, [], {}, {}
    for r in ORDEM_REDES:
        cfg = REDES[r]
        prods, vistos, nao_loja, filiais = [], set(), '', {}
        for chave in cfg['perfis']:
            if not perfil_existe(chave):
                print(f'[WARN] validade: perfil "{chave}" da rede {r} não encontrado no bucket')
                continue
            pb = carregar_perfil_bytes(chave)
            meta, ps = ler_perfil(pb)
            if chave == cfg.get('naoLoja'):
                nao_loja = _so_digitos(meta.get('cnpjPerfil'))
            for p in ps:
                cod = str(p['codInterno']).strip()
                if cod in vistos:
                    continue
                vistos.add(cod)
                prods.append({'cod': cod, 'nome': _nome_curto(p['nomePerfil']), 'kgCx': p['kgCx']})
            if chave == cfg['perfilLojas']:
                filiais = ler_filiais(pb)
                reps[r] = {'nome': meta.get('vendedor', ''), 'telefone': meta.get('telefone', '')}
        if not prods and not filiais:
            continue
        redes.append(_rede_meta(r))
        produtos[r] = prods
        unir = cfg.get('unir') or {}
        alias.update(unir)
        ls = []
        # Encarregados por loja. Ficam num dicionário à parte (cat['encs']) e
        # NÃO dentro de cada loja: a lista de lojas é devolvida ao navegador
        # (painel, promotor) e telefone de encarregado não pode ir junto.
        # Loja com dois CNPJs (PILARES) junta os encarregados das duas linhas.
        for cnpj, f in filiais.items():
            alvo = unir.get(cnpj, cnpj)
            for e in (f.get('encarregados') or []):
                if str(e.get('telefone') or '').strip():
                    encs.setdefault(alvo, []).append({'nome': e.get('nome', ''), 'telefone': e['telefone']})
        for cnpj, f in filiais.items():
            if cnpj in unir or cnpj == nao_loja or str(f['numero']).strip() in ('0', ''):
                continue   # alias da mesma loja, CD, ou linha sem número
            ls.append({'cnpj': cnpj, 'rede': r, 'nome': f['nome'], 'numero': f['numero'],
                       'regiao': f['regiao'], 'cidade': f['cidade'],
                       'lat': f.get('lat'), 'lng': f.get('lng')})
        ls.sort(key=lambda l: (int(l['numero']) if str(l['numero']).isdigit() else 9999, l['nome']))
        for l in ls:
            l['rotulo'] = f"{int(l['numero']):02d} · {l['nome']}" if str(l['numero']).isdigit() else l['nome']
        lojas.extend(ls)
    cat = {'redes': redes, 'produtos': produtos, 'lojas': lojas, 'alias': alias, 'reps': reps, 'encs': encs}
    _cache_cat['cat'], _cache_cat['t'] = cat, time.time()
    return cat


def _cnpj_loja(cnpj, cat):
    c = _so_digitos(cnpj)
    return cat['alias'].get(c, c)


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
    out = {k: q[k] for k in ('usuario', 'nome', 'papel', 'origem')}
    if q['papel'] == 'cliente':
        out['rede'] = q.get('rede')
    return jsonify(out)


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
    if q['papel'] == 'cliente':
        return jsonify({'erro': 'Este login é do painel do supermercado, não do promotor.'}), 403
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
    out = [{**{k: x for k, x in l.items() if k not in ('lat', 'lng')},
            'visitadaEm': feitas.get(l['cnpj'])} for l in lojas]
    return jsonify({'semana': seg, 'lojas': out, 'redes': cat['redes'],
                    'produtosPorRede': cat['produtos'],
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
    if q['papel'] == 'cliente':
        return jsonify({'erro': 'Este login não envia visitas'}), 403
    body = request.get_json(silent=True) or {}
    cat = _catalogo()
    cnpj = _cnpj_loja(body.get('loja'), cat)
    loja = next((l for l in cat['lojas'] if l['cnpj'] == cnpj), None)
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

    prods = {p['cod']: p for p in cat['produtos'].get(loja['rede'], [])}
    itens, erros = [], []
    for it in body.get('itens') or []:
        p = prods.get(str(it.get('cod')))
        if not p:
            continue
        rup = it.get('ruptura')
        if rup:
            itens.append({'cod': p['cod'], 'nome': p['nome'],
                          'ruptura': 'geladeira' if rup == 'geladeira' else 'total'})
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
        'loja': loja['cnpj'], 'lojaNome': loja['rotulo'], 'rede': loja['rede'],
        'promotor': q['usuario'], 'promotorNome': q['nome'],
        'coletadoEm': coletado.isoformat(timespec='seconds'),
        'recebidoEm': agora.isoformat(timespec='seconds'),
        'lat': _num(body.get('lat')), 'lng': _num(body.get('lng')), 'precisao': _num(body.get('precisao')),
        # posição lida quando o promotor ABRIU a loja no celular (reserva)
        'latAb': _num(body.get('latAb')), 'lngAb': _num(body.get('lngAb')), 'precAb': _num(body.get('precAb')),
        'idLocal': str(body.get('idLocal') or '')[:60],
        'semana': seg, 'itens': itens,
    }
    caminho = f'visitas/{seg}/{loja["cnpj"]}.json'
    anterior = _ler_json(caminho)            # visita da mesma loja já enviada nesta semana (ou None)
    ja_havia = anterior is not None
    _gravar_json(caminho, visita)
    print(f'[INFO] validade visita: {loja["rede"]} {loja["rotulo"]} por {q["usuario"]} semana {seg} '
          f'itens={len(itens)} gps={"sim" if visita["lat"] is not None else "não"}')
    _alerta_representante(loja, visita, cat, atualizacao=ja_havia,
                          anterior_tinha_itens=bool((anterior or {}).get('itens')))
    return jsonify({'ok': True, 'semana': seg, 'itens': len(itens)})


# ─────────────────────────────────────────────────────────────────────
# Alerta por WhatsApp pra representante da rede
# ─────────────────────────────────────────────────────────────────────
_ROMANOS = {'I', 'II', 'III', 'IV', 'V', 'VI'}


def _nome_loja_msg(loja):
    """'BANGU I' nº 5 -> 'Bangu I F5' (romanos continuam maiúsculos)."""
    palavras = [w if w.upper() in _ROMANOS else w.capitalize() for w in str(loja['nome']).split()]
    palavras = [w.upper() if w.upper() in _ROMANOS else w for w in palavras]
    return f"{' '.join(palavras)} F{loja['numero']}"


def _kg_txt(kg):
    kg = float(kg or 0)
    return (f'{kg:.0f}' if kg.is_integer() else f'{kg:.1f}'.replace('.', ',')) + ' kg'


def montar_alerta_visita(loja, visita, atualizacao=False):
    """Texto do alerta, no modelo combinado em 24/11/2026: só o nome da
    loja + nº da filial em negrito; uma linha por lote, em kg; rupturas
    no fim. Devolve '' quando a visita não tem nada a reportar."""
    hoje = _agora_brt().date()
    lotes, rup_tot, rup_gel = [], [], []
    for it in visita.get('itens', []):
        nome = str(it.get('nome', '')).lower()
        if it.get('ruptura') == 'geladeira':
            rup_gel.append(nome)
        elif it.get('ruptura'):
            rup_tot.append(nome)
        elif it.get('validade'):
            dias = (datetime.date.fromisoformat(it['validade']) - hoje).days
            lotes.append((dias, _kg_txt(it.get('kg')), nome))
    if not (lotes or rup_tot or rup_gel):
        return ''
    # 25/11/2026: uma linha em branco entre os blocos (vencimento / ruptura
    # total / ruptura geladeira). O WhatsApp via Botconversa engole "\n\n"
    # puro; o espaço de largura zero no meio segura a linha em branco
    # (mesmo truque do _QUEBRA do whatsapp.py).
    quebra = '\n\u200b\n'
    blocos = []
    venc = []
    for dias, kg, nome in sorted(lotes):
        if dias < 0:
            venc.append(f'{kg} de {nome} — VENCIDOS')
        elif dias == 0:
            venc.append(f'{kg} de {nome} — vence HOJE')
        else:
            venc.append(f'{kg} de {nome} com {dias} dia{"s" if dias != 1 else ""} para vencer')
    # 25/11/2026 (2): cada lote vira um bloco próprio -> linha em branco
    # entre um vencimento e outro.
    # 27/11/2026: frase de atenção antes dos lotes (só quando há lote).
    if venc:
        blocos.append('Atenção! Mercadorias com validade curta na câmara frigorífica:')
    blocos.extend(venc)
    if rup_tot:
        blocos.append('Ruptura total: ' + ', '.join(rup_tot))
    if rup_gel:
        blocos.append('Ruptura geladeira: ' + ', '.join(rup_gel))
    # 27/11/2026 (2): título passa a ser "Relatório <loja>".
    titulo = f'*Relatório {_nome_loja_msg(loja)}*' + (' (atualizado)' if atualizacao else '')
    # 25/11/2026 (2): linha em branco também logo depois do nome da loja.
    return titulo + quebra + quebra.join(blocos)


def _boas_vindas_encarregado(nome):
    """Apresentação mandada UMA vez ao encarregado, no primeiro contato
    (texto do Henrique, 28/11/2026). '\n\u200b\n' = linha em branco."""
    q = '\n\u200b\n'
    n = str(nome or '').strip()
    return q.join([
        f'Caro {n},' if n else 'Olá,',
        'Este é o número de WhatsApp da Pata Negra, para alertá-lo de eventos de mercadorias '
        'com pouco prazo de validade e/ou rupturas.',
        'Acreditamos que a informação sempre fresca, e na mão, pode melhorar muito nossos processos.',
        'Não se preocupe, será usado somente quando ocorrerem os problemas acima, mesmo assim, '
        'no máximo uma vez por semana. Não vamos encher vc de mensagens... rsrsrs',
        'Muito obrigado!!',
    ])


def _enviar_alerta(loja, texto, cat, destino='todos', n_itens=0, atualizacao=False):
    """Manda o texto pros destinatários da loja e devolve o que aconteceu
    com cada um: [{'papel','nome','resultado'}]. destino: 'todos' |
    'representante' | 'encarregados'. Mesmo telefone em dois papéis recebe
    uma vez só. Usado pelo envio da visita e pelo botão "Reenviar alerta"."""
    import whatsapp
    rede = loja['rede']
    dest, vistos, saida = [], set(), []
    if destino in ('todos', 'representante'):
        rep_ = (cat.get('reps') or {}).get(rede) or {}
        tel = str(rep_.get('telefone') or '').strip()
        if tel:
            dest.append(('representante', rep_.get('nome', ''), tel))
        else:
            print(f'[WARN] validade alerta: representante da rede {rede} SEM telefone no perfil (J3) ({loja["rotulo"]})')
            saida.append({'papel': 'representante', 'nome': rep_.get('nome', ''), 'resultado': 'sem_telefone'})
    if destino in ('todos', 'encarregados') and ALERTA_ENCARREGADOS:
        lista = (cat.get('encs') or {}).get(loja['cnpj']) or []
        if not lista:
            print(f'[INFO] validade alerta: loja sem encarregado com telefone no perfil — {rede} {loja["rotulo"]}')
        for e in lista:
            dest.append(('encarregado', e.get('nome', ''), str(e['telefone']).strip()))
    if not dest:
        print(f'[WARN] validade alerta: ninguém pra avisar — mensagem NÃO enviada ({rede} {loja["rotulo"]})')
        return saida
    for papel, nome, tel in dest:
        chave = whatsapp._normaliza_telefone(tel)
        if chave in vistos:
            continue
        vistos.add(chave)
        # um try por destinatário: falha num número não segura os outros
        try:
            print(f'[INFO] validade alerta: enviando pra {papel} {nome} ({tel}) — {rede} {loja["rotulo"]} '
                  f'({n_itens} itens{", atualização" if atualizacao else ""})')
            # 27/11/2026 (2): encarregado recebe a saudação com o nome dele
            # (como está no perfil) antes do relatório; representante não.
            msg = texto
            if papel == 'encarregado' and str(nome or '').strip():
                msg = f'Caro {str(nome).strip()},\n\u200b\n' + texto
            # 28/11/2026 (2): passa o nome -> se o telefone ainda não é contato
            # no Botconversa, o whatsapp.py cria o contato e envia.
            # 28/11/2026 (3): encarregado em PRIMEIRO contato recebe antes a
            # apresentação (só quando o contato é criado agora).
            bv = _boas_vindas_encarregado(nome) if papel == 'encarregado' else None
            res = whatsapp._enviar_whatsapp(tel, msg, nome=nome, boas_vindas=bv) or 'enviado'
        except Exception as e:
            print(f'[WARN] validade alerta: falha ao enviar pra {papel} {nome} ({tel}) — {rede} {loja["rotulo"]}: {e}')
            res = 'falha'
        saida.append({'papel': papel, 'nome': nome, 'resultado': res})
    return saida


def _alerta_representante(loja, visita, cat, atualizacao=False, anterior_tinha_itens=False):
    """Manda o alerta pra representante da rede da loja. Nunca levanta
    erro (a visita já está gravada) e SEMPRE loga o que aconteceu."""
    rede = loja['rede']
    try:
        if not ALERTA_REPRESENTANTE:
            print(f'[INFO] validade alerta: desligado (ALERTA_REPRESENTANTE=False) — {rede} {loja["rotulo"]}')
            return
        texto = montar_alerta_visita(loja, visita, atualizacao)
        if not texto and atualizacao and anterior_tinha_itens:
            # 25/11/2026: a representante já recebeu um alerta desta loja e o
            # promotor reenviou a visita SEM pendências -> avisa que zerou,
            # senão ela fica com um alerta que não vale mais.
            texto = f'*Relatório {_nome_loja_msg(loja)}* (atualizado)\n\u200b\nSem rupturas nem produtos por vencer.'
        if not texto:
            print(f'[INFO] validade alerta: nada a reportar, sem mensagem — {rede} {loja["rotulo"]}')
            return
        _enviar_alerta(loja, texto, cat, n_itens=len(visita.get('itens', [])), atualizacao=atualizacao)
    except Exception as e:
        print(f'[WARN] validade alerta: falha ao enviar ({rede} {loja["rotulo"]}): {e}')


@bp.route('/validade/alerta/reenviar', methods=['POST'])
def v_alerta_reenviar():
    """28/11/2026: manda de novo o relatório ATUAL de uma loja (visita da
    semana pedida). Só pra quem entra pelo BELLOTAS. Serve pra quando o
    contato foi cadastrado no Botconversa depois do envio, ou alguém diz
    que não recebeu. Devolve o resultado por destinatário."""
    q = _quem()
    if not q:
        return jsonify({'erro': 'sessão inválida — entre de novo'}), 401
    if not _eh_bellotas(q):
        return jsonify({'erro': 'Só o login do BELLOTAS reenvia alertas'}), 403
    body = request.get_json(silent=True) or {}
    destino = body.get('destino') if body.get('destino') in ('todos', 'representante', 'encarregados') else 'todos'
    cat = _catalogo()
    cnpj = _cnpj_loja(body.get('loja'), cat)
    loja = next((l for l in cat['lojas'] if l['cnpj'] == cnpj), None)
    if not loja:
        return jsonify({'erro': 'Loja desconhecida'}), 400
    try:
        seg = _segunda(datetime.date.fromisoformat(str(body.get('semana')))).isoformat()
    except Exception:
        seg = _segunda(_agora_brt().date()).isoformat()
    visita = _ler_json(f'visitas/{seg}/{loja["cnpj"]}.json')
    if not visita:
        return jsonify({'erro': 'Esta loja não tem visita nessa semana'}), 400
    texto = montar_alerta_visita(loja, visita)
    if not texto:
        return jsonify({'erro': 'A visita desta loja não tem ruptura nem validade curta — não há alerta a enviar'}), 400
    print(f'[INFO] validade alerta: REENVIO pedido por {q["usuario"]} — {loja["rede"]} {loja["rotulo"]} semana {seg} destino={destino}')
    res = _enviar_alerta(loja, texto, cat, destino=destino, n_itens=len(visita.get('itens', [])))
    return jsonify({'ok': True, 'loja': loja['rotulo'], 'resultados': res})


# ─────────────────────────────────────────────────────────────────────
# Rotas — painel (coordenador Lider / admin e operador do BELLOTAS)
# ─────────────────────────────────────────────────────────────────────
def _dados_painel(semana_txt, rede):
    hoje = _agora_brt().date()
    try:
        seg = _segunda(datetime.date.fromisoformat(semana_txt))
    except Exception:
        seg = _segunda(hoje)
    cat = _catalogo()
    if rede not in cat['produtos']:
        rede = cat['redes'][0]['id'] if cat['redes'] else rede
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
        if l['rede'] != rede:
            continue
        v = visitas.get(l['cnpj'])
        if v:
            for it in v['itens']:
                if it.get('validade'):
                    it['dias'] = (datetime.date.fromisoformat(it['validade']) - hoje).days
            v.update(_local_da_visita(v, l))
        lojas.append({**l, 'promotor': (v or {}).get('promotorNome') or dono.get(l['cnpj'], ''),
                      'visita': v})
    semanas = [(_segunda(hoje) - datetime.timedelta(weeks=i)).isoformat() for i in range(12)]
    return {'semana': seg.isoformat(), 'semanaFim': (seg + datetime.timedelta(days=6)).isoformat(),
            'semanas': semanas, 'hoje': hoje.isoformat(), 'rede': rede, 'redes': cat['redes'],
            'produtos': cat['produtos'].get(rede, []), 'lojas': lojas}


def _num(v):
    try:
        return float(v) if v is not None and v != '' else None
    except (TypeError, ValueError):
        return None


def _distancia_m(lat1, lng1, lat2, lng2):
    """Distância em metros entre dois pontos (haversine)."""
    import math
    r = 6371000.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lng2 - lng1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def _local_da_visita(v, loja):
    """{'naLoja': True|False|None, 'distLojaM': m|None, 'locOrigem':
    'envio'|'abertura'|None, 'raioM'} — ver bloco VISITA NA LOJA? no topo."""
    out = {'naLoja': None, 'distLojaM': None, 'locOrigem': None, 'raioM': RAIO_LOJA_M}
    lat, lng, origem = v.get('lat'), v.get('lng'), 'envio'
    if lat is None or lng is None:
        lat, lng, origem = v.get('latAb'), v.get('lngAb'), 'abertura'
    if lat is None or lng is None:
        return out
    out['locOrigem'] = origem
    if loja.get('lat') is None or loja.get('lng') is None:
        return out
    d = _distancia_m(float(lat), float(lng), float(loja['lat']), float(loja['lng']))
    out['distLojaM'] = round(d)
    out['naLoja'] = d <= RAIO_LOJA_M
    return out


def _sem_rastreio(d):
    """Supermercado não vê localização nem horário de recebimento."""
    for l in d['lojas']:
        l.pop('lat', None); l.pop('lng', None)
        v = l.get('visita')
        if v:
            for k in ('lat', 'lng', 'precisao', 'latAb', 'lngAb', 'precAb', 'recebidoEm', 'idLocal',
                      'naLoja', 'distLojaM', 'locOrigem', 'raioM'):
                v.pop(k, None)
    d['redes'] = [r for r in d['redes'] if r['id'] == d['rede']]
    return d


def v_rede_pedida(q):
    return q['rede'] if q['papel'] == 'cliente' else request.args.get('rede', 'guanabara')


@bp.route('/validade/painel')
def v_painel():
    q = _quem()
    if not _ve_painel(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    d = _dados_painel(request.args.get('semana', ''), v_rede_pedida(q))
    return jsonify(_sem_rastreio(d) if q['papel'] == 'cliente' else d)


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
    if not _ve_painel(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    semana = request.args.get('semana', '')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    d = None
    redes = _catalogo()['redes']
    if q['papel'] == 'cliente':
        redes = [r for r in redes if r['id'] == q['rede']]
    for rm in redes:
        d = _dados_painel(semana, rm['id'])
        ws = wb.create_sheet(rm['nome'][:31])
        ws.sheet_properties.tabColor = rm['cor'].lstrip('#')
        _aba_excel(ws, d)
    if d is None:
        return jsonify({'erro': 'nenhuma rede com perfil cadastrado'}), 404
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f'validade_semana_{d["semana"]}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def _aba_excel(ws, d):
    cinza = PatternFill(patternType='lightUp', fgColor='C9CFD6', bgColor='F4F6F8')   # sem visita: listrado
    verm = PatternFill('solid', fgColor='F4C9C5')
    amar = PatternFill('solid', fgColor='FBE7C2')
    rup_gel = PatternFill('solid', fgColor='D5D9DF')
    rup_tot = PatternFill('solid', fgColor='9AA1AA')
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
            c3 = ws.cell(3, j, dt.strftime('%d/%m/%Y %H:%M'))
            if v.get('naLoja') is True:
                c3.font = Font(bold=True, color='1F7A3D')     # enviada na loja
            elif v.get('naLoja') is False:
                c3.font = Font(bold=True, color='B71C1C')     # enviada longe da loja
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
                gel = it['ruptura'] == 'geladeira'
                c.value = 'RUPTURA\nGELADEIRA' if gel else 'RUPTURA\nTOTAL'
                c.fill = rup_gel if gel else rup_tot
                c.font = Font(bold=True, color='111111')
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
    ws.cell(n + 1, 1, 'Rosa: vence em menos de 15 dias · Bege: 15 a 29 dias · Cinza claro: RUPTURA GELADEIRA · Cinza escuro: RUPTURA TOTAL · Listrado: loja sem visita · Vazio: nada a reportar').font = Font(italic=True)


# ─────────────────────────────────────────────────────────────────────
# Rotas — cadastro de promotores / coordenadores
# ─────────────────────────────────────────────────────────────────────
def _publico(u):
    return {k: u.get(k) for k in ('usuario', 'nome', 'papel', 'ativo', 'lojas', 'rede')}


@bp.route('/validade/usuarios', methods=['GET'])
def v_usuarios():
    q = _quem()
    if not _eh_gestor(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    cat = _catalogo()
    lista = [u for u in _usuarios(fresco=True) if _eh_bellotas(q) or u.get('papel') != 'cliente']
    return jsonify({'usuarios': [_publico(u) for u in lista], 'podeSupermercado': _eh_bellotas(q),
                    'lojas': cat['lojas'], 'redes': cat['redes']})


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
    papel = b.get('papel') if b.get('papel') in ('promotor', 'coordenador', 'cliente') else 'promotor'
    existente = next((x for x in _usuarios(fresco=True) if x['usuario'] == usuario), None)
    if (papel == 'cliente' or (existente and existente.get('papel') == 'cliente')) and not _eh_bellotas(q):
        return jsonify({'erro': 'Logins de supermercado só podem ser criados ou alterados pelo BELLOTAS'}), 403
    rede = None
    if papel == 'cliente':
        rede = b.get('rede')
        if rede not in REDES:
            return jsonify({'erro': 'Escolha a rede do supermercado'}), 400
    # 16/11/2026: cadastro só com login — o login é o nome exibido. Usuário
    # antigo que tinha nome mantém o nome se o formulário não mandar outro.
    nome = str(b.get('nome') or '').strip()
    senha = str(b.get('senha') or '')
    cat = _catalogo()
    validas = {l['cnpj'] for l in cat['lojas']}
    lojas = [c for c in dict.fromkeys(_cnpj_loja(x, cat) for x in (b.get('lojas') or [])) if c in validas]
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
    u.update({'nome': nome or u.get('nome') or usuario, 'papel': papel,
              'ativo': bool(b.get('ativo', True)), 'lojas': lojas})
    if papel == 'cliente':
        u['rede'] = rede
    else:
        u.pop('rede', None)
    if senha:
        u['senhaHash'] = hash_senha(senha)
    for outro in lista:
        if outro is not u and outro.get('lojas'):
            outro['lojas'] = [c for c in outro['lojas'] if c not in lojas]
    _salvar_usuarios(lista)
    print(f'[INFO] validade cadastro: {usuario} ({papel}) lojas={len(lojas)} por {q["usuario"]}')
    return jsonify({'ok': True, 'usuario': _publico(u)})


@bp.route('/validade/usuarios/apagar', methods=['POST'])
def v_usuarios_apagar():
    """APAGA um usuário de vez (19/11/2026). As lojas dele ficam sem
    promotor; as visitas já enviadas continuam no histórico (cada visita
    guarda o login de quem enviou). Ninguém apaga a si mesmo."""
    q = _quem()
    if not _eh_gestor(q):
        return jsonify({'erro': 'acesso restrito'}), 403
    usuario = str((request.get_json(silent=True) or {}).get('usuario') or '').strip().lower()
    if q['origem'] == 'validade' and usuario == q['usuario'].lower():
        return jsonify({'erro': 'Você não pode apagar o seu próprio usuário'}), 400
    lista = _usuarios(fresco=True)
    alvo = next((u for u in lista if u['usuario'].lower() == usuario), None)
    if alvo and alvo.get('papel') == 'cliente' and not _eh_bellotas(q):
        return jsonify({'erro': 'Logins de supermercado só podem ser apagados pelo BELLOTAS'}), 403
    nova = [u for u in lista if u['usuario'].lower() != usuario]
    if len(nova) == len(lista):
        return jsonify({'erro': 'Usuário não encontrado'}), 404
    _salvar_usuarios(nova)
    print(f'[INFO] validade cadastro: APAGADO {usuario} por {q["usuario"]}')
    return jsonify({'ok': True})
