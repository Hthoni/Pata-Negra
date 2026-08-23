"""
Canal WhatsApp — roteamento por telefone e montagem das mensagens de
status/menu, conforme a Planta (planta_projeto_whatsapp.md) e o rascunho
original (rascunho_projeto_alertas_whatsapp.docx, 21/08/2026).

Ponto de entrada único (main.py: POST /whatsapp/webhook) recebe SÓ
{"telefone": "..."} do Botconversa (Bloco de Integração, botão "status
da entrega" no fluxo) e devolve {"mensagem": "..."} — nenhuma decisão é
tomada do lado do Botconversa, tudo aqui.

ESCOPO DESTA ENTREGA: a mensagem de ENTRADA — o que a pessoa recebe ao
clicar no botão / na primeira interação. Cobre:
  - vendedor  -> lista de status de todas as entregas dele, hoje
  - encarregado -> status da(s) entrega(s) da filial dele, hoje
  - motorista -> menu de 3 opções (chegada / conclusão / falha)
  - desconhecido -> mensagem padrão

AINDA PENDENTE (fora do escopo desta entrega, ver Planta seção 4 e 7):
  - Os passos SEGUINTES do motorista (escolher o número da entrega,
    confirmar SIM/NÃO) e o cadastro do encarregado desconhecido
    dependem de uma SESSÃO por telefone (pra saber "em que ponto da
    conversa" a pessoa está) — isso por sua vez depende de resolver a
    captura de "texto" no Botconversa (Bloco de Conteúdo + SALVAR),
    ainda não fechado do lado de lá.
  - A cascata de notificação pro vendedor/encarregado quando o
    motorista confirma chegada/conclusão/falha (fim do docx original).
"""
import re
import datetime
from storage import listar_romaneios, listar_entregas

_STATUS_LEGIVEL = {
    'em_rota': 'Em rota',
    'entregue': 'Entregue',
    'falhou': 'Falha na entrega',
}


def _normaliza_telefone(tel):
    """Só dígitos, com DDI 55 se faltar — mesmo padrão usado em avulso.py
    e no restante do projeto, pra bater com o formato que o WhatsApp usa
    (DDI+DDD+9+TELEFONE)."""
    d = re.sub(r'\D', '', str(tel or ''))
    if d and not d.startswith('55'):
        d = '55' + d
    return d


def _status_legivel(romaneio):
    if romaneio.get('chegouLocal'):
        return 'No local de entrega'
    return _STATUS_LEGIVEL.get(romaneio.get('status'), romaneio.get('status', ''))


# FIX (23/08/2026): _pedidos_ativos_hoje() e _entregas_idx() liam o
# bucket INTEIRO de romaneios/entregas toda vez que eram chamadas — e
# várias funções (identificar_papel, montar_mensagem_vendedor,
# montar_mensagem_encarregado, etc.) chamavam as duas de novo, do zero,
# dentro do MESMO request. Uma única mensagem de WhatsApp podia disparar
# 4-6 leituras completas do bucket pra responder. Agora tudo isso é
# cacheado por request, via _contexto_pedido() abaixo — chamado uma vez
# só (ou reaproveitado, se já foi chamado antes na mesma execução).
_cache_contexto = {}


def _contexto_pedido():
    """Lê romaneios + entregas do bucket UMA vez por execução da função
    (processo), reaproveitando entre todas as chamadas dentro do mesmo
    request. Evita reler o bucket inteiro várias vezes pra responder uma
    única mensagem."""
    if _cache_contexto:
        return _cache_contexto['romaneios'], _cache_contexto['entregas']
    romaneios = listar_romaneios()
    entregas = listar_entregas()
    _cache_contexto['romaneios'] = romaneios
    _cache_contexto['entregas'] = entregas
    return romaneios, entregas


def _limpar_cache_contexto():
    """Chamar depois de qualquer escrita (salvar_romaneio, atualizar
    status etc.) pra não servir dado desatualizado dentro do mesmo
    request, e no INÍCIO de cada request novo (main.py deve chamar isso
    antes de processar_mensagem_entrada)."""
    _cache_contexto.clear()


def _pedidos_ativos_hoje():
    """Pedidos de entregas DESPACHADAS hoje (fase='em_rota', despachadaEm
    de hoje) e que ainda estão com status='em_rota' (não entregues/
    falhados). 'Hoje' é o dia do despacho, não o dia em que o pedido foi
    originalmente processado."""
    romaneios, entregas = _contexto_pedido()
    hoje_iso = datetime.date.today().isoformat()
    entregas_hoje = [
        e for e in entregas
        if e.get('fase') == 'em_rota' and (e.get('despachadaEm') or '').startswith(hoje_iso)
    ]
    ids_hoje = set()
    for e in entregas_hoje:
        ids_hoje.update(e.get('pedidoIds', []))
    idx = {r['id']: r for r in romaneios}
    return [idx[pid] for pid in ids_hoje if pid in idx and idx[pid].get('status') == 'em_rota']


def _entregas_idx():
    _, entregas = _contexto_pedido()
    return {e['id']: e for e in entregas}


def _motorista_da_entrega(romaneio, idx_entregas):
    eid = romaneio.get('entregaId')
    e = idx_entregas.get(eid) if eid else None
    if e:
        return e.get('motorista', '—'), e.get('telefoneMotorista', '—')
    return '—', '—'


# ── identificação de papel (busca reversa por telefone) ──────────────
def identificar_papel(telefone):
    """Telefone -> ('vendedor'|'encarregado'|'motorista'|None, {dados}).
    Busca só entre pedidos/entregas ATIVOS HOJE (ver _pedidos_ativos_hoje)
    — vendedor/encarregado de um pedido de ontem que já saiu da rota não
    deveria mais responder a esse número por essa via."""
    alvo = _normaliza_telefone(telefone)
    pedidos = _pedidos_ativos_hoje()

    for r in pedidos:
        if _normaliza_telefone(r.get('telefoneVendedor', '')) == alvo:
            return 'vendedor', {'nome': r.get('vendedor', '')}
        for enc in (r.get('encarregados') or []):
            if _normaliza_telefone(enc.get('telefone', '')) == alvo:
                return 'encarregado', {'nome': enc.get('nome', '')}

    _, entregas = _contexto_pedido()
    for e in entregas:
        if _normaliza_telefone(e.get('telefoneMotorista', '')) == alvo:
            if e.get('fase') == 'em_rota':  # só motorista com entrega despachada hoje/ativa
                return 'motorista', {'nome': e.get('motorista', ''), 'entregaId': e.get('id')}

    return None, {}


# ── montagem de mensagem por papel (texto conforme o docx original) ──
def montar_mensagem_vendedor(nome_vendedor, telefone_alvo):
    pedidos = [r for r in _pedidos_ativos_hoje()
               if _normaliza_telefone(r.get('telefoneVendedor', '')) == telefone_alvo]
    if not pedidos:
        return f'Olá {nome_vendedor}, não encontrei nenhuma entrega em rota hoje pra você.'

    idx_ent = _entregas_idx()
    por_motorista = {}
    for r in pedidos:
        chave = _motorista_da_entrega(r, idx_ent)
        por_motorista.setdefault(chave, []).append(r)

    linhas = [f'Olá {nome_vendedor}, segue a relação de entregas e seus status:']
    for (mot_nome, mot_tel), lista in por_motorista.items():
        linhas.append('')
        linhas.append(f'Motorista {mot_nome} telefone - {mot_tel}')
        for r in lista:
            cliente = r.get('clienteNome') or r.get('cliente') or ''
            filial = r.get('filial', '')
            status = _status_legivel(r)
            encs = r.get('encarregados') or []
            if encs:
                enc_txt = f"{encs[0].get('nome', '—')} / {encs[0].get('telefone', '—')}"
            else:
                enc_txt = '— / —'
            linhas.append(f'{cliente} / {filial} / {status} / {enc_txt}')
    linhas.append('')
    linhas.append('Agradecemos toda ajuda no desembarque!')
    return '\n'.join(linhas)


def montar_mensagem_encarregado(nome_encarregado, telefone_alvo):
    """FIX (23/08/2026): reescrita pra agrupar por motorista (mesmo
    padrão já usado em montar_mensagem_vendedor) — antes repetia
    'Motorista ... telefone' e linhas em branco pra CADA pedido, ficando
    poluído quando havia mais de um pedido ativo pra mesma filial."""
    pedidos = [r for r in _pedidos_ativos_hoje()
               if any(_normaliza_telefone(e.get('telefone', '')) == telefone_alvo
                      for e in (r.get('encarregados') or []))]
    if not pedidos:
        return f'Olá {nome_encarregado}, não encontrei nenhuma entrega em rota hoje pra sua filial.'

    idx_ent = _entregas_idx()
    por_motorista = {}
    for r in pedidos:
        chave = _motorista_da_entrega(r, idx_ent)
        por_motorista.setdefault(chave, []).append(r)

    linhas = [f'Olá {nome_encarregado}, segue o status:']
    for (mot_nome, mot_tel), lista in por_motorista.items():
        linhas.append('')
        linhas.append(f'Motorista {mot_nome} telefone - {mot_tel}')
        for r in lista:
            cliente = r.get('clienteNome') or r.get('cliente') or ''
            filial = r.get('filial', '')
            status = _status_legivel(r)
            vend_nome = r.get('vendedor', '—')
            vend_tel = r.get('telefoneVendedor', '—')
            linhas.append(f'{cliente} / {filial} / {status} / Vendedor {vend_nome} - {vend_tel}')
    linhas.append('')
    linhas.append('Agradecemos toda ajuda no desembarque!')
    return '\n'.join(linhas)


def montar_menu_motorista(nome_motorista):
    return (
        f'Olá {nome_motorista}! O que você gostaria de fazer?\n\n'
        '1- Informar chegada ao local de entrega\n'
        '2- Informar conclusão de entrega\n'
        '3- Informar falha na entrega\n\n'
        'Digite o número correspondente.'
    )


# ── PARTE A: fluxo completo do motorista (menu -> escolher entrega -> confirmar) ──
_ACOES = {'1': 'chegada', '2': 'conclusao', '3': 'falha'}
_ACAO_VERBO = {'chegada': 'a chegada', 'conclusao': 'a conclusão', 'falha': 'a falha'}


def _pedidos_da_entrega(entrega_id):
    idx = {r['id']: r for r in _pedidos_ativos_hoje()}
    e = _entregas_idx().get(entrega_id)
    if not e:
        return []
    return [idx[pid] for pid in e.get('pedidoIds', []) if pid in idx]


def _pedidos_pendentes_da_acao(entrega_id, acao):
    pedidos = _pedidos_da_entrega(entrega_id)
    if acao == 'chegada':
        return [r for r in pedidos if not r.get('chegouLocal')]
    if acao == 'conclusao':
        return [r for r in pedidos if r.get('chegouLocal')]  # só quem já chegou pode concluir
    if acao == 'falha':
        return list(pedidos)  # qualquer pedido ainda em rota pode ser marcado como falha
    return []


def _rotulo_pedido(r):
    return f"{r.get('clienteNome') or r.get('cliente') or ''} / {r.get('filial', '')}"


def montar_lista_escolha(pedidos, acao):
    linhas = [f'Pra qual entrega você quer informar {_ACAO_VERBO[acao]}?', '']
    for i, r in enumerate(pedidos, start=1):
        linhas.append(f'{i}- {_rotulo_pedido(r)}')
    linhas.append('')
    linhas.append('Digite o número correspondente.')
    return '\n'.join(linhas)


def montar_confirmacao(acao, r):
    return (f'Confirma {_ACAO_VERBO[acao]} da entrega pra {_rotulo_pedido(r)}?\n\n'
            'Responda SIM ou NÃO.')


def _aplicar_acao(pedido_id, acao):
    """Aplica de fato a ação escolhida no romaneio, e devolve o pedido
    atualizado (ou None se não achou)."""
    from storage import salvar_romaneio, atualizar_status_romaneio, registrar_desfecho_entrega
    idx = {r['id']: r for r in listar_romaneios()}
    r = idx.get(pedido_id)
    if not r:
        return None

    if acao == 'chegada':
        r['chegouLocal'] = True
        r['chegouLocalEm'] = datetime.datetime.utcnow().isoformat()
        salvar_romaneio(pedido_id, r)
        return r

    novo_status = 'entregue' if acao == 'conclusao' else 'falhou'
    if novo_status == 'falhou':
        atualizar_status_romaneio(pedido_id, 'pendente', falha=True)
    else:
        atualizar_status_romaneio(pedido_id, novo_status)

    entrega_id = r.get('entregaId')
    if entrega_id:
        snap = {'cliente': r.get('clienteNome') or r.get('cliente') or '',
                'filial': r.get('filial', ''), 'kg': r.get('kgPlanejados', 0)}
        registrar_desfecho_entrega(entrega_id, pedido_id, novo_status, snap)

    _limpar_cache_contexto()  # acabamos de escrever -> invalida o cache pra próxima leitura
    idx2 = {rr['id']: rr for rr in listar_romaneios()}
    return idx2.get(pedido_id, r)


def _notificar_cascata(r, acao):
    """Avisa vendedor + encarregados do pedido quando o motorista confirma
    chegada/conclusão/falha (cascata do fim do rascunho)."""
    if acao == 'chegada':
        texto = f'Atualização: a entrega chegou no local — {_rotulo_pedido(r)}.'
    elif acao == 'conclusao':
        texto = f'Atualização: entrega CONCLUÍDA — {_rotulo_pedido(r)}. Obrigado!'
    else:
        texto = f'Atenção: FALHA registrada na entrega — {_rotulo_pedido(r)}. Favor verificar.'

    alvos = [(r.get('vendedor', ''), r.get('telefoneVendedor', ''))] + \
            [(e.get('nome', ''), e.get('telefone', '')) for e in (r.get('encarregados') or [])]
    for _, tel in alvos:
        if tel:
            _enviar_whatsapp(tel, texto)


# ── sessão de conversa por telefone (Cloud Storage, reaproveitando o
# mecanismo de perfil — mesma chave reservada usada em avulso.py/motoristas.py) ──
def _chave_sessao(telefone):
    return f'_sessao_wa_{_normaliza_telefone(telefone)}'


def _carregar_sessao(telefone):
    from storage import perfil_existe, carregar_perfil_bytes
    chave = _chave_sessao(telefone)
    if not perfil_existe(chave):
        return {}
    try:
        import json
        return json.loads(carregar_perfil_bytes(chave).decode('utf-8'))
    except Exception:
        return {}


def _salvar_sessao(telefone, dados):
    from storage import salvar_perfil
    import json
    dados = {**dados, 'atualizadoEm': datetime.datetime.utcnow().isoformat()}
    salvar_perfil(_chave_sessao(telefone), json.dumps(dados).encode('utf-8'), 'sessao.json')


def _limpar_sessao(telefone):
    _salvar_sessao(telefone, {'estado': None})


def _processar_motorista(telefone, nome, entrega_id, texto):
    sessao = _carregar_sessao(telefone)
    estado = sessao.get('estado')
    entrada = (texto or '').strip()

    # sem sessão / conversa nova -> mostra o menu
    if not estado:
        _salvar_sessao(telefone, {'estado': 'aguardando_opcao', 'entregaId': entrega_id})
        return montar_menu_motorista(nome)

    if estado == 'aguardando_opcao':
        acao = _ACOES.get(entrada)
        if not acao:
            return 'Não entendi. ' + montar_menu_motorista(nome)
        pendentes = _pedidos_pendentes_da_acao(entrega_id, acao)
        if not pendentes:
            _limpar_sessao(telefone)
            return 'Não há nenhuma entrega pendente com essa ação agora.'
        if len(pendentes) == 1:
            _salvar_sessao(telefone, {'estado': 'aguardando_confirmacao', 'entregaId': entrega_id,
                                      'acao': acao, 'pedidoId': pendentes[0]['id']})
            return montar_confirmacao(acao, pendentes[0])
        _salvar_sessao(telefone, {'estado': 'aguardando_entrega', 'entregaId': entrega_id, 'acao': acao})
        return montar_lista_escolha(pendentes, acao)

    if estado == 'aguardando_entrega':
        acao = sessao.get('acao')
        pendentes = _pedidos_pendentes_da_acao(entrega_id, acao)
        try:
            escolhido = pendentes[int(entrada) - 1]
        except (ValueError, IndexError):
            return 'Não entendi. ' + montar_lista_escolha(pendentes, acao)
        _salvar_sessao(telefone, {'estado': 'aguardando_confirmacao', 'entregaId': entrega_id,
                                  'acao': acao, 'pedidoId': escolhido['id']})
        return montar_confirmacao(acao, escolhido)

    if estado == 'aguardando_confirmacao':
        resp = entrada.upper()
        acao = sessao.get('acao')
        pedido_id = sessao.get('pedidoId')
        if resp == 'SIM':
            r = _aplicar_acao(pedido_id, acao)
            _limpar_sessao(telefone)
            if not r:
                return 'Não encontrei mais esse pedido — pode ter sido alterado. Digite algo pra ver o menu de novo.'
            _notificar_cascata(r, acao)
            return f'Confirmado! {_ACAO_VERBO[acao].capitalize()} foi registrada. Obrigado!'
        if resp in ('NAO', 'NÃO', 'N'):
            _limpar_sessao(telefone)
            return 'Ok, cancelado. Digite algo pra ver o menu de novo.'
        idx = {r['id']: r for r in _pedidos_da_entrega(entrega_id)}
        r = idx.get(pedido_id)
        return 'Não entendi. ' + (montar_confirmacao(acao, r) if r else 'Responda SIM ou NÃO.')

    # estado desconhecido -> reseta
    _limpar_sessao(telefone)
    return montar_menu_motorista(nome)


# ── mensagens de DESPACHO (disparadas quando a entrega vira 'em_rota') ──
def _linha_pessoa(rotulo, nome, telefone):
    """Uma linha 'Rótulo: nome - telefone', ou None se telefone vazio —
    princípio usado em toda mensagem cruzada (vendedor/encarregado/
    motorista): sem telefone, nem o nome aparece, fica limpo."""
    if not telefone:
        return None
    return f'{rotulo}: {nome or "—"} - {telefone}'


# espaço de largura zero na linha "vazia" — o Botconversa/WhatsApp colapsa
# \n\n puro (testado, confirmado); com o \u200b a linha deixa de ser
# vazia de verdade e a quebra sobrevive, mantendo aparência de linha em branco.
_QUEBRA = '\n\u200b\n'


def montar_mensagem_pedido_em_rota(nome, cliente, filial, papel,
                                   vendedor_nome='', vendedor_telefone='',
                                   encarregados=None, motorista_nome='', motorista_telefone=''):
    """Mensagem #1 do rascunho — despachada pra vendedor, cada
    encarregado, e (junto com a lista) o motorista. 'papel' é
    'vendedor'|'encarregado'|'motorista' — quem está RECEBENDO, pra
    saber quais dos outros 2 mostrar (nunca mostra info de si mesmo).
    Cada linha cruzada só aparece se tiver telefone (ver _linha_pessoa) —
    inclui o caso 'motorista indefinido': sem telefoneMotorista, a linha
    de motorista simplesmente não aparece pra ninguém."""
    linhas_extra = []
    if papel != 'vendedor':
        l = _linha_pessoa('Vendedor', vendedor_nome, vendedor_telefone)
        if l:
            linhas_extra.append(l)
    if papel != 'encarregado':
        for e in (encarregados or []):
            l = _linha_pessoa('Encarregado', e.get('nome'), e.get('telefone'))
            if l:
                linhas_extra.append(l)
    if papel != 'motorista':
        l = _linha_pessoa('Motorista', motorista_nome, motorista_telefone)
        if l:
            linhas_extra.append(l)

    bloco_extra = (_QUEBRA + '\n'.join(linhas_extra)) if linhas_extra else ''
    return (
        f'Olá {nome}, temos uma entrega da Pata Negra em rota para o cliente:{_QUEBRA}'
        f'{cliente}\n{filial}'
        f'{bloco_extra}{_QUEBRA}'
        f'Para atualizações desta rota, envie a palavra STATUS para este número.{_QUEBRA}'
        'Agradecemos toda ajuda no desembarque.'
    )


def montar_mensagem_lista_motorista(nome_motorista, pedidos):
    """Mensagem #2 do rascunho — lista do dia pro motorista, com
    vendedor + encarregado(s) de cada parada (omitidos quando sem
    telefone — mesmo princípio de _linha_pessoa)."""
    blocos = []
    for r in pedidos:
        cliente = r.get('clienteNome') or r.get('cliente') or ''
        filial = r.get('filial', '')
        linhas_parada = [f'{cliente} / {filial}']
        l = _linha_pessoa('Vendedor', r.get('vendedor'), r.get('telefoneVendedor'))
        if l:
            linhas_parada.append(l)
        for e in (r.get('encarregados') or []):
            l = _linha_pessoa('Encarregado', e.get('nome'), e.get('telefone'))
            if l:
                linhas_parada.append(l)
        blocos.append('\n'.join(linhas_parada))
    corpo = _QUEBRA.join(blocos)
    return f'Olá {nome_motorista}, temos as seguintes entregas hoje:{_QUEBRA}{corpo}{_QUEBRA}Bom trabalho!'


def notificar_despacho_entrega(entrega):
    """Chamada quando uma entrega passa pra fase 'em_rota' (o botão de
    despacho). Dispara a mensagem #1 (vendedor + encarregados) pra cada
    pedido, e a #2 (lista do dia) pro motorista — UMA vez cada.

    FIX (23/08/2026): antes, 'Motorista indefinido' pulava TUDO (nem
    vendedor/encarregado recebia nada) — errado, a mensagem #1 não
    menciona o motorista em lugar nenhum, não tem motivo pra depender
    dele. Agora só a mensagem #2 (lista do motorista) é pulada quando
    não há motorista definido; vendedor/encarregado sempre recebem a
    #1 normalmente."""
    idx = {r['id']: r for r in _pedidos_ativos_hoje()}
    pedidos = [idx[pid] for pid in entrega.get('pedidoIds', []) if pid in idx]
    if not pedidos:
        return

    for r in pedidos:
        mot_nome, mot_tel = _motorista_da_entrega(r, _entregas_idx())
        alvos = [(r.get('vendedor', ''), r.get('telefoneVendedor', ''), 'vendedor')] + \
                [(e.get('nome', ''), e.get('telefone', ''), 'encarregado') for e in (r.get('encarregados') or [])]
        for nome, tel, papel in alvos:
            if tel:
                _enviar_whatsapp(tel, montar_mensagem_pedido_em_rota(
                    nome, r.get('clienteNome') or r.get('cliente') or '', r.get('filial', ''), papel,
                    vendedor_nome=r.get('vendedor', ''), vendedor_telefone=r.get('telefoneVendedor', ''),
                    encarregados=r.get('encarregados'),
                    motorista_nome=mot_nome, motorista_telefone=mot_tel))

    telefone_motorista = (entrega.get('telefoneMotorista') or '').strip()
    if telefone_motorista:
        _enviar_whatsapp(telefone_motorista,
                         montar_mensagem_lista_motorista(entrega.get('motorista', ''), pedidos))


def _enviar_whatsapp(telefone, texto):
    """Envia mensagem PROATIVA via API do Botconversa (endpoint confirmado
    em 22/08/2026, via Swagger — backend.botconversa.com.br/swagger/).

    Fluxo em 2 passos, já que só temos o telefone:
      1. GET .../subscriber/get_by_phone/{telefone}/ -> devolve o
         subscriber_id (campo 'id' da resposta).
      2. POST .../subscriber/{subscriber_id}/send_message/ com
         {"type": "text", "value": texto}.

    Autenticação: header 'API-KEY' (a chave de Configurações →
    Integrações → seção "API" — NÃO a do Zapier nem a do RD Station,
    que são chaves separadas dentro da mesma tela)."""
    import os
    import requests

    api_key = os.environ.get('BOTCONVERSA_API_KEY', '')
    if not api_key:
        print(f'[WARN] BOTCONVERSA_API_KEY não configurada — mensagem NÃO enviada pra {telefone}')
        return

    base = 'https://backend.botconversa.com.br/api/v1/webhook'
    headers = {'API-KEY': api_key}
    tel_limpo = _normaliza_telefone(telefone)  # só dígitos, com DDI 55 (aceito com ou sem '+')

    try:
        resp_id = requests.get(f'{base}/subscriber/get_by_phone/{tel_limpo}/',
                               headers=headers, timeout=4)
        if resp_id.status_code == 404:
            # comum: telefone nunca mandou mensagem pro número do Botconversa
            # antes, então não existe como 'subscriber' ainda -- não é erro
            # de configuração, é esperado até a pessoa interagir 1x.
            print(f'[INFO] {telefone} ainda não é subscriber no Botconversa '
                  f'(precisa mandar mensagem pro número pelo menos 1 vez antes)')
            return
        resp_id.raise_for_status()
        subscriber_id = resp_id.json().get('id')
        if not subscriber_id:
            print(f'[WARN] Botconversa não achou subscriber pro telefone {telefone}')
            return

        resp_send = requests.post(f'{base}/subscriber/{subscriber_id}/send_message/',
                                  headers=headers,
                                  json={'type': 'text', 'value': texto},
                                  timeout=4)
        resp_send.raise_for_status()
    except Exception as e:
        print(f'[WARN] falha ao enviar WhatsApp pra {telefone}: {e}')


def processar_mensagem_entrada(telefone, texto=''):
    """Ponto de entrada único chamado pelo /whatsapp/webhook. 'texto' é o
    que a pessoa acabou de digitar (agora disponível, desde que o ciclo
    Integração -> Conteúdo/SALVAR -> Integração foi fechado no BC)."""
    _limpar_cache_contexto()  # request novo -> começa com dado fresco
    papel, dado = identificar_papel(telefone)
    alvo = _normaliza_telefone(telefone)

    if papel == 'vendedor':
        return montar_mensagem_vendedor(dado['nome'], alvo)
    if papel == 'encarregado':
        return montar_mensagem_encarregado(dado['nome'], alvo)
    if papel == 'motorista':
        return _processar_motorista(telefone, dado['nome'], dado['entregaId'], texto)

    return ('Não localizei seu número em nenhuma entrega ativa hoje. '
            'Se você é vendedor, encarregado ou motorista da Pata Negra e recebeu essa '
            'mensagem por engano, fala com a gente direto: '
            'https://wa.me/5521990111992')
