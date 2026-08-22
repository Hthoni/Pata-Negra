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


def _pedidos_ativos_hoje():
    """Pedidos de entregas DESPACHADAS hoje (fase='em_rota', despachadaEm
    de hoje) e que ainda estão com status='em_rota' (não entregues/
    falhados). 'Hoje' é o dia do despacho, não o dia em que o pedido foi
    originalmente processado."""
    hoje_iso = datetime.date.today().isoformat()
    entregas_hoje = [
        e for e in listar_entregas()
        if e.get('fase') == 'em_rota' and (e.get('despachadaEm') or '').startswith(hoje_iso)
    ]
    ids_hoje = set()
    for e in entregas_hoje:
        ids_hoje.update(e.get('pedidoIds', []))
    idx = {r['id']: r for r in listar_romaneios()}
    return [idx[pid] for pid in ids_hoje if pid in idx and idx[pid].get('status') == 'em_rota']


def _entregas_idx():
    return {e['id']: e for e in listar_entregas()}


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

    for e in listar_entregas():
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
    pedidos = [r for r in _pedidos_ativos_hoje()
               if any(_normaliza_telefone(e.get('telefone', '')) == telefone_alvo
                      for e in (r.get('encarregados') or []))]
    if not pedidos:
        return f'Olá {nome_encarregado}, não encontrei nenhuma entrega em rota hoje pra sua filial.'

    idx_ent = _entregas_idx()
    linhas = [f'Olá {nome_encarregado}, a entrega abaixo tem o seguinte status:', '']
    for r in pedidos:
        cliente = r.get('clienteNome') or r.get('cliente') or ''
        filial = r.get('filial', '')
        status = _status_legivel(r)
        mot_nome, mot_tel = _motorista_da_entrega(r, idx_ent)
        linhas.append(f'{cliente} / {filial} / {status}')
        linhas.append('')
        linhas.append(f'Motorista {mot_nome} telefone - {mot_tel}')
        linhas.append(f'Vendedor {r.get("vendedor", "—")} telefone {r.get("telefoneVendedor", "—")}')
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


def processar_mensagem_entrada(telefone):
    """Ponto de entrada único chamado pelo /whatsapp/webhook."""
    papel, dado = identificar_papel(telefone)
    alvo = _normaliza_telefone(telefone)

    if papel == 'vendedor':
        return montar_mensagem_vendedor(dado['nome'], alvo)
    if papel == 'encarregado':
        return montar_mensagem_encarregado(dado['nome'], alvo)
    if papel == 'motorista':
        return montar_menu_motorista(dado['nome'])

    return ('Não localizei seu número em nenhuma entrega ativa hoje. '
            'Se você é vendedor, encarregado ou motorista da Pata Negra e recebeu essa '
            'mensagem por engano, entre em contato com o atendimento.')
