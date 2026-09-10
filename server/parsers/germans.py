"""
Parser Germans (Supermercados Campeão) — formato TOTVS.
Pedido multi-página (2 págs por pedido). Itens numa linha, nome podendo
continuar na linha seguinte; algumas linhas sem Cód. Forn (só a Seq).

Robustez:
 - embalagem = primeiro token CX/KG SEGUIDO de número (evita o 'KG' que faz
   parte do nome do produto, ex.: 'BACON PORC KG');
 - Cód. Forn opcional (linha pode começar pela Seq, ex.: INGRED FEIJOADA);
 - kg físico = Valor Item ÷ Valor Unit (preço é por kg) — robusto a colunas
   extras que o TOTVS às vezes insere.

FIX (20/08/2026): item MINI COSTELA (o único faturado em KG, não em CX)
saía com quantidade zerada/em branco. Ver histórico do módulo — corrigido
extraindo preço/total pelo padrão estrutural (par Valor Item == Valor
Bruto), não por índice fixo de coluna.

FIX (10/09/2026): nome de produto pode quebrar em ATÉ 2 linhas de
continuação, não só 1 (ex.: "CHISPE" + "SALGADO PORC" + "KG - REF: 53" —
o "REF: 53" é resquício de referência a ignorar, o "KG" faz parte do nome
de verdade). O "REF:" às vezes vem grudado na linha ("KG - REF: 53"),
às vezes em linha própria ("REF: 83", depois de "DEF PORC KG -").

IMPORTANTE: isso NÃO relaxa o matching -- continua sendo exato, sempre.
O que essa lógica faz é só tentar RECONSTRUIR o texto (já que o PDF quebra
o nome de formas variáveis, imprevisíveis, sem relação com o pedido em
si) de algumas formas plausíveis (linha crua, sem "KG", sem sufixo
"- REF: N", e sem os dois; até 2 linhas seguintes) -- só aceita uma
reconstrução se ela bater CARACTERE POR CARACTERE com um nome já
cadastrado no Perfil. Se nenhuma reconstrução bater, erro claro, igual
sempre foi. Nunca aproxima produto diferente, nunca "adivinha" -- só dá
mais chances de achar a forma certa de juntar um nome que o PDF quebrou
em pedaços.
"""

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

CNPJ_DISTRIBUIDORA = '56.423.719'
CNPJ_INDUSTRIA = '10.171.633'


def _num(s):
    return float(str(s).replace('.', '').replace(',', '.'))


def _preco_e_total(nums):
    """Acha o preço unitário e o total a partir do padrão estrutural da
    linha (ver FIX no docstring do módulo), em vez de índice fixo de
    coluna. 'nums' é a lista de tokens numéricos após a embalagem."""
    floats = []
    for n in nums:
        try:
            floats.append(_num(n))
        except ValueError:
            floats.append(None)
    idx = None
    for i in range(len(floats) - 1):
        a, b = floats[i], floats[i + 1]
        if a is not None and b is not None and a > 0 and abs(a - b) < 0.005:
            idx = i  # mantém o ÚLTIMO par igual encontrado
    if idx is None:
        return 0.0, 0.0
    total = floats[idx]
    preco = floats[idx - 2] if idx - 2 >= 0 and floats[idx - 2] is not None else 0.0
    return preco, total


def _candidatos_continuacao(prox):
    """Gera as variações plausíveis de uma linha de continuação de nome:
    crua, sem 'KG', sem sufixo '- REF: N' (às vezes grudado na mesma
    linha, ex.: "KG - REF: 53"), e sem os dois -- sempre removendo um
    hífen solto que sobra no final (ex.: "DEF PORC KG -", quando o
    "REF: N" vem numa linha SEPARADA em vez de grudado). Devolve também
    o "fallback" mais seguro (sem REF/hífen solto) pra usar quando nada
    bate ainda, mas pode precisar de mais uma linha depois."""
    cands = [prox]
    sem_ref = re.sub(r'\s*-\s*REF:.*$', '', prox, flags=re.I).strip().rstrip('-').strip()
    if sem_ref and sem_ref not in cands:
        cands.append(sem_ref)
    sem_kg = re.sub(r'\bKG\b', '', prox).strip().rstrip('-').strip()
    if sem_kg and sem_kg not in cands:
        cands.append(sem_kg)
    sem_kg_ref = re.sub(r'\bKG\b', '', sem_ref).strip().rstrip('-').strip()
    if sem_kg_ref and sem_kg_ref not in cands:
        cands.append(sem_kg_ref)
    return cands, sem_ref


def _tenta_nome_completo(nome_base, linhas, i_prox, produtos, max_linhas_extra=2):
    """Tenta casar nome_base sozinho contra o Perfil; se não bater, vai
    mesclando as linhas seguintes (até max_linhas_extra), testando cada
    combinação -- só usa a versão mesclada quando resulta num MATCH
    EXATO de verdade. Pula/para se achar EANs/TOTAIS/REF ou início de
    outro item."""
    nome = nome_base
    if match_perfil(nome, produtos):
        return nome
    j = i_prox
    extra = 0
    while extra < max_linhas_extra and j < len(linhas):
        prox = linhas[j].strip()
        if not prox or re.match(r'^\d{4,6}\s', prox) or prox.startswith(('EANs', 'TOTAIS', 'REF:')):
            break
        cands, fallback = _candidatos_continuacao(prox)
        for c in cands:
            tentativa = (nome + ' ' + c).strip()
            if match_perfil(tentativa, produtos):
                return tentativa
        nome = (nome + ' ' + fallback).strip() if fallback else nome
        j += 1
        extra += 1
    return nome


def _parse_item(ln):
    """Extrai só o que dá pra tirar da PRÓPRIA linha do item -- nome
    ainda pode estar incompleto (ver _tenta_nome_completo, chamado
    depois, já com acesso ao Perfil pra confirmar a mesclagem).

    FIX (10/09/2026): o "Cod Forn" normalmente tem 4-6 dígitos
    (ex.: "000075"), mas um item real veio com um código de barras (EAN)
    de 14 dígitos no lugar ("17898611040078") -- com o limite de 6
    dígitos, a linha inteira era descartada em SILÊNCIO (sem erro
    nenhum), o item simplesmente sumia. Ampliado pra aceitar até 20
    dígitos, cobrindo tanto o código curto normal quanto um EAN
    completo."""
    parts = ln.split()
    if not parts or not re.match(r'^\d{4,20}$', parts[0]):
        return None
    emb_j = None
    for j, p in enumerate(parts):
        if p in ('CX', 'KG') and j + 1 < len(parts) and re.match(r'^[\d.,]+$', parts[j + 1]):
            emb_j = j
            break
    if emb_j is None:
        return None
    k = 0
    while k < len(parts) and re.match(r'^\d{3,20}$', parts[k]):
        k += 1
    nome = ' '.join(parts[k:emb_j])
    nums = parts[emb_j + 1:]
    preco, total = _preco_e_total(nums)
    if not preco:
        return None
    kg = round(total / preco, 3) if preco else 0.0
    return {'cod': parts[0], 'nome': nome, 'kg': kg, 'preco': preco, 'total': total}


def parse(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        paginas_txt = [(p.extract_text() or '') for p in pdf.pages]

    # Agrupa páginas pelo Nº do pedido repetido no cabeçalho -- pedido que
    # estoura de 2 páginas tem itens escaneados em TODAS as páginas do grupo.
    grupos = []
    pedido_atual = None
    for txt in paginas_txt:
        m = re.search(r'(\d{5,7}/[ML])', txt)
        num = m.group(1) if m else None
        if num and num == pedido_atual:
            grupos[-1].append(txt)
        else:
            grupos.append([txt])
            pedido_atual = num

    for paginas in grupos:
        txt1 = paginas[0]
        txt_all = '\n'.join(paginas)
        lines = txt1.split('\n')          # cabeçalho/filial só na 1ª página
        lines_itens = txt_all.split('\n')  # itens podem estar em qualquer página do grupo

        def fm(pat, txt=txt_all):
            m = re.search(pat, txt, re.I)
            return m.group(1).strip() if m else ''

        pedidoNum = fm(r'(\d{5,7}/[ML])')
        dataPedido = fm(r'Data da emiss[aã]o\s+([\d/]+)')
        dataEntrega = fm(r'Previs[aã]o de entrega\s+([\d/]+)')
        condPgto = fm(r'Prazo para pagamento\s+(\d+)')
        if condPgto:
            condPgto += ' dias'

        cnpj = ''
        for ln in lines:
            found = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', ln)
            if len(found) >= 2:
                cnpj = found[1]; break
            elif len(found) == 1 and CNPJ_INDUSTRIA not in found[0] and CNPJ_DISTRIBUIDORA not in found[0]:
                cnpj = found[0]; break

        cnpj_forn = fm(r'CNPJ\s+([\d./\- ]+?)\s+Inscri')
        empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj_forn.replace('.', '').replace('-', '').replace(' ', '') else 2

        filial_m = re.search(r'COMESTIVEI?\s+(.+?)$', txt1, re.M)
        filial = filial_m.group(1).strip() if filial_m else 'CAMPEAO - CORDOVIL'
        endereco = 'RUA CORDOVIL-1000, PARADA DE LUCAS'

        itens = []
        for i, ln in enumerate(lines_itens):
            d = _parse_item(ln.strip())
            if not d:
                continue
            nome_completo = _tenta_nome_completo(d['nome'], lines_itens, i + 1, produtos)
            # kg já é físico -> passa como KG p/ processar_item não multiplicar
            it = processar_item(d['cod'], nome_completo, 'KG', 1, d['kg'], d['preco'], d['total'], produtos)
            it['empresa'] = empresa
            itens.append(it)

        if itens:
            filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                            'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                            'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
