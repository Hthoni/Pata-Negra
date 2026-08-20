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
saía com quantidade zerada/em branco. Causa: a extração de Valor Unit./
Valor Item usava ÍNDICES FIXOS diferentes pra CX (nums[3]/nums[5]) e pra
KG (nums[2]/nums[4]) — presumindo que a coluna KG tivesse uma coluna a
menos. Isso não é mais verdade no layout atual do TOTVS: a estrutura de
colunas é a MESMA nos dois casos (Estoq., Qtde, Valor Unit., Valor Emb.,
Valor Item, Valor Bruto, ...), só que pra item em KG "Valor Unit." e
"Valor Emb." saem IGUAIS (fator de embalagem = 1), então os índices
antigos pegavam a coluna errada (Qtde como se fosse preço, e Valor Unit.
como se fosse total).

Corrigido pra não depender de índice fixo nenhum: "Valor Item" e "Valor
Bruto" sempre saem com o MESMO valor, lado a lado, quando não há IPI/
desconto (sempre o caso aqui) — acha esse par pela repetição, pega o
ÚLTIMO par assim (pra não cair no par "falso" Valor Unit./Valor Emb. que
os itens em KG têm mais cedo na linha), e o preço é o valor 2 posições
antes desse par. Funciona igual pra CX e pra KG, sem precisar de branch
separado — mais simples e mais resistente a colunas que o TOTVS mude de
novo no futuro.
"""

import io
import re
import pdfplumber
from perfil import processar_item

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


def _parse_item(ln, prox):
    parts = ln.split()
    if not parts or not re.match(r'^\d{4,6}$', parts[0]):
        return None
    # embalagem = primeiro CX/KG cujo próximo token é numérico
    emb_j = None
    for j, p in enumerate(parts):
        if p in ('CX', 'KG') and j + 1 < len(parts) and re.match(r'^[\d.,]+$', parts[j + 1]):
            emb_j = j
            break
    if emb_j is None:
        return None
    # nome = tokens entre os dígitos iniciais (cod/seq) e a embalagem
    k = 0
    while k < len(parts) and re.match(r'^\d{3,6}$', parts[k]):
        k += 1
    nome = ' '.join(parts[k:emb_j])
    nums = parts[emb_j + 1:]
    preco, total = _preco_e_total(nums)
    if not preco:
        return None
    # sufixo do nome na próxima linha (não puxa '- REF:', 'EANs', códigos)
    if prox and len(prox) < 30 and not prox.startswith(('EANs', 'TOTAIS', '- REF')):
        suf = re.sub(r'\bKG\b', '', prox).strip().rstrip('-').strip()
        if suf and not suf.upper().startswith('REF'):
            nome = (nome + ' ' + suf).strip()
    kg = round(total / preco, 3) if preco else 0.0
    return {'cod': parts[0], 'nome': nome, 'kg': kg, 'preco': preco, 'total': total}


def parse(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pags = len(pdf.pages)
        for pi in range(0, n_pags, 2):
            txt1 = pdf.pages[pi].extract_text() or ''
            txt2 = pdf.pages[pi + 1].extract_text() if pi + 1 < n_pags else ''
            lines = txt1.split('\n')
            txt_all = txt1 + '\n' + txt2

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
            for i, ln in enumerate(lines):
                prox = lines[i + 1].strip() if i + 1 < len(lines) else ''
                d = _parse_item(ln.strip(), prox)
                if not d:
                    continue
                # kg já é físico -> passa como KG p/ processar_item não multiplicar
                it = processar_item(d['cod'], d['nome'], 'KG', 1, d['kg'], d['preco'], d['total'], produtos)
                it['empresa'] = empresa
                itens.append(it)

            if itens:
                filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                                'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                                'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
