"""
Parser DOM Atacarejo — formato TOTVS, multi-filial.
PDF tem 2 páginas por filial: ímpar = dados do pedido, par = datas/assinaturas.

FIX (11/08/2026): erro de "produto não cadastrado" em quase todos os itens.
Causa: a descrição do produto nesse layout SEMPRE continua nas linhas
seguintes à linha do código (ex.: "BACON BARRIGA KG 1 20,00..." na 1ª linha,
"PATA NEGRA KG" / "DEF PORC - REF:" / "22 PATA NEGRA" nas linhas de baixo) —
a regex antiga só operava dentro de uma linha só (via re.M/^...$), então
capturava apenas o fragmento inicial do nome ("BACON BARRIGA"), nunca o nome
completo cadastrado no perfil ("BACON BARRIGA PATA NEGRA KG DEF PORC").

Corrigido reconstruindo o nome linha a linha: depois de casar a linha do
item, junta as linhas seguintes (até a próxima linha de item ou "TOTAIS")
e corta tudo a partir de "REF:" (número de referência interno do DOM, que
vem sempre colado ao fim do nome e não é parte da descrição do produto).
"""

__cliente_nome__ = "DOM Atacarejo"

import io
import re
import pdfplumber
from perfil import processar_item

_RE_ITEM_INICIO = re.compile(
    r'^(\d{5,6})\s+\d+\s+(.+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)'
)
_RE_PROX_ITEM = re.compile(r'^\d{5,6}\s+\d+\s')


def _reconstruir_nome(linhas, idx_item, primeira_parte):
    """Junta a linha do item às linhas de continuação seguintes (até o
    próximo item ou 'TOTAIS'), e corta a partir de 'REF:' (referência
    interna do DOM, não é parte do nome do produto)."""
    partes = [primeira_parte]
    i = idx_item + 1
    while i < len(linhas):
        ln = linhas[i].strip()
        if not ln or ln.startswith('TOTAIS') or _RE_PROX_ITEM.match(ln):
            break
        partes.append(ln)
        i += 1
    nome = ' '.join(partes)
    nome = re.split(r'\s*-?\s*REF:', nome, maxsplit=1)[0]
    return re.sub(r'\s+', ' ', nome).strip()


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

            pedidoNum = fm(r'(\d{5,7}/[CL])')
            filial = ''
            for ln in lines:
                if 'DOM ATACAREJO SA' in ln:
                    filial = re.sub(r'\s+R\..*$', '', re.sub(r'.*DOM ATACAREJO SA\s+', '', ln)).strip()
                    break
            cnpj = ''
            for ln in lines:
                found = re.findall(r'\d{2}\.\d{3}\.\d{3}\s*/\s*\d{4}\s*-\s*\d{2}', ln)
                if len(found) >= 2:
                    cnpj = found[1]
                    break
                elif len(found) == 1:
                    cnpj = found[0]
                    break
            endereco = ''
            for i, ln in enumerate(lines):
                if 'ENDEREÇO PARA ENTREGA' in ln:
                    for j in range(i + 1, min(i + 5, len(lines))):
                        if lines[j].startswith('Endereço'):
                            endereco = re.split(r'\s{2,}', lines[j].replace('Endereço', '').strip())[0].strip()
                            break
                    break
            dataPedido = fm(r'Data da emiss[aã]o\s+([\d/]+)')
            dataEntrega = fm(r'Previs[aã]o de entrega\s+([\d/]+)')
            condPgto = fm(r'Prazo para pagamento\s+(\d+)')
            if condPgto:
                condPgto += ' dias'

            itens = []
            for idx, ln in enumerate(lines):
                m = _RE_ITEM_INICIO.match(ln.strip())
                if not m:
                    continue
                nome = _reconstruir_nome(lines, idx, m.group(2))
                qtde_ped = float(m.group(5).replace('.', '').replace(',', '.'))
                preco = float(m.group(6).replace('.', '').replace(',', '.'))
                total = float(m.group(7).replace('.', '').replace(',', '.'))
                it = processar_item(int(m.group(1)), nome, m.group(3),
                                     int(m.group(4)), qtde_ped, preco, total, produtos)
                itens.append(it)

            if filial and itens:
                filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                                 'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                                 'condPgto': condPgto, 'empresa': 2, 'itens': itens})
    return filiais
