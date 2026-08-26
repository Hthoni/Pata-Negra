"""
Parser DOM Atacarejo — formato TOTVS, multi-filial.
Páginas por filial VARIAM (1 ou 2) — agrupadas dinamicamente pelo Nº do
pedido repetido no cabeçalho de cada página (ver FIX 26/08/2026 abaixo).

CONVENÇÃO (11/08/2026): o nome do produto usado pro matching é SÓ o
fragmento que aparece na MESMA linha do código/valores (antes da coluna
Emb. KG/CX) — nunca o nome completo reconstruído das linhas de baixo.
Isso evita qualquer lógica de "juntar continuação + cortar REF" (frágil,
gera exceção atrás de exceção conforme cada produto quebra de um jeito
diferente entre as linhas). O perfil precisa cadastrar o produto com esse
MESMO fragmento curto — ver lista de referência no fim do arquivo.
"""

__cliente_nome__ = "DOM Atacarejo"

import io
import re
import pdfplumber
from perfil import processar_item


def parse(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        paginas_txt = [(p.extract_text() or '') for p in pdf.pages]

    # Agrupa páginas pelo Nº do pedido repetido no cabeçalho ("PEDIDO DE
    # COMPRAS NNNNNN/X"). FIX (26/08/2026): o PDF NÃO tem sempre 2 páginas
    # por filial -- pedido curto cabe numa página só (itens+totais+
    # assinatura); pedido longo estoura pra uma 2ª página (repete cabeçalho
    # + assinatura, sem itens novos). O código antigo assumia par fixo
    # (0,1)/(2,3)/... e desalinhava a cada filial de 1 página só, perdendo
    # a filial seguinte inteira (derrubou 15 de 26 filiais num teste real).
    grupos = []
    pedido_atual = None
    for txt in paginas_txt:
        m = re.search(r'PEDIDO DE COMPRAS\s+(\d{5,7}/[A-Z])', txt)
        num = m.group(1) if m else None
        if num and num == pedido_atual:
            grupos[-1].append(txt)
        else:
            grupos.append([txt])
            pedido_atual = num

    for paginas in grupos:
        txt1 = paginas[0]
        txt_all = '\n'.join(paginas)
        lines = txt1.split('\n')

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

        # Nome = só a mesma linha, até a coluna Emb. (KG|CX). Sem
        # reconstrução das linhas de continuação — ver docstring.
        reItem = re.compile(r'^(\d{5,6})\s+\d+\s+(.+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)', re.M)
        itens = []
        for m in reItem.finditer(txt1):
            nome_raw = re.sub(r'\s+', ' ', m.group(2)).strip()
            qtde_ped = float(m.group(5).replace('.', '').replace(',', '.'))
            preco = float(m.group(6).replace('.', '').replace(',', '.'))
            total = float(m.group(7).replace('.', '').replace(',', '.'))
            it = processar_item(int(m.group(1)), nome_raw, m.group(3),
                                 int(m.group(4)), qtde_ped, preco, total, produtos)
            itens.append(it)

        if filial and itens:
            filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                             'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                             'condPgto': condPgto, 'empresa': 2, 'itens': itens})
    return filiais
