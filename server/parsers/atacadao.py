"""
Parser Atacadão — formato próprio com tabela em pipes (|).
Cada produto tem linha principal (RF.) + linha de referência (KG/GR/CXA).
Linhas GR (gramas) são duplicatas e devem ser ignoradas.
"""
import io
import re
import pdfplumber
from perfil import processar_item
CNPJ_INDUSTRIA = '10.171.633'
CNPJ_RE = r'\d{2}\.?\d{3}\.?\d{3}\s*/\s*\d{4}\s*-\s*\d{2}'  # pontos opcionais (Atacadão não pontua) + espaço opcional ao redor de / e - (artefato do pdfplumber)
def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)
    lines = txt.split('\n')
    def fm(pat):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''
    pedidoNum = fm(r'Numero:\s*(\d+)')
    cnpj = fm(rf'Local de Entrega:\s*({CNPJ_RE})')
    empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj.replace('.', '').replace('-', '').replace(' ', '') else 2
    cond_m = re.search(r'\|\s+(\d+)\s*\+?\s*\n.*?da data de recebimento', txt, re.S)
    condPgto = cond_m.group(1) + ' dias' if cond_m else fm(r'Condicoes de Pagto.*?\n\|\s*(\d+)')
    if condPgto and 'dias' not in condPgto:
        condPgto += ' dias'
    dataEntrega = fm(r'RF\.[^|]+\s+(\d{2}/\d{2})\s+\d')
    if dataEntrega:
        p = dataEntrega.split('/')
        dataEntrega = f'{p[0]}/{p[1]}/2026' if len(p) == 2 else dataEntrega
    endereco = ''
    for i, ln in enumerate(lines):
        if 'Local de Entrega:' in ln:
            for j in range(i + 1, min(i + 4, len(lines))):
                m2 = re.search(r'\|\s+([A-Z][^|]+?)\s+\|', lines[j])
                if m2 and len(m2.group(1).strip()) > 5:
                    endereco = m2.group(1).strip()
                    break
            break
    filial_num = re.search(r'/\s*(\d{4})\s*-', cnpj)
    filial = f'ATACADÃO FILIAL {int(filial_num.group(1))}' if filial_num else 'ATACADÃO'
    reItem = re.compile(r'\|\s+(RF\.\w[^|]+?)\s+(\d{2}/\d{2})\s+(\d+)\s+([\d,.]+)\s+0,.*?([\d,.]+)\s+S\s+\|')
    reRef = re.compile(r'\|\s+(\d{8}/\d+)\s+(KG|GR|CXA)\s+(\S+)')
    itens = []
    pending = None
    for ln in lines:
        mm2 = reItem.match(ln)
        if mm2:
            pending = {'nome': mm2.group(1).strip(),
                       'qtd': float(mm2.group(3).replace('.', '').replace(',', '.')),
                       'preco': float(mm2.group(4).replace('.', '').replace(',', '.'))}
            continue
        m2 = reRef.search(ln)
        if m2 and pending:
            if m2.group(2) == 'GR':
                pending = None
                continue
            emb_tipo = 'CX' if m2.group(2) == 'CXA' else m2.group(2)
            qtde_emb = int(re.search(r'X(\d+)X', m2.group(3) or '1X20X1').group(1)) if re.search(r'X(\d+)X', m2.group(3) or '') else 20
            total = round(pending['qtd'] * pending['preco'], 2)
            it = processar_item(m2.group(1), pending['nome'], emb_tipo,
                                 qtde_emb, pending['qtd'], pending['preco'], total, produtos)
            itens.append(it)
            pending = None
    if not itens:
        return []
    return [{'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj, 'endereco': endereco,
              'dataPedido': '', 'dataEntrega': dataEntrega, 'condPgto': condPgto,
              'empresa': empresa, 'itens': itens}]
