"""
Parser Superprix — formato texto corrido próprio (não-TOTVS),
uma linha por item.
"""
import io
import re
import pdfplumber
from perfil import processar_item

CNPJ_INDUSTRIA = '10.171.633'


def parse(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    def fm(pat):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'Pedido:\s*(\d+-\d+)')
    dataPedido = fm(r'Emis:\s*([\d/]+)')
    dataEntrega = fm(r'Entrega\s*:\s*([\d/]+)')
    condPgto = fm(r'PGTO\s+(\d+)\s+DIAS')
    if condPgto:
        condPgto += ' dias'
    cnpj = fm(r'CNPJ\s*:\s*([\d./\-]+)\s+INSC')
    cnpj_forn = fm(r'CNPJ\s*:\s*[\d./\-]+.*?CNPJ\s*:\s*([\d./\-]+)')
    empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj_forn.replace('.', '').replace('-', '') else 2

    filial_m = re.search(r'Bairro\s*:\s*([A-Z][A-Z\s]+?)\s+Cidade', txt)
    filial = filial_m.group(1).strip() if filial_m else 'SUPERPRIX'
    end_m = re.search(r'Endereço\s*:\s*([^\n]+)', txt)
    endereco = end_m.group(1).strip() if end_m else ''

    # itens: Nome  Emb  Cod  Qtde  Preco  ...  Total
    reItem = re.compile(
        r'^([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][^\n]+?)\s+(KG/\d+|CX/\d+)\s+\S+\s+(\d+)\s+([\d,.]+)\s+[\d,.]+\s+([\d,.]+)',
        re.M
    )
    itens = []
    for m in reItem.finditer(txt):
        nome = m.group(1).strip()
        emb_raw = m.group(2)
        emb_tipo = 'CX' if emb_raw.startswith('CX') else 'KG'
        qtde_emb = int(emb_raw.split('/')[1]) if '/' in emb_raw else 1
        qtde_ped = float(m.group(3).replace('.', '').replace(',', '.'))
        preco = float(m.group(4).replace('.', '').replace(',', '.'))
        total = float(m.group(5).replace('.', '').replace(',', '.'))
        it = processar_item('', nome, emb_tipo, qtde_emb, qtde_ped, preco, total, produtos)
        itens.append(it)

    if itens:
        filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                         'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                         'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
