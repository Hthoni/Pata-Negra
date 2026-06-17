"""
Parser Assaí — formato Consinco/TOTVS, uma página por filial.
Código do produto vem colado ao nome (ex: '1156510BACON LOMBO...').
"""
import re
import pdfplumber
from perfil import processar_item

CNPJ_INDUSTRIA = '10.171.633'


def parse(pdf_bytes, produtos):
    import io
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ''

            def fm(pat):
                m = re.search(pat, txt, re.I)
                return m.group(1).strip() if m else ''

            pedidoNum = fm(r'PEDIDO DE COMPRAS\s+(\S+)')
            cnpj_m = re.search(r'CNPJ\s+([\d./\-]+)\s+Cidade.*?CNPJ\s+([\d./\-]+)', txt, re.S)
            cnpj_forn = cnpj_m.group(1) if cnpj_m else ''
            cnpj_loja = cnpj_m.group(2) if cnpj_m else ''
            empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj_forn.replace('.', '') else 2
            filial_m = re.search(r'R\. Social SENDAS.*?LJ\d+\s+\d+\s+(.+?)$', txt, re.M)
            filial = filial_m.group(1).strip() if filial_m else 'ASSAÍ'
            end_m = re.search(r'ENDEREÇO PARA ENTREGA.*?Endereço\s+(.+?)\s+Endereço', txt, re.S)
            endereco = end_m.group(1).strip() if end_m else ''
            dataPedido = fm(r'Data da emiss[aã]o\s+([\d/]+)')
            dataEntrega = fm(r'Previs[aã]o de entrega\s+([\d/]+)')
            cond_m2 = re.search(r'pagamento\s+(\d+)\s*\(', txt)
            condPgto = cond_m2.group(1) + ' dias' if cond_m2 else ''

            reItem = re.compile(
                r'^(\d{7})([A-Z][^\n]+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)',
                re.M)
            itens = []
            for m in reItem.finditer(txt):
                nome_raw = re.sub(r'\s+(FRAC\s*KG|KG)\s*$', '', re.sub(r'\s+', ' ', m.group(2))).strip()
                qtde_ped = float(m.group(5).replace('.', '').replace(',', '.'))
                preco = float(m.group(6).replace('.', '').replace(',', '.'))
                total = float(m.group(7).replace('.', '').replace(',', '.'))
                it = processar_item(m.group(1), nome_raw, m.group(3),
                                     int(m.group(4)), qtde_ped, preco, total, produtos)
                it['empresa'] = empresa
                itens.append(it)

            if itens:
                filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj_loja,
                                 'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                                 'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
