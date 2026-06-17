"""
Parser Torre Central — formato TOTVS, nome do produto na linha
ANTERIOR ao código+dados (diferente do padrão DOM).
"""
import io
import re
import pdfplumber
from perfil import processar_item

CNPJ_INDUSTRIA = '10.171.633'


def parse(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pags = len(pdf.pages)
        for pi in range(0, n_pags, 2):
            txt1 = pdf.pages[pi].extract_text() or ''
            txt2 = pdf.pages[pi + 1].extract_text() if pi + 1 < n_pags else ''
            lines = [l.strip() for l in txt1.split('\n')]
            txt_all = txt1 + '\n' + txt2

            def fm(pat, txt=txt_all):
                m = re.search(pat, txt, re.I)
                return m.group(1).strip() if m else ''

            pedidoNum = fm(r'Nº\s*([\d]+/[ML])')
            dataPedido = fm(r'Data Emiss[aã]o:\s*([\d/]+)')
            dataEntrega = fm(r'Previs[aã]o Entrega:\s*([\d/]+)')
            condPgto = fm(r'Prazo para pagamento:\s*(\d+)')
            if condPgto:
                condPgto += ' dias'

            cnpj = fm(r'CNPJ:\s*([\d./\-]+)')

            filial_m = re.search(r'TORRE\s*&\s*CIA\s+SUPERMERCADOS\s+S/A\s+.+?–\s*([A-ZÁÉÍÓÚ][A-Z\s]+?)\s+(?:RIO DE|SÃO|NITERÓI)', txt_all, re.I)
            filial = filial_m.group(1).strip() if filial_m else 'TORRE CENTRAL'

            end_m = re.search(r'(AV\.[^–\n]+)', txt_all)
            endereco = end_m.group(1).strip() if end_m else ''

            cnpj_forn = fm(r'CNPJ Fornecedor:\s*([\d./\-]+)')
            empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj_forn.replace('.', '').replace('-', '') else 2

            # Parser de itens: nome vem na linha anterior ao código+dados
            reData = re.compile(
                r'^(\d{4,6})(?:\s+\d{4,6})?\s+(KG|CX)\s+(\d+)\s+([\d.]+,\d+)\s+([\d.]+,\d+)\s+([\d.]+,\d+)'
            )
            reInline = re.compile(
                r'^(\d{4,6})\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ\s]+?)\s+(KG|CX)\s+(\d+)\s+([\d.]+,\d+)\s+([\d.]+,\d+)\s+([\d.]+,\d+)'
            )
            reSufixo = re.compile(r'^[A-Z0-9]{1,5}$')

            itens = []
            pending_nome = None
            for i, ln in enumerate(lines):
                m = reData.match(ln)
                if m:
                    sufixo = ''
                    if i + 1 < len(lines) and reSufixo.match(lines[i + 1]):
                        sufixo = lines[i + 1]
                    nome = ((pending_nome or '') + (' ' + sufixo if sufixo else '')).strip()
                    qtde = float(m.group(4).replace('.', '').replace(',', '.'))
                    preco = float(m.group(5).replace('.', '').replace(',', '.'))
                    total = float(m.group(6).replace('.', '').replace(',', '.'))
                    it = processar_item(m.group(1), nome, m.group(2), int(m.group(3)), qtde, preco, total, produtos)
                    itens.append(it)
                    pending_nome = None
                else:
                    m2 = reInline.match(ln)
                    if m2:
                        qtde = float(m2.group(5).replace('.', '').replace(',', '.'))
                        preco = float(m2.group(6).replace('.', '').replace(',', '.'))
                        total = float(m2.group(7).replace('.', '').replace(',', '.'))
                        it = processar_item(m2.group(1), m2.group(2).strip(), m2.group(3), int(m2.group(4)), qtde, preco, total, produtos)
                        itens.append(it)
                        pending_nome = None
                    elif re.match(r'^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ\s]{4,}$', ln) and 'TORRE' not in ln and 'PEDIDO' not in ln and 'DADOS' not in ln:
                        pending_nome = ln

            if itens:
                filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                                 'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                                 'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
