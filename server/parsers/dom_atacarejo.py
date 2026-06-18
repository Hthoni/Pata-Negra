"""
Parser DOM Atacarejo — formato TOTVS, multi-filial.
PDF tem 2 páginas por filial: ímpar = dados do pedido, par = datas/assinaturas.
"""
import io
import re
import pdfplumber
from perfil import processar_item
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
                # tolera espaços ao redor de / e - (artefato de kerning do pdfplumber em alguns PDFs TOTVS)
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
            reItem = re.compile(r'^(\d{5,6})\s+\d+\s+(.+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)', re.M)
            itens = []
            for m in reItem.finditer(txt1):
                qtde_ped = float(m.group(5).replace('.', '').replace(',', '.'))
                preco = float(m.group(6).replace('.', '').replace(',', '.'))
                total = float(m.group(7).replace('.', '').replace(',', '.'))
                it = processar_item(int(m.group(1)), m.group(2), m.group(3),
                                     int(m.group(4)), qtde_ped, preco, total, produtos)
                itens.append(it)
            if filial and itens:
                filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                                 'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                                 'condPgto': condPgto, 'empresa': 2, 'itens': itens})
    return filiais
