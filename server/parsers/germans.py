"""
Parser Germans — formato TOTVS, itens em uma única linha (nome entre
seq e embalagem), sufixo do nome na linha seguinte.
"""
import io
import re
import pdfplumber
from perfil import processar_item

CNPJ_DISTRIBUIDORA = '56.423.719'
CNPJ_INDUSTRIA = '10.171.633'


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
                    cnpj = found[1]
                    break
                elif len(found) == 1 and CNPJ_INDUSTRIA not in found[0] and CNPJ_DISTRIBUIDORA not in found[0]:
                    cnpj = found[0]
                    break

            cnpj_forn = fm(r'CNPJ\s+([\d./\-]+)\s+Inscrição')
            empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj_forn.replace('.', '').replace('-', '') else 2

            filial_m = re.search(r'COMESTIVEI\s+(.+?)$', txt1, re.M)
            filial = filial_m.group(1).strip() if filial_m else 'CAMPEAO - CORDOVIL'
            endereco = fm(r'Endereço (RUA CORDOVIL[^\n]+?)\s{2,}')
            if not endereco:
                endereco = 'RUA CORDOVIL-1000, PARADA DE LUCAS'

            reLinhaItem = re.compile(r'^(\d{5,6})\s+\d+\s+')
            itens = []
            for i, ln in enumerate(lines):
                if not reLinhaItem.match(ln.strip()):
                    continue
                parts = ln.strip().split()
                cod = parts[0]
                for j, p in enumerate(parts):
                    if p in ('KG', 'CX') and j >= 3:
                        nome_raw = ' '.join(parts[2:j])
                        emb_tipo = p
                        try:
                            qtde_emb = int(parts[j + 1])
                            qtde_ped = float(parts[j + 2].replace('.', '').replace(',', '.'))
                            preco = float(parts[j + 4].replace('.', '').replace(',', '.'))
                            total = float(parts[j + 6].replace('.', '').replace(',', '.'))
                        except (IndexError, ValueError):
                            break
                        if i + 1 < len(lines):
                            prox = lines[i + 1].strip()
                            if prox and not prox.startswith('EANs') and not prox.startswith('TOTAIS') and len(prox) < 30:
                                sufixo = re.sub(r'\bKG\b', '', prox).strip()
                                if sufixo:
                                    nome_raw = (nome_raw + ' ' + sufixo).strip()
                        it = processar_item(cod, nome_raw, emb_tipo, qtde_emb, qtde_ped, preco, total, produtos)
                        itens.append(it)
                        break

            if itens:
                filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                                 'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                                 'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
