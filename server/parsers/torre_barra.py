"""
Parser Torre Barra — formato TOTVS com nome em múltiplas linhas,
estrutura de colunas ligeiramente diferente do Torre Central.
Usa split() em vez de regex fixo para robustez com 1 ou 2 códigos
no início da linha.
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
            lines = [l.strip() for l in txt1.split('\n')]
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

            cnpj = fm(r'CNPJ\s+([\d./\-]+)\s+Inscrição Estadual\d')
            if not cnpj:
                for ln in lines:
                    found = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', ln)
                    validos = [c for c in found if CNPJ_DISTRIBUIDORA not in c and CNPJ_INDUSTRIA not in c]
                    if validos:
                        cnpj = validos[0]
                        break

            cnpj_forn = fm(r'CNPJ\s+([\d./\-]+)\s+Insc')
            empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in cnpj_forn.replace('.', '').replace('-', '') else 2

            filial_m = re.search(r'TORRE\s*&\s*CIA\s+SUPERMERCADOS\s+S/A\s+(.+?)$', txt1, re.M)
            filial = filial_m.group(1).strip() if filial_m else 'BARRA DA TIJUCA'
            end_m = re.search(r'Endereço\s+(AVENIDA[^\n]+)', txt1)
            endereco = end_m.group(1).strip() if end_m else ''

            reLinhaItem = re.compile(r'^(\d{4,6})(?:\s+\d{4,6})?\s+[A-ZÁÉÍÓÚ]')

            itens = []
            i = 0
            while i < len(lines):
                ln = lines[i]
                if not reLinhaItem.match(ln):
                    i += 1
                    continue
                parts = ln.split()
                cod = parts[0]
                emb_pos = None
                for j, p in enumerate(parts):
                    if p in ('KG', 'CX') and j >= 2:
                        emb_pos = j
                        break
                if emb_pos is None:
                    i += 1
                    continue
                try:
                    is_two_codes = re.match(r'^\d{4,6}$', parts[1]) if len(parts) > 1 else False
                    nome_start = 2 if is_two_codes else 1
                    nome_raw = ' '.join(parts[nome_start:emb_pos])
                    emb_tipo = parts[emb_pos]
                    qtde_emb = int(parts[emb_pos + 1])
                    qtde_ped = float(parts[emb_pos + 2].replace('.', '').replace(',', '.'))
                    # estrutura sempre: qtdePed + bonif(0,00) + precoUnit + valorItem
                    preco = float(parts[emb_pos + 4].replace('.', '').replace(',', '.'))
                    total = float(parts[emb_pos + 5].replace('.', '').replace(',', '.'))
                except (IndexError, ValueError):
                    i += 1
                    continue
                # coletar sufixos até EANs ou próximo item
                j = i + 1
                sufixo_parts = []
                while j < len(lines) and not lines[j].startswith('EANs') \
                        and not re.match(r'^\d{4,6}', lines[j]) \
                        and not lines[j].startswith('TOTAIS'):
                    s = re.sub(r'\bKG\b', '', lines[j]).strip()
                    if s:
                        sufixo_parts.append(s)
                    j += 1
                if sufixo_parts:
                    nome_raw = (nome_raw + ' ' + ' '.join(sufixo_parts)).strip()
                it = processar_item(cod, nome_raw, emb_tipo, qtde_emb, qtde_ped, preco, total, produtos)
                itens.append(it)
                i = j

            if itens:
                filiais.append({'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                                 'endereco': endereco, 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
                                 'condPgto': condPgto, 'empresa': empresa, 'itens': itens})
    return filiais
