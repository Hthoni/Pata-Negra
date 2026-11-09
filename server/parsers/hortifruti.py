"""
Parser Hortifruti (Neogrid) — formato "PEDIDO DE COMPRA" da Neogrid, usado
pelas redes Americanas / Hortifruti (Natural da Terra).

Layout de tabela; descrições quebram em varias linhas. Usa extract_tables():
cada item vem como um campo (possivelmente multi-linha) do tipo
   [desc-A]\\n[N?] [codigo] [desc-B] (Quilograma|Unidade) [8 numeros]\\n[desc-C]
As vezes o N e o codigo se fundem (ex.: '1,007898611040613'); por isso a
descricao e reconstruida por TOKENS (descartando os puramente numericos) e a
quantidade/preco sao lidos DEPOIS da unidade de medida.

Unidade:
 - 'Quilograma' -> Qtde Pedida ja em KG.
 - 'Unidade'    -> Qtde Pedida em unidades; converte pelo peso no nome
                   (ex.: 500G -> 0,5 kg/un, 400G -> 0,4 kg/un).

Empresa (faturamento): detectada pelo CNPJ do Fornecedor:
 - 10.171.633 -> Industria  (empresa 1)
 - 56.423.719 -> Distribuidora (empresa 2)

FIX (29/07/2026): _UNID tinha \\b nos dois lados ("\\b(Quilograma|Unidade)\\b").
Quando o pdfplumber cola o fim do nome do produto direto na coluna de unidade
sem espaço (ex.: "...PATAQuilograma", visto no pedido Hortifruti 4507891618,
item "LOMBO SUINO DEF CANADENSE PATA NEGRA kg"), não existe fronteira de
palavra entre duas letras -> o \\b da esquerda falha -> o item inteiro é
descartado em silêncio no `continue`, sem erro nenhum. Removido o \\b da
esquerda (mantido o da direita, que continua protegendo contra casar no meio
de outra palavra). Comportamento para o caso comum (com espaço antes da
unidade) não muda.

FIX (09/11/2026) — LAYOUT NOVO: a partir do pedido 4508048625 (06/11) o
Hortifruti passou a mandar outro PDF (gerado por SAP/PDFium, cabeçalho
"It. Cód.Mat. Desc.Material Mat. Forn. EAN Qtde. Un. ..."). Visualmente
parecido, mas sem nenhum dos marcadores do Neogrid ("Número Pedido:",
"CNPJ do Local de Entrega:", unidade "Quilograma"/"Unidade") -> o parser
antigo devolvia 0 itens e o sistema acusava "filial não encontrada".
Agora parse() detecta o layout: Neogrid segue no código antigo
(_parse_neogrid, sem mudança); o novo vai pra _parse_sap. No novo, a
soma dos itens é conferida contra o "Total R$" do rodapé — se não bater,
para com erro claro (item perdido no regex nunca passa em silêncio).
"""

__cliente_nome__ = "Hortifruti"

import io
import re
import pdfplumber
from perfil import processar_item, match_perfil

CNPJ_INDUSTRIA = '10.171.633'
CNPJ_DISTRIBUIDORA = '56.423.719'
_UNID = re.compile(r'(Quilograma|Unidade)\b', re.I)   # \b só à direita (ver FIX acima)
_NUM = re.compile(r'\d[\d.,]*')


def _num(s):
    return float((s or '0').replace('.', '').replace(',', '.'))


# Layout SAP (09/11/2026): It. | Cód.Mat. | Descrição | Mat.Forn (opcional)
# | EAN | Qtde | Un. | Preço Unit. | Desconto | Total | Data Remessa | IPI
_ITEM_SAP = re.compile(
    r'^\s*(\d+)\s+(\d+)\s+(.+?)\s+(?:(\d+)\s+)?(\d{8,14})\s+'
    r'([\d.,]+)\s+([A-Z]{2,3})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+'
    r'(\d{2}\.\d{2}\.\d{4})', re.M)


def _parse_sap(full_text, produtos):
    def fm(pat):
        m = re.search(pat, full_text, re.I | re.M)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'^\s*(\d{6,})\s*$\s*^\s*Pedido:')
    # a data vem 1-2 linhas abaixo do rótulo (o pdfplumber intercala a razão social)
    dataEntrega = fm(r'Data de Entrega[\s\S]{0,200}?(\d{2}\.\d{2}\.\d{4})').replace('.', '/')
    cnpj = fm(r'^CNPJ:\s*([\d./\-]+)')              # 1º CNPJ = local de entrega (HNT)
    condPgto = fm(r'Cond\.\s*Pagto:\s*([^\n]+)')
    forn = re.sub(r'\D', '', fm(r'Forn\.:[^\n]*?CNPJ:\s*([\d./\-]+)'))
    empresa = 1 if forn.startswith(CNPJ_INDUSTRIA.replace('.', '')) else 2

    itens = []
    for m in _ITEM_SAP.finditer(full_text):
        nome = re.sub(r'\s+', ' ', m.group(3)).strip()
        qtde = _num(m.group(6))
        un = m.group(7).upper()
        preco = _num(m.group(8))
        total = _num(m.group(10))
        if un == 'KG':
            kg = qtde
        else:
            # UN/PC: peso pelo nome (500G -> 0,5 kg/un), mesma regra do Neogrid
            mg = re.search(r'(\d+)\s*G\b', nome, re.I)
            if not mg:
                raise ValueError(f'Hortifruti: item "{nome}" veio em {un} sem peso no nome — não sei converter pra kg')
            kg = qtde * int(mg.group(1)) / 1000.0
        it = processar_item('', nome, 'KG', 1, kg, preco, total, produtos)
        if un != 'KG':
            it['qtdeMultipl'] = qtde        # nº de unidades (unidade de faturamento)
        it['empresa'] = empresa
        itens.append(it)

    # Reconciliação com o rodapé ("Tot. Produtos" / "Total R$")
    tot_txt = fm(r'Tot\.\s*Produtos\s+([\d.,]+)') or fm(r'Total R\$\s*([\d.,]+)')
    if itens and tot_txt:
        soma = round(sum(i['valorPedido'] for i in itens), 2)
        if abs(soma - _num(tot_txt)) > 0.05:
            raise ValueError(f'Hortifruti pedido {pedidoNum}: soma dos itens lidos (R$ {soma:.2f}) '
                             f'≠ total do PDF (R$ {_num(tot_txt):.2f}) — algum item não foi lido')
    if not itens:
        return []
    return [{
        'filial': 'HORTIFRUTI', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
        'endereco': '', 'dataPedido': '', 'dataEntrega': dataEntrega,
        'condPgto': condPgto, 'empresa': empresa, 'itens': itens
    }]


def parse(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
        if 'Desc.Material' in full_text and 'Número Pedido' not in full_text:
            return _parse_sap(full_text, produtos)
        tabelas = []
        for pg in pdf.pages:
            tabelas.extend(pg.extract_tables() or [])
    return _parse_neogrid(full_text, tabelas, produtos)


def _parse_neogrid(full_text, tabelas, produtos):

    def fm(pat, txt=full_text):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'Número Pedido:\s*(\d+)')
    dataPedido = fm(r'Data de Emissão:\s*([\d/ ]+)').replace(' ', '')
    dataEntrega = fm(r'Data Inicial:\s*([\d/ ]+)').replace(' ', '')
    cnpj = fm(r'CNPJ do Local de Entrega:\s*([\d./ \-]+)').replace(' ', '')

    # Empresa pela origem de faturamento (2o CNPJ da linha Comprador/Fornecedor).
    cnpjs_linha = re.search(r'CNPJ:\s*([\d./ \-]+)\s+CNPJ:\s*([\d./ \-]+)', full_text)
    forn = (cnpjs_linha.group(2) if cnpjs_linha else '').replace(' ', '')
    empresa = 1 if CNPJ_INDUSTRIA.replace('.', '') in forn.replace('.', '').replace('-', '') else 2

    itens = []
    for t in tabelas:
        for row in t:
            for cell in row:
                if not cell:
                    continue
                mu = _UNID.search(cell)
                if not mu:
                    continue
                linha = cell.replace('\n', ' ').strip()
                mu = _UNID.search(linha)
                if not mu:
                    continue
                un = mu.group(1)
                antes = linha[:mu.start()].strip()
                depois = linha[mu.end():].strip()

                # separa as COLUNAS numéricas (logo após a unidade) da desc-C:
                # os tokens numéricos iniciais são as colunas; o resto é descrição.
                dtoks = depois.split()
                nums, resto = [], []
                for i, w in enumerate(dtoks):
                    if not resto and re.match(r'^[\d.,]+$', w):
                        nums.append(w)
                    else:
                        resto.append(w)
                if len(nums) < 6:
                    continue
                qped = _num(nums[1])                       # Qtde Pedida
                preco = _num(nums[4]) if len(nums) >= 5 else 0.0   # Preço Líquido
                total = _num(nums[7]) if len(nums) >= 8 else _num(nums[-1])  # Valor Total
                descC = ' '.join(resto).strip()

                # desc-A/B = tokens não-puramente-numéricos do "antes"
                toks = [w for w in antes.split() if not re.match(r'^[\d.,]+$', w)]
                nome = ' '.join(toks + ([descC] if descC else [])).strip()
                nome = re.sub(r'\s+', ' ', nome)

                if un.lower().startswith('quilog'):
                    kg = qped
                else:
                    mg = re.search(r'(\d+)\s*G\b', nome, re.I)
                    gramas = int(mg.group(1)) if mg else 1000
                    kg = qped * gramas / 1000.0

                it = processar_item('', nome, 'KG', 1, kg, preco, total, produtos)
                it['empresa'] = empresa
                itens.append(it)

    if not itens:
        return []
    return [{
        'filial': 'HORTIFRUTI', 'pedidoNum': pedidoNum, 'cnpj': cnpj,
        'endereco': '', 'dataPedido': dataPedido, 'dataEntrega': dataEntrega,
        'condPgto': '', 'empresa': empresa, 'itens': itens
    }]
