"""
Gerador do PDF de expedição — espelho das colunas A-I do Excel,
para o encarregado de embarque conferir e preencher kg embarcados.
"""
import io
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, KeepTogether, Image, PageBreak, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import master


def _kg_por_pacote(it):
    """Peso de 1 pacote em kg = kg por caixa ÷ pacotes por caixa.
    Os pacotes por caixa vêm da embalagem no formato 'CX-N' (ex.: CX-40 -> 40);
    kg por caixa vem de kgCx. Ex.: CX-40 com kgCx=20 -> 0,5; CX-50 -> 0,4.
    Fallback: gramas explícitas na embalagem ('400G'), senão 0,5."""
    import re
    emb = str(it.get('embalagem', '')).upper()
    kgcx = float(it.get('kgCx', 0) or 0)
    m = re.search(r'CX[-\s]?(\d+)', emb)
    if m and kgcx:
        n = int(m.group(1))
        if n:
            return kgcx / n
    m = re.search(r'(\d+)\s*G\b', emb)
    if m:
        return int(m.group(1)) / 1000.0
    return 0.5


def _kg_pdf(it):
    """Kg a exibir no PDF de expedição. Itens vendidos em pacote
    (unidFat='pct') são convertidos de nº de pacotes para kg; os demais
    já estão em kg. O Excel e o pedido seguem em pacotes — só o PDF converte."""
    base = float(it.get('kgPlanejados', 0) or 0)
    if str(it.get('unidFat', '')).lower() == 'pct':
        return base * _kg_por_pacote(it)
    return base


def _nome_prod(it):
    """Nome MASTER do produto (coluna unica). Fallback: nome do cliente + aviso."""
    nm = master.nome_master(it.get('codInterno'), '')
    return nm if nm else (str(it.get('nomeProduto', '')).strip() + '  (SEM MASTER)')


def gerar_pdf(dados, empresa_override=None, logo_bytes=None):
    """Gera o PDF completo. Se empresa_override for passado, filtra
    apenas os itens daquela empresa (usado no modo split)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                             leftMargin=10 * mm, rightMargin=10 * mm,
                             topMargin=8 * mm, bottomMargin=8 * mm)

    cli = dados.get('clienteNome', '')
    vend = dados.get('vendedor', '')
    tel = dados.get('telefone', '')

    COR_TIT = colors.HexColor('#E0E0E0')
    COR_SUB = colors.HexColor('#F0F0F0')
    COR_META = colors.HexColor('#F8F8F8')
    COR_HDR = colors.HexColor('#D0D0D0')
    COR_PAR = colors.HexColor('#F5F5F5')
    COR_CZ = colors.HexColor('#E8E8E8')
    COR_RG = colors.HexColor('#8B1C1C')

    col_w = [8 * mm, 18 * mm, 95 * mm, 24 * mm, 18 * mm, 22 * mm, 26 * mm, 20 * mm, 46 * mm]
    W = sum(col_w)

    # --- Linhas 1 e 2: cliente em destaque (maior) e tipo de doc. secundário ---
    ST_TIT = ParagraphStyle('t', fontSize=10, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=12, textColor=colors.HexColor('#555555'))
    ST_SUB = ParagraphStyle('s', fontSize=16, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=19)
    # --- Campo Região (canto direito, alinhado à coluna Obs.) ---
    ST_RGL = ParagraphStyle('rgl', fontSize=8, fontName='Helvetica-Bold', alignment=TA_LEFT, leading=10, textColor=colors.HexColor('#7A746E'))
    ST_RGV = ParagraphStyle('rgv', fontSize=14, fontName='Helvetica-Bold', alignment=TA_LEFT, leading=16, textColor=COR_RG)
    ST_MB = ParagraphStyle('mb', fontSize=10, fontName='Helvetica-Bold', alignment=TA_LEFT, leading=12)
    ST_MV = ParagraphStyle('mv', fontSize=10, fontName='Helvetica', alignment=TA_LEFT, leading=12)
    ST_HDR = ParagraphStyle('h', fontSize=10, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=11)
    ST_IT = ParagraphStyle('i', fontSize=10, fontName='Helvetica', alignment=TA_LEFT, leading=11)
    ST_ITC = ParagraphStyle('ic', fontSize=10, fontName='Helvetica', alignment=TA_CENTER, leading=11)
    ST_ITR = ParagraphStyle('ir', fontSize=10, fontName='Helvetica', alignment=TA_RIGHT, leading=11)
    ST_TOT = ParagraphStyle('to', fontSize=10, fontName='Helvetica-Bold', alignment=TA_LEFT, leading=11)
    ST_TOTR = ParagraphStyle('tr', fontSize=10, fontName='Helvetica-Bold', alignment=TA_RIGHT, leading=11)
    ST_NOTA = ParagraphStyle('n', fontSize=8, fontName='Helvetica', textColor=colors.HexColor('#666666'))

    story = []

    for fd in dados['filiais']:
        if empresa_override:
            its = [it for it in fd['itens']
                   if (it.get('empresa') or empresa_override) == empresa_override]
        else:
            its = fd['itens']
        if not its:
            continue
        n = len(its)
        emp_fd = empresa_override if empresa_override else fd.get('empresa', dados.get('empresa', 2))
        tit_fd = 'PEDIDO PATA NEGRA DISTRIBUIDORA' if emp_fd == 2 else 'PEDIDO INDÚSTRIA PATA NEGRA'
        nota_empresa = 'Pata Negra Distribuidora' if emp_fd == 2 else 'Indústria Pata Negra'
        tkg = sum(_kg_pdf(it) for it in its)
        subtitulo = (cli + ' — ' + fd['filial']) if cli else fd['filial']

        # --- Cabeçalho: linhas 1 e 2 DESMESCLADAS + campo Região alinhado à coluna Obs. ---
        REGIAO_W = col_w[8]  # largura da coluna Obs. -> alinha o campo Região por ela
        regiao_val = fd.get('regiao', '') or '—'
        regiao_cell = [Paragraph('REGIÃO', ST_RGL), Paragraph(regiao_val, ST_RGV)]
        cab = [
            [Paragraph(tit_fd, ST_TIT), regiao_cell],
            [Paragraph(subtitulo, ST_SUB), ''],
        ]
        tbl_cab = Table(cab, colWidths=[W - REGIAO_W, REGIAO_W])
        tbl_cab.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), COR_TIT),
            ('BACKGROUND', (0, 1), (0, 1), COR_SUB),
            ('SPAN', (1, 0), (1, 1)),
            ('BACKGROUND', (1, 0), (1, 1), colors.white),
            ('BOX', (1, 0), (1, 1), 1, COR_RG),
            ('ALIGN', (0, 0), (0, 1), 'CENTER'),
            ('VALIGN', (0, 0), (0, 1), 'MIDDLE'),
            ('VALIGN', (1, 0), (1, 1), 'MIDDLE'),
            ('LEFTPADDING', (1, 0), (1, 1), 8),
            ('TOPPADDING', (0, 0), (0, 0), 5), ('BOTTOMPADDING', (0, 0), (0, 0), 2),
            ('TOPPADDING', (0, 1), (0, 1), 2), ('BOTTOMPADDING', (0, 1), (0, 1), 5),
            ('TOPPADDING', (1, 0), (1, 1), 4), ('BOTTOMPADDING', (1, 0), (1, 1), 4),
        ]))

        def ml(t):
            return Paragraph(t, ST_MB)

        def mv(t):
            return Paragraph(str(t) if t else '—', ST_MV)

        def fl(t):
            return Paragraph(str(t) if t else '—', ST_MV)

        meta = [
            [ml('Pedido Nº:'), mv(fd.get('pedidoNum', '')), ml('Data Pedido:'), mv(fd.get('dataPedido', ''))],
            [ml('CNPJ:'), mv(fd.get('cnpj', '')), ml('Data Entrega:'), mv(fd.get('dataEntrega', ''))],
            [ml('Filial:'), Paragraph((str(fd['filial']) + ('   |   Núm. Filial: ' + str(fd['numFilial']) if fd.get('numFilial') is not None else '')), ST_MV), ml('Solicitante:'), mv(fd.get('solicitante', ''))],
            [ml('Endereço:'), mv(fd.get('endereco', '')), ml('Vendedor:'), mv(vend)],
            [ml('Cond. Pgto.:'), mv(fd.get('condPgto', '')), ml('Tel. Vendedor:'), mv(tel)],
        ]
        LOGO_W = 38 * mm
        VAL4_W = W - 25 * mm - 100 * mm - 32 * mm - LOGO_W
        if logo_bytes:
            try:
                # Preservar proporção: deixar ReportLab calcular dimensões reais
                # e escalar respeitando max_w e max_h
                _tmp = Image(io.BytesIO(logo_bytes))
                orig_w, orig_h = _tmp.imageWidth, _tmp.imageHeight
                max_h = 22 * mm
                max_w = LOGO_W - 4 * mm
                ratio = min(max_w / orig_w, max_h / orig_h)
                logo_img = Image(io.BytesIO(logo_bytes), width=orig_w * ratio, height=orig_h * ratio)
                logo_img.hAlign = 'CENTER'
                for row in meta:
                    row.append('')
                meta[0][4] = logo_img
                tbl_meta = Table(meta, colWidths=[25*mm, 100*mm, 32*mm, VAL4_W, LOGO_W])
                tbl_meta.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), COR_META),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('LINEBELOW', (0, 4), (-1, 4), 0.5, colors.grey),
                    ('SPAN', (4, 0), (4, 4)),
                    ('VALIGN', (4, 0), (4, 4), 'MIDDLE'),
                    ('ALIGN', (4, 0), (4, 4), 'CENTER'),
                ]))
            except Exception:
                tbl_meta = Table(meta, colWidths=[25*mm, 100*mm, 32*mm, W - 25*mm - 100*mm - 32*mm])
                tbl_meta.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), COR_META),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('LINEBELOW', (0, 4), (-1, 4), 0.5, colors.grey),
                ]))
        else:
            tbl_meta = Table(meta, colWidths=[25*mm, 100*mm, 32*mm, W - 25*mm - 100*mm - 32*mm])
            tbl_meta.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), COR_META),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('LINEBELOW', (0, 4), (-1, 4), 0.5, colors.grey),
            ]))

        header = [
            Paragraph('#', ST_HDR), Paragraph('Cód.\nInterno', ST_HDR),
            Paragraph('Produtos', ST_HDR), Paragraph('Formato', ST_HDR),
            Paragraph('Caixa', ST_HDR), Paragraph('Kg\nPlanejados', ST_HDR),
            Paragraph('Kgs\nEmbarcados', ST_HDR), Paragraph('Nº\nCaixas', ST_HDR),
            Paragraph('Obs.', ST_HDR),
        ]
        rows = [header]
        for idx, it in enumerate(its):
            kgcx = it.get('kgCx', 20)
            kgPlan = _kg_pdf(it)
            nrCx = int(round(kgPlan / kgcx, 0)) if kgcx else ""
            rows.append([
                Paragraph(str(idx + 1), ST_ITC),
                Paragraph(str(it.get('codInterno', '')), ST_ITC),
                Paragraph(_nome_prod(it), ST_IT),
                Paragraph(str(it.get('formato', '') or ''), ST_ITC),
                Paragraph(str(it.get('embalagem', '')), ST_ITC),
                Paragraph(f"{kgPlan:.1f}".replace(".",",") if kgPlan else "", ST_ITR),
                Paragraph('', ST_ITC),
                Paragraph(str(nrCx) if nrCx else '', ST_ITR),
                Paragraph(str(it.get('obs', '') or ''), ST_IT),
            ])

        rows.append([
            Paragraph(f'TOTAL — {n} itens', ST_TOT), '', '', '', '',
            Paragraph(f"{tkg:.1f}".replace(".",","), ST_TOTR), '', '', ''
        ])

        tbl_itens = Table(rows, colWidths=col_w, repeatRows=1)
        cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), COR_HDR),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -2), 0.5, colors.HexColor('#BBBBBB')),
            ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('BACKGROUND', (6, 1), (6, -2), COR_CZ),
            ('BACKGROUND', (0, -1), (-1, -1), COR_HDR),
            ('SPAN', (0, -1), (4, -1)),
            ('LINEABOVE', (0, -1), (-1, -1), 0.8, colors.grey),
            ('LINEBELOW', (0, -1), (-1, -1), 0.8, colors.grey),
        ]
        for i in range(1, len(rows) - 1):
            if i % 2 == 0:
                cmds.append(('BACKGROUND', (0, i), (-1, i), COR_PAR))
        tbl_itens.setStyle(TableStyle(cmds))

        nota = Paragraph(
            f'Col. G — Kg Embarcados: preenchida pelo encarregado de embarque.   |   {nota_empresa}',
            ST_NOTA)

        story.append(KeepTogether([tbl_cab, tbl_meta, tbl_itens, nota]))
        story.append(PageBreak())

    # Remover o último PageBreak (não precisa após a última filial)
    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story)
    buf.seek(0)
    return buf.read()


def gerar_pdf_totais(produtos, pedidos, meta):
    """PDF de simulacao de entregas: soma por produto (nome MASTER, A->Z) +
    a lista dos pedidos incluidos. produtos = [{nome, kg, alerta}], ja ordenado.
    pedidos = [{cliente, filial}]. meta = {data, nPedidos, totalKg, semItens}."""
    def _fmt(v):
        return (f"{float(v):.1f}".replace(".", ",")) if v else "0"

    ACCENT = colors.HexColor('#8B1C1C')
    GREYBG = colors.HexColor('#F1EFEA')
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=16 * mm, bottomMargin=16 * mm,
                            leftMargin=16 * mm, rightMargin=16 * mm,
                            title='Simulacao de Entregas')
    S_TIT = ParagraphStyle('tit', fontName='Helvetica-Bold', fontSize=15, textColor=ACCENT, spaceAfter=7)
    S_SUB = ParagraphStyle('sub', fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#555555'), leading=14, spaceAfter=12)
    S_WARN = ParagraphStyle('warn', fontName='Helvetica-Bold', fontSize=9, textColor=ACCENT, spaceBefore=4, spaceAfter=2)
    S_SEC = ParagraphStyle('sec', fontName='Helvetica-Bold', fontSize=11, textColor=colors.black, spaceBefore=14, spaceAfter=6)
    S_HDR = ParagraphStyle('hdr', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.white)
    S_HDRR = ParagraphStyle('hdrr', parent=S_HDR, alignment=TA_RIGHT)
    S_IT = ParagraphStyle('it', fontName='Helvetica', fontSize=10, textColor=colors.black)
    S_ITR = ParagraphStyle('itr', parent=S_IT, alignment=TA_RIGHT)
    S_AL = ParagraphStyle('al', fontName='Helvetica-Bold', fontSize=10, textColor=ACCENT)
    S_PED = ParagraphStyle('ped', fontName='Helvetica', fontSize=9.5, textColor=colors.HexColor('#333333'), leading=14)

    story = []
    titulo = meta.get('titulo') or 'Simulacao de Entregas &mdash; Totais por Produto'
    story.append(Paragraph(titulo, S_TIT))
    story.append(Paragraph(f"{meta.get('data','')} &nbsp;&bull;&nbsp; {meta.get('nPedidos',0)} pedido(s) &nbsp;&bull;&nbsp; Total: <b>{_fmt(meta.get('totalKg',0))} kg</b>", S_SUB))
    if meta.get('semItens'):
        story.append(Paragraph(f"&#9888; {meta['semItens']} pedido(s) sem detalhamento (gerados antes desta atualizacao) &mdash; nao somados.", S_WARN))

    story.append(Spacer(1, 5 * mm))   # respiro entre o cabeçalho e a faixa vermelha
    data = [[Paragraph('Produto', S_HDR), Paragraph('Kg total', S_HDRR)]]
    for p in produtos:
        st = S_AL if p.get('alerta') else S_IT
        data.append([Paragraph(str(p.get('nome', '')), st), Paragraph(_fmt(p.get('kg', 0)), S_ITR)])
    tbl = Table(data, colWidths=[132 * mm, 46 * mm], repeatRows=1)
    style = [
        ('BACKGROUND', (0, 0), (-1, 0), ACCENT),
        ('TOPPADDING', (0, 0), (-1, -1), 5), ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 8), ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('LINEBELOW', (0, 0), (-1, -1), 0.4, colors.HexColor('#DDDDDD')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]
    for i in range(1, len(data) + 1):
        if i % 2 == 0:
            style.append(('BACKGROUND', (0, i), (-1, i), GREYBG))
    tbl.setStyle(TableStyle(style))
    story.append(tbl)

    if pedidos:
        story.append(Paragraph(f"Pedidos incluidos ({len(pedidos)})", S_SEC))
        for pd in sorted(pedidos, key=lambda x: (x.get('cliente', '') or '').lower()):
            cli = pd.get('cliente', '') or ''
            fil = pd.get('filial', '') or ''
            story.append(Paragraph(f"&bull; {cli} &mdash; {fil}" if fil else f"&bull; {cli}", S_PED))

    doc.build(story)
    buf.seek(0)
    return buf.read()
