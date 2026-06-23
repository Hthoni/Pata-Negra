"""
Gerador do PDF de expedição — espelho das colunas A-I do Excel,
para o encarregado de embarque conferir e preencher kg embarcados.
"""
import io
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, KeepTogether, Image, PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT


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

    col_w = [8 * mm, 18 * mm, 95 * mm, 24 * mm, 18 * mm, 22 * mm, 26 * mm, 20 * mm, 46 * mm]
    W = sum(col_w)

    ST_TIT = ParagraphStyle('t', fontSize=13, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=16)
    ST_SUB = ParagraphStyle('s', fontSize=11, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=14)
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
        tkg = sum(float(it.get('kgPlanejados', 0)) for it in its)
        subtitulo = (cli + ' — ' + fd['filial']) if cli else fd['filial']

        tbl_tit = Table([[Paragraph(tit_fd, ST_TIT)]], colWidths=[W])
        tbl_tit.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), COR_TIT),
            ('TOPPADDING', (0, 0), (-1, -1), 5), ('BOTTOMPADDING', (0, 0), (-1, -1), 5)]))

        tbl_sub = Table([[Paragraph(subtitulo, ST_SUB)]], colWidths=[W])
        tbl_sub.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), COR_SUB),
            ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))

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
            Paragraph('Nome Produto no Cliente', ST_HDR), Paragraph('Formato', ST_HDR),
            Paragraph('Caixa', ST_HDR), Paragraph('Kg\nPlanejados', ST_HDR),
            Paragraph('Kgs\nEmbarcados', ST_HDR), Paragraph('Nº\nCaixas', ST_HDR),
            Paragraph('Obs.', ST_HDR),
        ]
        rows = [header]
        for idx, it in enumerate(its):
            kgcx = it.get('kgCx', 20)
            kgPlan = it.get('kgPlanejados', 0) or 0
            nrCx = int(round(kgPlan / kgcx, 0)) if kgcx else ""
            rows.append([
                Paragraph(str(idx + 1), ST_ITC),
                Paragraph(str(it.get('codInterno', '')), ST_ITC),
                Paragraph(str(it.get('nomeProduto', '')), ST_IT),
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

        story.append(KeepTogether([tbl_tit, tbl_sub, tbl_meta, tbl_itens, nota]))
        story.append(PageBreak())

    # Remover o último PageBreak (não precisa após a última filial)
    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story)
    buf.seek(0)
    return buf.read()
