from flask import Flask, request, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import pdfplumber
import re, io, base64, json, os, traceback
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ════════════════════════════════════════════════
# ESTILOS EXCEL
# ════════════════════════════════════════════════
def thin(): return Side(style='thin')
def brd(t=0,b=0,l=0,r=0): return Border(
    top=thin()if t else Side(), bottom=thin()if b else Side(),
    left=thin()if l else Side(), right=thin()if r else Side())
def fill(rgb): return PatternFill(fill_type='solid', fgColor=rgb)
def fnt(bold=False,sz=9,color='FF000000'): return Font(bold=bold,size=sz,color=color)
def aln(h='left',wrap=False): return Alignment(horizontal=h,vertical='center',wrap_text=wrap)

BALL=brd(1,1,1,1); BBOT=brd(b=1)
BG_TIT=fill('FFE0E0E0'); BG_SUB=fill('FFF0F0F0'); BG_META=fill('FFF8F8F8')
BG_HDR=fill('FFD0D0D0'); BG_AZL=fill('FF2980B9'); BG_AZLT=fill('FFEBF5FB')
BG_CZ=fill('FFE8E8E8'); BG_PAR=fill('FFF5F5F5'); BG_TOTV=fill('FFD5F5E3')
F_TIT=fnt(True,11); F_SUB=fnt(True,10); F_MB=fnt(True,9); F_M=fnt(sz=9)
F_HW=fnt(True,9,'FFFFFFFF'); F_HB=fnt(True,9); F_IT=fnt(sz=9); F_TOT=fnt(True,9)
F_NT=fnt(sz=8,color='FF666666'); F_NR=fnt(sz=8,color='FFC0392B'); F_CX=fnt(True,9,'FF854F0B')
FMT_MONEY='#,##0.00'; FMT_NUM='0.0'

def ap(cell, val=None, fn=None, bg=None, bo=None, al=None, fmt=None):
    if val is not None: cell.value = val
    if fn:  cell.font = fn
    if bg:  cell.fill = bg
    if bo:  cell.border = bo
    if al:  cell.alignment = al
    if fmt: cell.number_format = fmt

# ════════════════════════════════════════════════
# GERADOR EXCEL
# ════════════════════════════════════════════════
def gerar_excel(dados):
    emp  = dados['empresa']
    cv   = str(dados['codVend'])
    cc   = str(dados['codCond'])
    vend = dados.get('vendedor','')
    tel  = dados.get('telefone','')
    tit  = 'PEDIDO PATA NEGRA DISTRIBUIDORA' if emp==2 else 'PEDIDO INDUSTRIA PATANEGRA'

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for fd in dados['filiais']:
        its = fd['itens']; n = len(its)
        numPed = fd.get('pedidoNum','').replace('/C','').replace('/L','').strip()
        nomeAba = ((numPed+' - '+fd['filial']) if numPed else fd['filial'])[:31]
        ws = wb.create_sheet(nomeAba)

        for i,w in enumerate([4,10,38,12,7,12,12,9,14,2,11,12,14,12,10,8],1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # L1 A1:I1
        ws.merge_cells('A1:I1')
        ap(ws['A1'], tit, F_TIT, BG_TIT, al=aln('center'))

        # L2 A2:I2
        ws.merge_cells('A2:I2')
        ap(ws['A2'], 'DOM Atacarejo — '+fd['filial'], F_SUB, BG_SUB, al=aln('center'))

        # L3-L7: A:B, C:F, H:I
        for r in [3,4,5,6,7]:
            ws.merge_cells(f'A{r}:B{r}')
            ws.merge_cells(f'C{r}:F{r}')
            ws.merge_cells(f'H{r}:I{r}')

        ap(ws['A3'],'Pedido Nº:',F_MB,BG_META,al=aln('left'))
        ap(ws['C3'],fd.get('pedidoNum',''),F_M,BG_META)
        ap(ws['G3'],'Data Pedido:',F_MB,BG_META,al=aln('left'))
        ap(ws['H3'],fd.get('dataPedido',''),F_M,BG_META,al=aln('left'))

        ap(ws['A4'],'CNPJ:',F_MB,BG_META,al=aln('left'))
        ap(ws['C4'],fd.get('cnpj',''),F_M,BG_META)
        ap(ws['G4'],'Data Entrega:',F_MB,BG_META,al=aln('left'))
        ap(ws['H4'],fd.get('dataEntrega',''),F_M,BG_META,al=aln('left'))

        ap(ws['A5'],'Filial:',F_MB,BG_META,al=aln('left'))
        ap(ws['C5'],fd['filial'],F_M,BG_META)
        ap(ws['G5'],'Solicitante:',F_MB,BG_META,al=aln('left'))
        ap(ws['K5'],'Código vend.',F_HW,BG_AZL,BALL,aln('center',True))
        ap(ws['L5'],'código cond', F_HW,BG_AZL,BALL,aln('center',True))
        ap(ws['M5'],'empresa',     F_HW,BG_AZL,BALL,aln('center',True))

        ap(ws['A6'],'Endereço:',F_MB,BG_META,al=aln('left'))
        ap(ws['C6'],fd.get('endereco',''),F_M,BG_META)
        ap(ws['G6'],'Vendedor:',F_MB,BG_META,al=aln('left'))
        ap(ws['H6'],vend,F_M,BG_META,al=aln('left'))
        ws['K6'].value = int(cv) if cv.isdigit() else cv
        ws['L6'].value = int(cc) if cc.isdigit() else cc
        ws['M6'].value = emp

        ap(ws['A7'],'Cond. Pgto.:',F_MB,BG_META,BBOT,aln('left'))
        ap(ws['C7'],fd.get('condPgto',''),F_M,BG_META,BBOT)
        for col in ['B7','D7','E7','F7']: ws[col].border = BBOT
        ap(ws['G7'],'Tel. Vendedor:',F_MB,BG_META,al=aln('left'))
        ap(ws['H7'],tel,F_M,BG_META,al=aln('left'))

        # L8 Headers
        for col,h,az in [('A','#',0),('B','Cód.\nInterno',0),('C','Nome Produto\nno Cliente',0),
            ('D','Formato',0),('E','Caixa',0),('F','Kg\nPlanejados',0),('G','Kg\nEmbarcados',0),
            ('H','Nº\nCaixas',0),('I','Obs.',0),('K','Qtde\nMultipl.',1),('L','Preço Unit.\n(R$)',1),
            ('M','Valor Pedido\n(R$)',1),('N','Preço\nSistema',1),('O','Dif.\nPreço',1),('P','Unid.\nfat.',1)]:
            ap(ws[f'{col}8'],h,F_HW if az else F_HB,BG_AZL if az else BG_HDR,BALL,aln('center',True))
        ws.row_dimensions[8].height = 28

        # Itens
        for idx,it in enumerate(its):
            r = idx+9; par=(idx%2==1); bg=BG_PAR if par else None
            isCx=(it.get('unidFat','kg')=='cx'); kgcx=it.get('kgCx',20)
            def c(col): return ws[f'{col}{r}']
            ap(c('A'),idx+1,F_IT,bg,BALL,aln('center'))
            ap(c('B'),it.get('codInterno'),F_IT,bg,BALL,aln('center'))
            ap(c('C'),it.get('nomeProduto',''),F_IT,bg,BALL,aln('left'))
            ap(c('D'),it.get('formato') or None,F_IT,bg,BALL,aln('center'))
            ap(c('E'),it.get('embalagem',''),F_IT,bg,BALL,aln('center'))
            ap(c('F'),it.get('kgPlanejados'),F_IT,bg,BALL,aln('center'),FMT_NUM)
            ap(c('G'),None,F_IT,BG_CZ,BALL,aln('center'))
            c('H').value = f'=IF(G{r}<>"",G{r}/{kgcx},F{r}/{kgcx})'
            ap(c('H'),None,F_IT,bg,BALL,aln('center'),FMT_NUM)
            ap(c('I'),it.get('obs') or None,F_IT,bg,BALL,aln('center'))
            ap(c('K'),it.get('qtdeMultipl'),F_IT,BG_AZLT,BALL,aln('center'))
            ap(c('L'),it.get('precoUnit'),F_IT,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            c('M').value = f'=L{r}*K{r}' if isCx else f'=IF(G{r}<>"",L{r}*G{r},L{r}*F{r})'
            ap(c('M'),None,F_IT,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            ap(c('N'),it.get('precoSistema') or 0,F_IT,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            c('O').value = f'=IF(AND(L{r}<>"",N{r}<>"",L{r}<>N{r}),L{r}-N{r},"")'
            ap(c('O'),None,F_IT,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            ap(c('P'),it.get('unidFat','kg'),F_CX if isCx else F_IT,BG_AZLT,BALL,aln('center'))

        # TOTAL
        rT=n+9; r1=9; r2=n+8
        ws.merge_cells(f'A{rT}:E{rT}')
        ap(ws[f'A{rT}'],f'TOTAL  —  {n} itens',F_TOT,None,brd(1,1,1,0),aln('left'))
        for col in ['B','C','D','E']: ws[f'{col}{rT}'].border = brd(1,1,0,0)
        for col in ['F','G','H']:
            ws[f'{col}{rT}'].value = f'=SUM({col}{r1}:{col}{r2})'
            ap(ws[f'{col}{rT}'],None,F_TOT,BG_HDR,BALL,aln('center'),FMT_NUM)
        ws[f'M{rT}'].value = f'=SUM(M{r1}:M{r2})'
        ap(ws[f'M{rT}'],None,F_TOT,BG_TOTV,BALL,aln('center'),FMT_MONEY)

        # Instrução
        rI = rT+2
        ws.merge_cells(f'A{rI}:I{rI}')
        ws.merge_cells(f'K{rI}:P{rI}')
        ap(ws[f'A{rI}'],f'Imprimir: selecionar A1:I{rT-1} → Imprimir Seleção',F_NT,al=aln('left'))
        ap(ws[f'K{rI}'],'Col K=qtde faturamento | Col P=unidade | Dif. Preço vermelho=revisar',F_NR,al=aln('left'))

    # RESUMO GERAL
    wsr = wb.create_sheet('RESUMO GERAL')
    for i,w in enumerate([22,12,22,6,12,14],1): wsr.column_dimensions[get_column_letter(i)].width=w
    for ci,h in enumerate(['Filial','Pedido Nº','CNPJ','Itens','Kg Plan.','Valor (R$)'],1):
        ap(wsr.cell(1,ci),h,F_HB,BG_HDR,BALL,aln('center'))
    skg=sv=sit=0
    for ri,fd in enumerate(dados['filiais'],2):
        its2=fd['itens']
        kg=sum(i.get('kgPlanejados',0) for i in its2)
        val=sum(i.get('valorPedido',0) for i in its2)
        skg+=kg; sv+=val; sit+=len(its2)
        par=(ri%2==0); bg2=BG_PAR if par else None
        for ci,v in enumerate([fd['filial'],fd.get('pedidoNum',''),fd.get('cnpj',''),
                                len(its2),round(kg,1),round(val,2)],1):
            fmt2=FMT_MONEY if ci==6 else (FMT_NUM if ci==5 else None)
            ap(wsr.cell(ri,ci),v,F_IT,bg2,BALL,fmt=fmt2)
    rTR=len(dados['filiais'])+2
    for ci,v in enumerate(['TOTAL','','',sit,round(skg,1),round(sv,2)],1):
        fmt2=FMT_MONEY if ci==6 else (FMT_NUM if ci==5 else None)
        ap(wsr.cell(rTR,ci),v,F_TOT,BG_HDR,BALL,fmt=fmt2)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()

# ════════════════════════════════════════════════
# GERADOR PDF — espelho das colunas A:I do Excel
# ════════════════════════════════════════════════
def gerar_pdf(dados):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=8*mm, bottomMargin=8*mm)

    emp = dados['empresa']
    tit = 'PEDIDO PATA NEGRA DISTRIBUIDORA' if emp==2 else 'PEDIDO INDUSTRIA PATANEGRA'
    vend = dados.get('vendedor','')

    COR_TIT  = colors.HexColor('#E0E0E0')
    COR_SUB  = colors.HexColor('#F0F0F0')
    COR_META = colors.HexColor('#F8F8F8')
    COR_HDR  = colors.HexColor('#D0D0D0')
    COR_PAR  = colors.HexColor('#F5F5F5')
    COR_CZ   = colors.HexColor('#E8E8E8')

    W = landscape(A4)[0] - 20*mm

    ST_TIT  = ParagraphStyle('t', fontSize=11, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=14)
    ST_SUB  = ParagraphStyle('s', fontSize=10, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=12)
    ST_MB   = ParagraphStyle('mb',fontSize=8,  fontName='Helvetica-Bold', alignment=TA_LEFT)
    ST_MV   = ParagraphStyle('mv',fontSize=8,  fontName='Helvetica',      alignment=TA_LEFT)
    ST_HDR  = ParagraphStyle('h', fontSize=8,  fontName='Helvetica-Bold', alignment=TA_CENTER, leading=9)
    ST_IT   = ParagraphStyle('i', fontSize=8,  fontName='Helvetica',      alignment=TA_LEFT,   leading=9)
    ST_ITC  = ParagraphStyle('ic',fontSize=8,  fontName='Helvetica',      alignment=TA_CENTER, leading=9)
    ST_ITR  = ParagraphStyle('ir',fontSize=8,  fontName='Helvetica',      alignment=TA_RIGHT,  leading=9)
    ST_TOT  = ParagraphStyle('to',fontSize=8,  fontName='Helvetica-Bold', alignment=TA_LEFT,   leading=9)
    ST_NOTA = ParagraphStyle('n', fontSize=7,  fontName='Helvetica',      textColor=colors.HexColor('#666666'))

    # Larguras das colunas A:I
    # A(#) B(Cód) C(Nome) D(Formato) E(Caixa) F(KgPlan) G(KgEmb) H(NrCx) I(Obs)
    col_w = [8*mm, 16*mm, 88*mm, 24*mm, 16*mm, 20*mm, 20*mm, 14*mm, 14*mm]

    story = []

    for fd in dados['filiais']:
        its = fd['itens']; n = len(its)
        tkg = sum(float(it.get('kgPlanejados',0)) for it in its)

        # Título
        tbl_tit = Table([[Paragraph(tit, ST_TIT)]], colWidths=[W])
        tbl_tit.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),COR_TIT),
            ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]))

        # Subtítulo
        tbl_sub = Table([[Paragraph('DOM Atacarejo — '+fd['filial'], ST_SUB)]], colWidths=[W])
        tbl_sub.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),COR_SUB),
            ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3)]))

        # Metadados
        def ml(t): return Paragraph(t, ST_MB)
        def mv(t): return Paragraph(str(t) if t else '—', ST_MV)

        meta = [
            [ml('Pedido Nº:'), mv(fd.get('pedidoNum','')),  ml('Data Pedido:'),  mv(fd.get('dataPedido',''))],
            [ml('CNPJ:'),      mv(fd.get('cnpj','')),       ml('Data Entrega:'), mv(fd.get('dataEntrega',''))],
            [ml('Filial:'),    mv(fd['filial']),             ml('Solicitante:'),  mv('')],
            [ml('Endereço:'),  mv(fd.get('endereco','')),   ml('Vendedor:'),     mv(vend)],
            [ml('Cond. Pgto.:'), mv(fd.get('condPgto','')), ml('Tel. Vendedor:'),mv(dados.get('telefone',''))],
        ]
        tbl_meta = Table(meta, colWidths=[22*mm, 88*mm, 28*mm, 82*mm])
        tbl_meta.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),COR_META),
            ('FONTSIZE',(0,0),(-1,-1),8),
            ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
            ('LEFTPADDING',(0,0),(-1,-1),3),
            ('LINEBELOW',(0,4),(-1,4),0.5,colors.grey),
        ]))

        # Tabela de itens — espelho A:I do Excel
        header = [
            Paragraph('#',ST_HDR), Paragraph('Cód.\nInterno',ST_HDR),
            Paragraph('Nome Produto no Cliente',ST_HDR), Paragraph('Formato',ST_HDR),
            Paragraph('Caixa',ST_HDR), Paragraph('Kg\nPlanejados',ST_HDR),
            Paragraph('Kg\nEmbarcados',ST_HDR), Paragraph('Nº\nCaixas',ST_HDR),
            Paragraph('Obs.',ST_HDR),
        ]
        rows = [header]
        for idx,it in enumerate(its):
            kgcx = it.get('kgCx',20)
            kgPlan = it.get('kgPlanejados',0)
            nrCx = round(kgPlan/kgcx,1) if kgcx else ''
            rows.append([
                Paragraph(str(idx+1),ST_ITC),
                Paragraph(str(it.get('codInterno','')),ST_ITC),
                Paragraph(str(it.get('nomeProduto','')),ST_IT),
                Paragraph(str(it.get('formato','') or ''),ST_ITC),
                Paragraph(str(it.get('embalagem','')),ST_ITC),
                Paragraph(f"{kgPlan:.1f}" if kgPlan else '',ST_ITR),
                Paragraph('',ST_ITC),  # Kg Embarcados — vazio (preenchido à mão)
                Paragraph(str(nrCx) if nrCx else '',ST_ITR),
                Paragraph(str(it.get('obs','') or ''),ST_IT),
            ])

        # Linha TOTAL
        rows.append([
            Paragraph(f'TOTAL — {n} itens', ST_TOT),'','','','',
            Paragraph(f'{tkg:.1f}', ParagraphStyle('tr',fontSize=8,fontName='Helvetica-Bold',alignment=TA_RIGHT)),
            '','',''
        ])

        tbl_itens = Table(rows, colWidths=col_w, repeatRows=1)
        cmds = [
            ('BACKGROUND',(0,0),(-1,0),COR_HDR),
            ('FONTSIZE',(0,0),(-1,-1),8),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('GRID',(0,0),(-1,-2),0.4,colors.HexColor('#CCCCCC')),
            ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
            ('LEFTPADDING',(0,0),(-1,-1),2),('RIGHTPADDING',(0,0),(-1,-1),2),
            # Col G (Kg Embarcados) cinza
            ('BACKGROUND',(6,1),(6,-2),COR_CZ),
            # TOTAL
            ('BACKGROUND',(0,-1),(-1,-1),COR_HDR),
            ('SPAN',(0,-1),(4,-1)),
            ('LINEABOVE',(0,-1),(-1,-1),0.5,colors.grey),
            ('LINEBELOW',(0,-1),(-1,-1),0.5,colors.grey),
        ]
        for i in range(1,len(rows)-1):
            if i%2==0: cmds.append(('BACKGROUND',(0,i),(-1,i),COR_PAR))
        tbl_itens.setStyle(TableStyle(cmds))

        nota = Paragraph(
            'Col. G — Kg Embarcados: preenchida pelo encarregado de embarque.   |   '
            'Pata Negra Distribuidora',
            ST_NOTA)

        story.append(KeepTogether([tbl_tit, tbl_sub, tbl_meta, tbl_itens, nota]))

    doc.build(story)
    buf.seek(0)
    return buf.read()

# ════════════════════════════════════════════════
# PARSER PDF DOM ATACAREJO
# ════════════════════════════════════════════════
def match_perfil(nome, produtos):
    n = nome.lower().strip()
    best, score = None, 0
    for p in produtos:
        a = (p.get('nomePerfil','') or '').lower().strip()
        if not a: continue
        if a == n: s = 100
        elif n in a: s = 50+len(a)
        elif a in n: s = 40+len(n)
        else:
            common = len(set(a.split()) & set(n.split()))
            s = common*10 if common>=2 else 0
        if s > score: score,best = s,p
    return best

def parse_dom_atacarejo(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pags = len(pdf.pages)
        for pi in range(0, n_pags, 2):
            txt1 = pdf.pages[pi].extract_text() or ''
            txt2 = pdf.pages[pi+1].extract_text() if pi+1 < n_pags else ''
            lines = txt1.split('\n')
            txt_all = txt1+'\n'+txt2

            def fm(pat, txt):
                m = re.search(pat, txt, re.I)
                return m.group(1) if m else ''

            pedidoNum = fm(r'(\d{5,7}/[CL])', txt_all)
            filial = ''
            for ln in lines:
                if 'DOM ATACAREJO SA' in ln:
                    filial = re.sub(r'\s+R\..*$','', re.sub(r'.*DOM ATACAREJO SA\s+','',ln)).strip()
                    break
            cnpj = ''
            for ln in lines:
                found = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', ln)
                if len(found)>=2: cnpj=found[1]; break
                elif len(found)==1: cnpj=found[0]; break
            endereco = ''
            for i,ln in enumerate(lines):
                if 'ENDEREÇO PARA ENTREGA' in ln:
                    for j in range(i+1,min(i+5,len(lines))):
                        if lines[j].startswith('Endereço'):
                            endereco = re.split(r'\s{2,}', lines[j].replace('Endereço','').strip())[0].strip()
                            break
                    break
            dataPedido  = fm(r'Data da emiss[aã]o\s+([\d/]+)', txt_all)
            dataEntrega = fm(r'Previs[aã]o de entrega\s+([\d/]+)', txt_all)
            condPgto    = fm(r'Prazo para pagamento\s+(\d+)', txt_all)
            if condPgto: condPgto += ' dias'

            reItem = re.compile(r'^(\d{5,6})\s+\d+\s+(.+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)')
            itens = []
            for ln in lines:
                mm = reItem.match(ln)
                if not mm: continue
                embTipo = mm.group(3); qtdeEmb = int(mm.group(4))
                qtdePed   = float(mm.group(5).replace('.','').replace(',','.'))
                precoUnit = float(mm.group(6).replace('.','').replace(',','.'))
                valorTot  = float(mm.group(7).replace('.','').replace(',','.'))
                nomeRaw = mm.group(2).strip()
                pf = match_perfil(nomeRaw, produtos)
                kgCx     = pf['kgCx']      if pf else 20
                embalagem= pf['embalagem']  if pf else ('CX-'+str(qtdeEmb) if embTipo=='CX' else 'CX-20')
                if embTipo=='CX':
                    kgPlan=qtdePed*kgCx; nrCx=qtdePed; qtdeMult=qtdePed; unidFat='cx'
                else:
                    kgPlan=qtdePed; nrCx=round(kgPlan/kgCx,1) if kgCx else 0; qtdeMult=kgPlan; unidFat='kg'
                itens.append({
                    'codInterno':  pf['codInterno']  if pf else int(mm.group(1)),
                    'nomeProduto': pf['nomePerfil']  if pf else nomeRaw,
                    'formato':     pf.get('formato','')  if pf else '',
                    'embalagem':   embalagem, 'kgCx': kgCx,
                    'kgPlanejados': kgPlan, 'nrCaixas': nrCx,
                    'obs':         pf.get('obs','') if pf else '',
                    'qtdeMultipl': qtdeMult, 'precoUnit': precoUnit,
                    'valorPedido': valorTot,
                    'precoSistema': pf['precoUnit'] if pf else 0,
                    'unidFat':     unidFat,
                })
            if filial and itens:
                filiais.append({
                    'filial': filial, 'pedidoNum': pedidoNum, 'cnpj': cnpj,
                    'endereco': endereco, 'dataPedido': dataPedido,
                    'dataEntrega': dataEntrega, 'condPgto': condPgto,
                    'itens': itens,
                })
    return filiais

# ════════════════════════════════════════════════
# ROTAS
# ════════════════════════════════════════════════
@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

@app.route('/processar', methods=['POST'])
def processar():
    try:
        perfil_file = request.files.get('perfil')
        pedido_file = request.files.get('pedido')
        cliente     = request.form.get('cliente', 'dom_atacarejo')

        if not perfil_file or not pedido_file:
            return jsonify({'erro': 'Envie perfil e pedido'}), 400

        # Ler Perfil Excel
        wb_p = openpyxl.load_workbook(io.BytesIO(perfil_file.read()), data_only=True)
        pws  = wb_p[wb_p.sheetnames[0]]
        pdata = list(pws.iter_rows(values_only=True))
        meta = {
            'empresa':  pdata[0][9],
            'codVend':  str(pdata[2][6]),
            'codCond':  str(pdata[3][6]),
            'vendedor': pdata[2][8],
            'telefone': pdata[2][9],
        }
        produtos = []
        for r in pdata[7:]:
            if not r[2]: break
            produtos.append({
                'codInterno':  r[1],
                'nomePerfil':  str(r[2]).strip(),
                'formato':     str(r[3] or '').strip(),
                'embalagem':   str(r[4]).strip(),
                'kgCx':        float(r[6] or 20),
                'unidFat':     str(r[7] or 'kg').strip(),
                'precoUnit':   float(r[8] or 0),
                'obs':         str(r[9] or '').strip(),
            })

        # Parser do pedido
        pdf_bytes = pedido_file.read()
        if cliente == 'dom_atacarejo':
            filiais = parse_dom_atacarejo(pdf_bytes, produtos)
        else:
            return jsonify({'erro': f'Cliente {cliente} não implementado'}), 400

        if not filiais:
            return jsonify({'erro': 'Nenhuma filial encontrada no pedido'}), 400

        # Montar dados para geradores
        dados = {**meta, 'filiais': filiais}

        # Gerar Excel e PDF
        excel_bytes = gerar_excel(dados)
        pdf_bytes_out = gerar_pdf(dados)

        return jsonify({
            'ok': True,
            'filiais': len(filiais),
            'itens': sum(len(f['itens']) for f in filiais),
            'totalKg': round(sum(i['kgPlanejados'] for f in filiais for i in f['itens']),1),
            'totalValor': round(sum(i['valorPedido'] for f in filiais for i in f['itens']),2),
            'resumo': [{'filial': f['filial'], 'pedidoNum': f['pedidoNum'],
                        'itens': len(f['itens']),
                        'kg': round(sum(i['kgPlanejados'] for i in f['itens']),1),
                        'valor': round(sum(i['valorPedido'] for i in f['itens']),2)}
                       for f in filiais],
            'excel': base64.b64encode(excel_bytes).decode(),
            'pdf':   base64.b64encode(pdf_bytes_out).decode(),
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'erro': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
