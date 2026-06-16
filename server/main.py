from flask import Flask, request, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import pdfplumber
import re, io, base64, os, traceback
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

CNPJ_DISTRIBUIDORA = '56.423.719'
CNPJ_INDUSTRIA     = '10.171.633'

# ════════════════════════════════════════════════
# PERFIS SALVOS NO SERVIDOR
# ════════════════════════════════════════════════
PERFIS_DIR = os.path.join(os.path.dirname(__file__), 'perfis')
os.makedirs(PERFIS_DIR, exist_ok=True)

def perfil_path(cliente):
    return os.path.join(PERFIS_DIR, f'{cliente}.xlsx')

def meta_path(cliente):
    return os.path.join(PERFIS_DIR, f'{cliente}_meta.json')

def perfil_existe(cliente):
    return os.path.exists(perfil_path(cliente))

def salvar_perfil(cliente, file_bytes, filename=None):
    with open(perfil_path(cliente), 'wb') as f:
        f.write(file_bytes)
    if filename:
        import json as _json
        with open(meta_path(cliente), 'w') as f:
            _json.dump({'filename': filename}, f)

def carregar_perfil_bytes(cliente):
    with open(perfil_path(cliente), 'rb') as f:
        return f.read()

def perfil_filename(cliente):
    import json as _json
    mp = meta_path(cliente)
    if os.path.exists(mp):
        return _json.load(open(mp)).get('filename','')
    return 

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
F_NT=fnt(sz=8,color='FF666666'); F_NR=fnt(sz=8,color='FFC0392B')
F_CX=fnt(True,9,'FF854F0B'); F_DIF=fnt(True,9,'FFCC0000')

FMT_MONEY='#,##0.00'; FMT_NUM='0.0'

def ap(cell, val=None, fn=None, bg=None, bo=None, al=None, fmt=None):
    if val is not None: cell.value = val
    if fn:  cell.font = fn
    if bg:  cell.fill = bg
    if bo:  cell.border = bo
    if al:  cell.alignment = al
    if fmt: cell.number_format = fmt

# ════════════════════════════════════════════════
# MATCH PRODUTO PDF → PERFIL
# ════════════════════════════════════════════════
def match_perfil(nome, produtos):
    n = (nome or '').lower().strip()
    best, score = None, 0
    for p in produtos:
        a = (p.get('nomePerfil','') or '').lower().strip()
        if not a: continue
        if a == n: s = 100
        elif n in a: s = 50 + len(a)
        elif a in n: s = 40 + len(n)
        else:
            common = len(set(a.split()) & set(n.split()))
            s = common * 10 if common >= 2 else 0
        if s > score: score, best = s, p
    return best

def detectar_empresa(txt):
    m = re.search(r'CNPJ\s+([\d.]+)', txt)
    if m and CNPJ_INDUSTRIA.replace('.','') in m.group(1).replace('.',''):
        return 1
    return 2

# ════════════════════════════════════════════════
# LER PERFIL EXCEL → meta + produtos
# ════════════════════════════════════════════════
def ler_perfil(perfil_bytes):
    wb_p  = openpyxl.load_workbook(io.BytesIO(perfil_bytes), data_only=True)
    pws   = wb_p[wb_p.sheetnames[0]]
    pdata = list(pws.iter_rows(values_only=True))
    meta  = {
        'empresa':  pdata[0][9] if pdata[0][9] else 2,
        'codVend':  str(pdata[2][6]) if pdata[2][6] else '',
        'codCond':  str(pdata[3][6]) if pdata[3][6] else '',
        'vendedor': pdata[2][8] or '',
        'telefone': pdata[2][9] or '',
    }
    produtos = []
    for r in pdata[7:]:
        if not r or not r[2]: break
        produtos.append({
            'empresa':     int(r[0]) if r[0] in (1,2) else None,  # coluna A: Fat. (1=Indústria, 2=Distrib.)
            'codInterno':  r[1],
            'nomePerfil':  str(r[2]).strip(),
            'formato':     str(r[3] or '').strip(),
            'embalagem':   str(r[4]).strip(),
            'kgCx':        float(r[6] or 20),
            'unidFat':     str(r[7] or 'kg').strip(),
            'precoUnit':   float(r[8] or 0),
            'obs':         str(r[9] or '').strip(),
        })
    return meta, produtos

# ════════════════════════════════════════════════
# GERADOR EXCEL
# ════════════════════════════════════════════════
def gerar_excel(dados, empresa_override=None):
    emp   = empresa_override if empresa_override else dados.get('empresa', 2)
    cv    = str(dados.get('codVend',''))
    cc    = str(dados.get('codCond',''))
    vend  = dados.get('vendedor','')
    tel   = dados.get('telefone','')
    tit   = 'PEDIDO PATA NEGRA DISTRIBUIDORA' if emp==2 else 'PEDIDO INDUSTRIA PATANEGRA'
    cli   = dados.get('clienteNome','')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for fd in dados['filiais']:
        # filtrar itens pela empresa quando em modo split
        if empresa_override:
            its = [it for it in fd['itens']
                   if (it.get('empresa') or empresa_override) == empresa_override]
        else:
            its = fd['itens']
        if not its: continue
        n = len(its)
        emp_fd = empresa_override if empresa_override else fd.get('empresa', emp)
        tit_fd = 'PEDIDO PATA NEGRA DISTRIBUIDORA' if emp_fd==2 else 'PEDIDO INDUSTRIA PATANEGRA'
        nota_empresa = 'Pata Negra Distribuidora' if emp_fd==2 else 'Indústria Pata Negra'
        numPed = re.sub(r'[/\\*?\[\]:]', '-', fd.get('pedidoNum','')).strip().strip('-')
        nomeAba = ((numPed+' - '+fd['filial']) if numPed else fd['filial'])[:31]
        ws = wb.create_sheet(nomeAba)

        for i,w in enumerate([4,10,38,12,7,12,12,9,14,2,11,12,14,12,10,8],1):
            ws.column_dimensions[get_column_letter(i)].width = w

        ws.merge_cells('A1:I1')
        ap(ws['A1'], tit_fd, F_TIT, BG_TIT, al=aln('center'))
        ws.merge_cells('A2:I2')
        subtitulo = (cli+' — '+fd['filial']) if cli else fd['filial']
        ap(ws['A2'], subtitulo, F_SUB, BG_SUB, al=aln('center'))

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
        ap(ws['H5'],fd.get('solicitante',''),F_M,BG_META,al=aln('left'))
        ap(ws['K5'],'Código vend.',F_HW,BG_AZL,BALL,aln('center',True))
        ap(ws['L5'],'código cond', F_HW,BG_AZL,BALL,aln('center',True))
        ap(ws['M5'],'empresa',     F_HW,BG_AZL,BALL,aln('center',True))

        ap(ws['A6'],'Endereço:',F_MB,BG_META,al=aln('left'))
        ap(ws['C6'],fd.get('endereco',''),F_M,BG_META)
        ap(ws['G6'],'Vendedor:',F_MB,BG_META,al=aln('left'))
        ap(ws['H6'],vend,F_M,BG_META,al=aln('left'))
        ws['K6'].value = int(cv) if cv.isdigit() else cv
        ws['L6'].value = int(cc) if cc.isdigit() else cc
        ws['M6'].value = emp_fd

        ap(ws['A7'],'Cond. Pgto.:',F_MB,BG_META,BBOT,aln('left'))
        ap(ws['C7'],fd.get('condPgto',''),F_M,BG_META,BBOT)
        for col in ['B7','D7','E7','F7']: ws[col].border = BBOT
        ap(ws['G7'],'Tel. Vendedor:',F_MB,BG_META,al=aln('left'))
        ap(ws['H7'],tel,F_M,BG_META,al=aln('left'))

        for col,h,az in [('A','#',0),('B','Cód.\nInterno',0),('C','Nome Produto\nno Cliente',0),
            ('D','Formato',0),('E','Caixa',0),('F','Kg\nPlanejados',0),('G','Kg\nEmbarcados',0),
            ('H','Nº\nCaixas',0),('I','Obs.',0),('K','Qtde\nMultipl.',1),('L','Preço Unit.\n(R$)',1),
            ('M','Valor Pedido\n(R$)',1),('N','Preço\nSistema',1),('O','Dif.\nPreço',1),('P','Unid.\nfat.',1)]:
            ap(ws[f'{col}8'],h,F_HW if az else F_HB,BG_AZL if az else BG_HDR,BALL,aln('center',True))
        ws.row_dimensions[8].height = 28

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
            if isCx: c('K').value = f'=IF(G{r}<>"",G{r}/{kgcx},F{r}/{kgcx})'
            else:     c('K').value = f'=IF(G{r}<>"",G{r},F{r})'
            ap(c('K'),None,F_IT,BG_AZLT,BALL,aln('center'))
            c('H').value = f'=IF(G{r}<>"",G{r}/{kgcx},F{r}/{kgcx})'
            ap(c('H'),None,F_IT,bg,BALL,aln('center'),FMT_NUM)
            ap(c('I'),it.get('obs') or None,F_IT,bg,BALL,aln('center'))
            ap(c('L'),it.get('precoUnit'),F_IT,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            c('M').value = f'=L{r}*K{r}' if isCx else f'=IF(G{r}<>"",L{r}*G{r},L{r}*F{r})'
            ap(c('M'),None,F_IT,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            ap(c('N'),it.get('precoSistema') or 0,F_IT,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            c('O').value = f'=IF(AND(L{r}<>"",N{r}<>"",L{r}<>N{r}),L{r}-N{r},"")'
            ap(c('O'),None,F_DIF,BG_AZLT,BALL,aln('center'),FMT_MONEY)
            ap(c('P'),it.get('unidFat','kg'),F_CX if isCx else F_IT,BG_AZLT,BALL,aln('center'))

        ws.conditional_formatting.add(
            f'O9:O{n+8}',
            FormulaRule(formula=[f'O9=""'], font=Font(bold=False,size=9,color='FF000000'))
        )

        rT=n+9; r1=9; r2=n+8
        ws.merge_cells(f'A{rT}:E{rT}')
        ap(ws[f'A{rT}'],f'TOTAL  —  {n} itens',F_TOT,None,brd(1,1,1,0),aln('left'))
        for col in ['B','C','D','E']: ws[f'{col}{rT}'].border = brd(1,1,0,0)
        for col in ['F','G','H']:
            ws[f'{col}{rT}'].value = f'=SUM({col}{r1}:{col}{r2})'
            ap(ws[f'{col}{rT}'],None,F_TOT,BG_HDR,BALL,aln('center'),FMT_NUM)
        ws[f'M{rT}'].value = f'=SUM(M{r1}:M{r2})'
        ap(ws[f'M{rT}'],None,F_TOT,BG_TOTV,BALL,aln('center'),FMT_MONEY)

        rI = rT+2
        ws.merge_cells(f'A{rI}:I{rI}')
        ws.merge_cells(f'K{rI}:P{rI}')
        ap(ws[f'A{rI}'],f'Imprimir: selecionar A1:I{rT-1} → Imprimir Seleção',F_NT,al=aln('left'))
        ap(ws[f'K{rI}'],'Col K=qtde faturamento | Col P=unidade | Dif. Preço vermelho=revisar',F_NR,al=aln('left'))

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
# GERADOR PDF
# ════════════════════════════════════════════════
def gerar_pdf(dados, empresa_override=None):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=8*mm, bottomMargin=8*mm)

    cli  = dados.get('clienteNome','')
    vend = dados.get('vendedor','')
    tel  = dados.get('telefone','')

    COR_TIT = colors.HexColor('#E0E0E0')
    COR_SUB = colors.HexColor('#F0F0F0')
    COR_META= colors.HexColor('#F8F8F8')
    COR_HDR = colors.HexColor('#D0D0D0')
    COR_PAR = colors.HexColor('#F5F5F5')
    COR_CZ  = colors.HexColor('#E8E8E8')

    col_w = [8*mm, 18*mm, 95*mm, 24*mm, 18*mm, 22*mm, 22*mm, 20*mm, 50*mm]
    W = sum(col_w)

    ST_TIT  = ParagraphStyle('t',  fontSize=13, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=16)
    ST_SUB  = ParagraphStyle('s',  fontSize=11, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=14)
    ST_MB   = ParagraphStyle('mb', fontSize=10, fontName='Helvetica-Bold', alignment=TA_LEFT,   leading=12)
    ST_MV   = ParagraphStyle('mv', fontSize=10, fontName='Helvetica',      alignment=TA_LEFT,   leading=12)
    ST_HDR  = ParagraphStyle('h',  fontSize=10, fontName='Helvetica-Bold', alignment=TA_CENTER, leading=11)
    ST_IT   = ParagraphStyle('i',  fontSize=10, fontName='Helvetica',      alignment=TA_LEFT,   leading=11)
    ST_ITC  = ParagraphStyle('ic', fontSize=10, fontName='Helvetica',      alignment=TA_CENTER, leading=11)
    ST_ITR  = ParagraphStyle('ir', fontSize=10, fontName='Helvetica',      alignment=TA_RIGHT,  leading=11)
    ST_TOT  = ParagraphStyle('to', fontSize=10, fontName='Helvetica-Bold', alignment=TA_LEFT,   leading=11)
    ST_TOTR = ParagraphStyle('tr', fontSize=10, fontName='Helvetica-Bold', alignment=TA_RIGHT,  leading=11)
    ST_NOTA = ParagraphStyle('n',  fontSize=8,  fontName='Helvetica',      textColor=colors.HexColor('#666666'))

    story = []

    for fd in dados['filiais']:
        if empresa_override:
            its = [it for it in fd['itens']
                   if (it.get('empresa') or empresa_override) == empresa_override]
        else:
            its = fd['itens']
        if not its: continue
        n = len(its)
        emp_fd = empresa_override if empresa_override else fd.get('empresa', dados.get('empresa', 2))
        tit_fd = 'PEDIDO PATA NEGRA DISTRIBUIDORA' if emp_fd==2 else 'PEDIDO INDUSTRIA PATANEGRA'
        tkg = sum(float(it.get('kgPlanejados',0)) for it in its)
        subtitulo = (cli+' — '+fd['filial']) if cli else fd['filial']

        tbl_tit = Table([[Paragraph(tit_fd, ST_TIT)]], colWidths=[W])
        tbl_tit.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),COR_TIT),
            ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))

        tbl_sub = Table([[Paragraph(subtitulo, ST_SUB)]], colWidths=[W])
        tbl_sub.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),COR_SUB),
            ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]))

        def ml(t): return Paragraph(t, ST_MB)
        def mv(t): return Paragraph(str(t) if t else '—', ST_MV)

        meta = [
            [ml('Pedido Nº:'), mv(fd.get('pedidoNum','')),  ml('Data Pedido:'),  mv(fd.get('dataPedido',''))],
            [ml('CNPJ:'),      mv(fd.get('cnpj','')),       ml('Data Entrega:'), mv(fd.get('dataEntrega',''))],
            [ml('Filial:'),    mv(fd['filial']),             ml('Solicitante:'),  mv('')],
            [ml('Endereço:'),  mv(fd.get('endereco','')),   ml('Vendedor:'),     mv(vend)],
            [ml('Cond. Pgto.:'), mv(fd.get('condPgto','')),ml('Tel. Vendedor:'), mv(tel)],
        ]
        tbl_meta = Table(meta, colWidths=[25*mm, 100*mm, 32*mm, W-25*mm-100*mm-32*mm])
        tbl_meta.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),COR_META),
            ('FONTSIZE',(0,0),(-1,-1),10),
            ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
            ('LEFTPADDING',(0,0),(-1,-1),4),
            ('LINEBELOW',(0,4),(-1,4),0.5,colors.grey),
        ]))

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
            kgPlan = it.get('kgPlanejados',0) or 0
            nrCx = round(kgPlan/kgcx,1) if kgcx else ''
            rows.append([
                Paragraph(str(idx+1), ST_ITC),
                Paragraph(str(it.get('codInterno','')), ST_ITC),
                Paragraph(str(it.get('nomeProduto','')), ST_IT),
                Paragraph(str(it.get('formato','') or ''), ST_ITC),
                Paragraph(str(it.get('embalagem','')), ST_ITC),
                Paragraph(f"{kgPlan:.1f}" if kgPlan else '', ST_ITR),
                Paragraph('', ST_ITC),
                Paragraph(str(nrCx) if nrCx else '', ST_ITR),
                Paragraph(str(it.get('obs','') or ''), ST_IT),
            ])

        rows.append([
            Paragraph(f'TOTAL — {n} itens', ST_TOT),'','','','',
            Paragraph(f'{tkg:.1f}', ST_TOTR),'','',''
        ])

        tbl_itens = Table(rows, colWidths=col_w, repeatRows=1)
        cmds = [
            ('BACKGROUND',(0,0),(-1,0),COR_HDR),
            ('FONTSIZE',(0,0),(-1,-1),10),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('GRID',(0,0),(-1,-2),0.5,colors.HexColor('#BBBBBB')),
            ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
            ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
            ('BACKGROUND',(6,1),(6,-2),COR_CZ),
            ('BACKGROUND',(0,-1),(-1,-1),COR_HDR),
            ('SPAN',(0,-1),(4,-1)),
            ('LINEABOVE',(0,-1),(-1,-1),0.8,colors.grey),
            ('LINEBELOW',(0,-1),(-1,-1),0.8,colors.grey),
        ]
        for i in range(1, len(rows)-1):
            if i % 2 == 0:
                cmds.append(('BACKGROUND',(0,i),(-1,i),COR_PAR))
        tbl_itens.setStyle(TableStyle(cmds))

        nota = Paragraph(
            f'Col. G — Kg Embarcados: preenchida pelo encarregado de embarque.   |   {nota_empresa}',
            ST_NOTA)

        story.append(KeepTogether([tbl_tit, tbl_sub, tbl_meta, tbl_itens, nota]))

    doc.build(story)
    buf.seek(0)
    return buf.read()

# ════════════════════════════════════════════════
# PARSER — FUNÇÃO AUXILIAR
# ════════════════════════════════════════════════
def processar_item(cod_cli, nome_raw, emb_tipo, qtde_emb, qtde_ped, preco, total, produtos):
    nome_raw = re.sub(r'\s+', ' ', nome_raw).strip()
    pf = match_perfil(nome_raw, produtos)
    kgCx     = pf['kgCx']      if pf else 20
    embalagem= pf['embalagem'] if pf else ('CX-'+str(qtde_emb) if emb_tipo in ['CX','CXA'] else 'CX-20')
    if emb_tipo in ['CX','CXA']:
        kgPlan=qtde_ped*kgCx; nrCx=qtde_ped; qtdeMult=qtde_ped; unidFat='cx'
    else:
        kgPlan=qtde_ped; nrCx=round(kgPlan/kgCx,1) if kgCx else 0; qtdeMult=kgPlan; unidFat='kg'
    return {
        'empresa':      pf.get('empresa') if pf else None,  # herda do perfil (coluna A)
        'codInterno':   pf['codInterno']    if pf else cod_cli,
        'nomeProduto':  pf['nomePerfil']    if pf else nome_raw,
        'formato':      pf.get('formato','') if pf else '',
        'embalagem':    embalagem, 'kgCx': kgCx,
        'kgPlanejados': kgPlan, 'nrCaixas': nrCx,
        'obs':          pf.get('obs','')    if pf else '',
        'qtdeMultipl':  qtdeMult, 'precoUnit': preco,
        'valorPedido':  total,
        'precoSistema': pf['precoUnit']     if pf else 0,
        'unidFat':      unidFat,
    }

# ════════════════════════════════════════════════
# PARSERS
# ════════════════════════════════════════════════
def parse_dom_atacarejo(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pags = len(pdf.pages)
        for pi in range(0, n_pags, 2):
            txt1 = pdf.pages[pi].extract_text() or ''
            txt2 = pdf.pages[pi+1].extract_text() if pi+1 < n_pags else ''
            lines = txt1.split('\n')
            txt_all = txt1+'\n'+txt2

            def fm(pat, txt=txt_all):
                m = re.search(pat, txt, re.I)
                return m.group(1).strip() if m else ''

            pedidoNum = fm(r'(\d{5,7}/[CL])')
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
            dataPedido  = fm(r'Data da emiss[aã]o\s+([\d/]+)')
            dataEntrega = fm(r'Previs[aã]o de entrega\s+([\d/]+)')
            condPgto    = fm(r'Prazo para pagamento\s+(\d+)')
            if condPgto: condPgto += ' dias'

            reItem = re.compile(r'^(\d{5,6})\s+\d+\s+(.+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)', re.M)
            itens = []
            for m in reItem.finditer(txt1):
                qtde_ped = float(m.group(5).replace('.','').replace(',','.'))
                preco    = float(m.group(6).replace('.','').replace(',','.'))
                total    = float(m.group(7).replace('.','').replace(',','.'))
                it = processar_item(int(m.group(1)), m.group(2), m.group(3),
                                    int(m.group(4)), qtde_ped, preco, total, produtos)
                itens.append(it)

            if filial and itens:
                filiais.append({'filial':filial,'pedidoNum':pedidoNum,'cnpj':cnpj,
                    'endereco':endereco,'dataPedido':dataPedido,'dataEntrega':dataEntrega,
                    'condPgto':condPgto,'empresa':2,'itens':itens})
    return filiais

def parse_atacadao(pdf_bytes, produtos):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)
    lines = txt.split('\n')

    def fm(pat):
        m = re.search(pat, txt, re.I)
        return m.group(1).strip() if m else ''

    pedidoNum = fm(r'Numero:\s*(\d+)')
    cnpj      = fm(r'Local de Entrega:\s*([\d./\-]+)')
    empresa   = 1 if CNPJ_INDUSTRIA.replace('.','') in cnpj.replace('.','').replace('-','') else 2

    cond_m = re.search(r'\|\s+(\d+)\s*\+?\s*\n.*?da data de recebimento', txt, re.S)
    condPgto = cond_m.group(1)+' dias' if cond_m else fm(r'Condicoes de Pagto.*?\n\|\s*(\d+)')
    if condPgto and 'dias' not in condPgto: condPgto += ' dias'

    dataEntrega = fm(r'RF\.[^|]+\s+(\d{2}/\d{2})\s+\d')
    if dataEntrega:
        p = dataEntrega.split('/')
        dataEntrega = f'{p[0]}/{p[1]}/2026' if len(p)==2 else dataEntrega

    endereco = ''
    for i,ln in enumerate(lines):
        if 'Local de Entrega:' in ln:
            for j in range(i+1,min(i+4,len(lines))):
                m2 = re.search(r'\|\s+([A-Z][^|]+?)\s+\|', lines[j])
                if m2 and len(m2.group(1).strip()) > 5:
                    endereco = m2.group(1).strip()
                    break
            break

    filial_num = re.search(r'/(\d{4})-', cnpj)
    filial = f'ATACADÃO FILIAL {int(filial_num.group(1))}' if filial_num else 'ATACADÃO'

    reItem = re.compile(r'\|\s+(RF\.\w[^|]+?)\s+(\d{2}/\d{2})\s+(\d+)\s+([\d,.]+)\s+0,.*?([\d,.]+)\s+S\s+\|')
    reRef  = re.compile(r'\|\s+(\d{8}/\d+)\s+(KG|GR|CXA)\s+(\S+)')

    itens = []; pending = None
    for ln in lines:
        mm2 = reItem.match(ln)
        if mm2:
            pending = {'nome': mm2.group(1).strip(),
                       'qtd':  float(mm2.group(3).replace('.','').replace(',','.')),
                       'preco':float(mm2.group(4).replace('.','').replace(',','.'))}
            continue
        m2 = reRef.search(ln)
        if m2 and pending:
            if m2.group(2) == 'GR': pending=None; continue
            emb_tipo = 'CX' if m2.group(2)=='CXA' else m2.group(2)
            qtde_emb = int(re.search(r'X(\d+)X', m2.group(3) or '1X20X1').group(1)) if re.search(r'X(\d+)X', m2.group(3) or '') else 20
            total = round(pending['qtd'] * pending['preco'], 2)
            it = processar_item(m2.group(1), pending['nome'], emb_tipo,
                                qtde_emb, pending['qtd'], pending['preco'], total, produtos)
            itens.append(it)
            pending = None

    if not itens: return []
    return [{'filial':filial,'pedidoNum':pedidoNum,'cnpj':cnpj,'endereco':endereco,
             'dataPedido':'','dataEntrega':dataEntrega,'condPgto':condPgto,
             'empresa':empresa,'itens':itens}]

def parse_assai(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ''

            def fm(pat):
                m = re.search(pat, txt, re.I)
                return m.group(1).strip() if m else ''

            pedidoNum   = fm(r'PEDIDO DE COMPRAS\s+(\S+)')
            cnpj_m      = re.search(r'CNPJ\s+([\d./\-]+)\s+Cidade.*?CNPJ\s+([\d./\-]+)', txt, re.S)
            cnpj_forn   = cnpj_m.group(1) if cnpj_m else ''
            cnpj_loja   = cnpj_m.group(2) if cnpj_m else ''
            empresa     = 1 if CNPJ_INDUSTRIA.replace('.','') in cnpj_forn.replace('.','') else 2
            filial_m    = re.search(r'R\. Social SENDAS.*?LJ\d+\s+\d+\s+(.+?)$', txt, re.M)
            filial      = filial_m.group(1).strip() if filial_m else 'ASSAÍ'
            end_m       = re.search(r'ENDEREÇO PARA ENTREGA.*?Endereço\s+(.+?)\s+Endereço', txt, re.S)
            endereco    = end_m.group(1).strip() if end_m else ''
            dataPedido  = fm(r'Data da emiss[aã]o\s+([\d/]+)')
            dataEntrega = fm(r'Previs[aã]o de entrega\s+([\d/]+)')
            cond_m2     = re.search(r'pagamento\s+(\d+)\s*\(', txt)
            condPgto    = cond_m2.group(1)+' dias' if cond_m2 else ''

            reItem = re.compile(
                r'^(\d{7})([A-Z][^\n]+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)',
                re.M)
            itens = []
            for m in reItem.finditer(txt):
                nome_raw = re.sub(r'\s+(FRAC\s*KG|KG)\s*$','', re.sub(r'\s+',' ', m.group(2))).strip()
                qtde_ped = float(m.group(5).replace('.','').replace(',','.'))
                preco    = float(m.group(6).replace('.','').replace(',','.'))
                total    = float(m.group(7).replace('.','').replace(',','.'))
                it = processar_item(m.group(1), nome_raw, m.group(3),
                                    int(m.group(4)), qtde_ped, preco, total, produtos)
                it['empresa'] = empresa
                itens.append(it)

            if itens:
                filiais.append({'filial':filial,'pedidoNum':pedidoNum,'cnpj':cnpj_loja,
                    'endereco':endereco,'dataPedido':dataPedido,'dataEntrega':dataEntrega,
                    'condPgto':condPgto,'empresa':empresa,'itens':itens})
    return filiais

def parse_torre_central(pdf_bytes, produtos):
    filiais = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pags = len(pdf.pages)
        for pi in range(0, n_pags, 2):
            txt1 = pdf.pages[pi].extract_text() or ''
            txt2 = pdf.pages[pi+1].extract_text() if pi+1 < n_pags else ''
            lines = txt1.split('\n')
            txt_all = txt1+'\n'+txt2

            def fm(pat, txt=txt_all):
                m = re.search(pat, txt, re.I)
                return m.group(1).strip() if m else ''

            pedidoNum = fm(r'Nº\s*([\d]+/[ML])')
            if not pedidoNum:
                pedidoNum = fm(r'PEDIDO DE COMPRAS\s*\n\s*Nº\s*([\d]+/[ML]?)')

            # filial: nome da loja no endereço
            filial = ''
            for ln in lines:
                if 'TORRE' in ln and 'CIA' in ln:
                    m2 = re.search(r'TORRE\s*&\s*CIA\s+SUPERMERCADOS\s+S/A\s+(.+)', ln)
                    if m2:
                        filial = m2.group(1).strip()
                        break
            if not filial:
                filial = fm(r'TORRE\s*&\s*CIA\s+SUPERMERCADOS\s+S/A\s+([A-Z\s]+?)(?:\s+AV\.|\s+RUA|\s+R\.)')
            if not filial: filial = 'TORRE'

            cnpj = ''
            for ln in lines:
                found = re.findall(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', ln)
                cnpjs_validos = [c for c in found if '56.423.719' not in c and '97.760.885' not in c.replace('07.760.885','')]
                if cnpjs_validos: cnpj = cnpjs_validos[0]; break
                elif found: cnpj = found[-1]; break

            endereco = fm(r'(?:AV\.|RUA|R\.)\s+[^–\n]+?–\s*([A-Z\s]+)\s+RIO DE JANEIRO')
            if not endereco:
                endereco = fm(r'ENDEREÇO PARA ENTREGA[:\s]+(.+?)(?:\n|ENDEREÇO PARA COBRANÇA)')

            dataPedido  = fm(r'Data Emiss[aã]o:\s*([\d/]+)')
            dataEntrega = fm(r'Previs[aã]o Entrega:\s*([\d/]+)')
            condPgto    = fm(r'Prazo para pagamento:\s*(\d+)')
            if condPgto: condPgto += ' dias'

            # empresa pelo CNPJ do fornecedor
            cnpj_forn = fm(r'CNPJ Fornecedor:\s*([\d./\-]+)')
            empresa = 1 if CNPJ_INDUSTRIA.replace('.','') in cnpj_forn.replace('.','').replace('-','') else 2

            reItem = re.compile(
                r'^(\d{5,6})\s+(\d{5,6})\s+(.+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)',
                re.M)
            # também tenta sem código duplo
            reItem2 = re.compile(
                r'^(\d{5,6})\s+(.+?)\s+(KG|CX)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)',
                re.M)

            itens = []
            for m in reItem.finditer(txt1):
                qtde_ped = float(m.group(6).replace('.','').replace(',','.'))
                preco    = float(m.group(7).replace('.','').replace(',','.'))
                total    = round(qtde_ped * preco, 2)
                it = processar_item(m.group(1), m.group(3), m.group(4),
                                    int(m.group(5)), qtde_ped, preco, total, produtos)
                itens.append(it)

            if not itens:
                for m in reItem2.finditer(txt1):
                    qtde_ped = float(m.group(5).replace('.','').replace(',','.'))
                    preco    = float(m.group(6).replace('.','').replace(',','.'))
                    total    = round(qtde_ped * preco, 2)
                    it = processar_item(m.group(1), m.group(2), m.group(3),
                                        int(m.group(4)), qtde_ped, preco, total, produtos)
                    itens.append(it)

            if itens:
                filiais.append({'filial':filial,'pedidoNum':pedidoNum,'cnpj':cnpj,
                    'endereco':endereco,'dataPedido':dataPedido,'dataEntrega':dataEntrega,
                    'condPgto':condPgto,'empresa':empresa,'itens':itens})
    return filiais

# ════════════════════════════════════════════════
# ROTAS
# ════════════════════════════════════════════════
@app.route('/health')
def health():
    perfis = {}
    for c in ['dom_atacarejo','atacadao','assai','torre_central']:
        if perfil_existe(c):
            perfis[c] = perfil_filename(c)
    return jsonify({'status':'ok', 'perfis': perfis})

@app.route('/perfil/<cliente>', methods=['POST'])
def upload_perfil(cliente):
    """Salva ou atualiza o perfil de um cliente no servidor."""
    clientes_validos = ['dom_atacarejo','atacadao','assai','torre_central']
    if cliente not in clientes_validos:
        return jsonify({'erro': f'Cliente inválido: {cliente}'}), 400
    f = request.files.get('perfil')
    if not f:
        return jsonify({'erro': 'Envie o arquivo perfil'}), 400
    filename = f.filename or ''
    salvar_perfil(cliente, f.read(), filename)
    return jsonify({'ok': True, 'cliente': cliente, 'filename': filename, 'mensagem': 'Perfil salvo com sucesso'})

@app.route('/logo/<cliente>')
def logo(cliente):
    """Retorna a logo extraída do perfil Excel do cliente."""
    if not perfil_existe(cliente):
        return jsonify({'erro': 'Perfil não encontrado'}), 404
    try:
        wb = openpyxl.load_workbook(perfil_path(cliente))
        ws = wb[wb.sheetnames[0]]
        if not ws._images:
            return jsonify({'erro': 'Sem imagem no perfil'}), 404
        img = ws._images[0]
        img.ref.seek(0)
        data = img.ref.read()
        return send_file(io.BytesIO(data), mimetype='image/png')
    except Exception as e:
        return jsonify({'erro': str(e)}), 500

@app.route('/processar', methods=['POST'])
def processar():
    try:
        perfil_file = request.files.get('perfil')
        pedido_file = request.files.get('pedido')
        cliente     = request.form.get('cliente','dom_atacarejo')

        if not pedido_file:
            return jsonify({'erro': 'Envie o pedido'}), 400

        # Perfil: usa o enviado agora (e salva) ou o salvo no servidor
        if perfil_file:
            perfil_bytes = perfil_file.read()
            salvar_perfil(cliente, perfil_bytes, perfil_file.filename)  # atualiza o salvo
        elif perfil_existe(cliente):
            perfil_bytes = carregar_perfil_bytes(cliente)
        else:
            return jsonify({'erro': f'Nenhum perfil disponível para {cliente}. Faça upload do perfil primeiro.'}), 400

        meta, produtos = ler_perfil(perfil_bytes)
        pdf_bytes = pedido_file.read()

        nomes_cliente = {
            'dom_atacarejo': 'DOM Atacarejo',
            'atacadao':      'Atacadão',
            'assai':         'Assaí',
            'torre_central': 'Torre Central',
        }
        if cliente == 'dom_atacarejo':
            filiais = parse_dom_atacarejo(pdf_bytes, produtos)
        elif cliente == 'atacadao':
            filiais = parse_atacadao(pdf_bytes, produtos)
        elif cliente == 'assai':
            filiais = parse_assai(pdf_bytes, produtos)
            if filiais: meta['empresa'] = filiais[0]['empresa']
        elif cliente == 'torre_central':
            filiais = parse_torre_central(pdf_bytes, produtos)
        else:
            return jsonify({'erro': f'Cliente {cliente} não implementado'}), 400

        if not filiais:
            return jsonify({'erro': 'Nenhuma filial encontrada no pedido'}), 400

        dados = {**meta, 'filiais': filiais, 'clienteNome': nomes_cliente.get(cliente, cliente)}

        # detectar se há itens de empresas diferentes (split)
        empresas_nos_itens = set(
            it.get('empresa') or dados.get('empresa', 2)
            for f in filiais for it in f['itens']
        )
        empresas_nos_itens.discard(None)
        if not empresas_nos_itens: empresas_nos_itens = {dados.get('empresa', 2)}

        arquivos = []
        for emp_split in sorted(empresas_nos_itens):
            eb = gerar_excel(dados, empresa_override=emp_split if len(empresas_nos_itens)>1 else None)
            pb = gerar_pdf(dados,   empresa_override=emp_split if len(empresas_nos_itens)>1 else None)
            label = ('Indústria' if emp_split==1 else 'Distribuidora') if len(empresas_nos_itens)>1 else ''
            arquivos.append({
                'empresa': emp_split,
                'label':   label,
                'excel':   base64.b64encode(eb).decode(),
                'pdf':     base64.b64encode(pb).decode(),
            })

        todos_itens = [i for f in filiais for i in f['itens']]
        return jsonify({
            'ok':         True,
            'split':      len(empresas_nos_itens) > 1,
            'filiais':    len(filiais),
            'itens':      len(todos_itens),
            'totalKg':    round(sum(i['kgPlanejados'] for i in todos_itens),1),
            'totalValor': round(sum(i['valorPedido']  for i in todos_itens),2),
            'resumo': [{'filial':f['filial'],'pedidoNum':f.get('pedidoNum',''),
                        'itens':len(f['itens']),
                        'kg':   round(sum(i['kgPlanejados'] for i in f['itens']),1),
                        'valor':round(sum(i['valorPedido']  for i in f['itens']),2)}
                       for f in filiais],
            'arquivos': arquivos,
            # compatibilidade retroativa (caso simples)
            'excel': base64.b64encode(gerar_excel(dados)).decode() if len(empresas_nos_itens)==1 else '',
            'pdf':   base64.b64encode(gerar_pdf(dados)).decode()   if len(empresas_nos_itens)==1 else '',
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'erro': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
