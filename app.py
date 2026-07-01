import streamlit as st
import re
import io

# ─── CONFIGURAÇÃO DA PÁGINA ────────────────────────────────────────
st.set_page_config(
    page_title="Gerador Batch — PRODAM",
    page_icon="📊",
    layout="wide",
)

# ─── CSS CUSTOMIZADO ───────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { max-width: 1200px; padding-top: 1rem; }
    .header-bar {
        background: #1F4E79; color: #A8C8E8; padding: 12px 20px;
        border-radius: 6px; margin-bottom: 16px;
        display: flex; align-items: center; justify-content: space-between;
    }
    .header-bar h2 { color: #fff; margin: 0; font-size: 18px; }
    .header-bar .badge { background: #2E75B6; color: #A8C8E8; font-size: 11px; padding: 3px 10px; border-radius: 12px; }
    .success-box { background: #E8F5E9; border-left: 3px solid #1B5E20; padding: 10px 14px; border-radius: 4px; margin: 8px 0; }
    .info-box { background: #E3F0FA; border-left: 3px solid #1565A0; padding: 10px 14px; border-radius: 4px; margin: 8px 0; }
    .warn-box { background: #FFF3E0; border-left: 3px solid #E65100; padding: 10px 14px; border-radius: 4px; margin: 8px 0; }
</style>
""", unsafe_allow_html=True)

# ─── CABEÇALHO ─────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <div>
        <h2>■ PRODAM — Gerador de Documentação Batch</h2>
        <div style="font-size:12px;color:#7AABCC;margin-top:2px">Upload do(s) script(s) SQL → geração automática da Característica (.docx) e do Fluxograma (SVG)</div>
    </div>
    <span class="badge">SIGA SAÚDE · V3.0</span>
</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
# PARSER DE SQL
# ═══════════════════════════════════════════════════════════════════

def extrair_dados_sql(texto: str, nome_arquivo: str = '') -> dict:
    """Extrai informações do script SQL."""
    dados = {
        'tabelas': [],
        'arquivo_saida': '',
        'nome_script': nome_arquivo,
        'tem_select': False,
        'tem_insert': False,
        'tem_update': False,
        'tem_delete': False,
    }

    # Tabelas (FROM e JOIN) — exclui DUAL
    tabelas = set()
    for m in re.finditer(r'(?:FROM|JOIN)\s+(TB_\w+|VW_\w+)', texto, re.IGNORECASE):
        tabelas.add(m.group(1).upper())
    dados['tabelas'] = sorted(tabelas)

    # Arquivo de saída (SPOOL) — ignora SPOOL ON/OFF/OUT
    for m in re.finditer(r'SPOOL\s+(\S+)', texto, re.IGNORECASE):
        val = m.group(1).strip()
        if val.upper() not in ('ON', 'OFF', 'OUT'):
            dados['arquivo_saida'] = val
            break

    # Tipo de operação
    dados['tem_select'] = bool(re.search(r'\bSELECT\b', texto, re.IGNORECASE))
    dados['tem_insert'] = bool(re.search(r'\bINSERT\b', texto, re.IGNORECASE))
    dados['tem_update'] = bool(re.search(r'\bUPDATE\b', texto, re.IGNORECASE))
    dados['tem_delete'] = bool(re.search(r'\bDELETE\b', texto, re.IGNORECASE))

    return dados


def descrever_operacao_sql(dados_sql: dict) -> str:
    """Gera descrição textual do que o script faz."""
    ops = []
    if dados_sql['tem_select']:
        ops.append('consulta (SELECT)')
    if dados_sql['tem_insert']:
        ops.append('inserção (INSERT)')
    if dados_sql['tem_update']:
        ops.append('atualização (UPDATE)')
    if dados_sql['tem_delete']:
        ops.append('exclusão (DELETE)')
    if ops:
        return 'Script faz ' + ', '.join(ops) + ' no banco.'
    return 'Script SQL.'


# ═══════════════════════════════════════════════════════════════════
# GERADOR DE CARACTERÍSTICA (.docx)
# ═══════════════════════════════════════════════════════════════════

def gerar_caracteristica_docx(dados: dict) -> bytes:
    """Gera o documento Característica no formato PRODAM usando python-docx."""
    from docx import Document
    from docx.shared import Pt, Inches, Cm, RGBColor, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml

    doc = Document()

    # Configurar margens
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    BLU = '1F4E79'
    WHT = 'FFFFFF'
    LBL = '9BB8D8'

    cod = dados.get('cod', '')
    sis = dados.get('sis', 'SIGA SAUDE')
    sub = dados.get('sub', '')
    den = dados.get('den', '')
    amb = dados.get('amb', 'UNIX')
    dat = dados.get('dat', '')
    hora = dados.get('hora', '00h00')
    etapas = dados.get('etapas', [])
    total_folhas = len(etapas) + 1  # etapa 00 + etapas

    def set_cell_bg(cell, color):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}" w:val="clear"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_text(cell, text, size=9, bold=False, color=WHT, align=WD_ALIGN_PARAGRAPH.LEFT):
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = align
        run = p.add_run(str(text))
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = 'Arial Narrow'
        run.font.color.rgb = RGBColor.from_string(color)
        # Reduzir espaçamento
        pf = p.paragraph_format
        pf.space_before = Pt(1)
        pf.space_after = Pt(1)

    def add_label_value(cell, label, value, size=9):
        """Adiciona label pequeno + valor em negrito na mesma célula."""
        cell.text = ''
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run_lbl = p.add_run(label + '\n')
        run_lbl.font.size = Pt(7)
        run_lbl.font.name = 'Arial Narrow'
        run_lbl.font.color.rgb = RGBColor.from_string(LBL)
        run_val = p.add_run(str(value))
        run_val.font.size = Pt(size)
        run_val.font.bold = True
        run_val.font.name = 'Arial Narrow'
        run_val.font.color.rgb = RGBColor.from_string(WHT)

    def criar_pagina_etapa(etapa_num, etapa_data, folha_num):
        """Cria uma página de Característica para uma etapa."""

        # ─── CABEÇALHO PRINCIPAL ────────────────────────────
        tbl = doc.add_table(rows=3, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Larguras das colunas
        for row in tbl.rows:
            row.cells[0].width = Cm(3)
            row.cells[1].width = Cm(10)
            row.cells[2].width = Cm(5)

        # Linha 1: PRODAM | CARACTERÍSTICA... | CÓDIGO/ETAPA
        for cell in tbl.rows[0].cells:
            set_cell_bg(cell, BLU)
        set_cell_text(tbl.rows[0].cells[0], 'PRODAM', 14, True, WHT, WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(tbl.rows[0].cells[1], 'CARACTERÍSTICA DE PROGRAMA/UTILITÁRIO', 11, True, WHT, WD_ALIGN_PARAGRAPH.CENTER)
        add_label_value(tbl.rows[0].cells[2], 'CÓDIGO DA ROTINA', f'{cod}\nETAPA {str(etapa_num).zfill(2)}', 11)

        # Linha 2: SISTEMA | SUBSISTEMA | CÓD. PROGRAMA
        for cell in tbl.rows[1].cells:
            set_cell_bg(cell, BLU)
        add_label_value(tbl.rows[1].cells[0], 'SISTEMA', sis, 9)
        add_label_value(tbl.rows[1].cells[1], 'SUBSISTEMA', sub, 9)
        add_label_value(tbl.rows[1].cells[2], 'CÓD. PROGRAMA', 'SQLPLUS', 9)

        # Linha 3: DENOMINAÇÃO | DATA/FOLHA
        for cell in tbl.rows[2].cells:
            set_cell_bg(cell, BLU)
        # Mesclar as 2 primeiras colunas da linha 3
        tbl.rows[2].cells[0].merge(tbl.rows[2].cells[1])
        add_label_value(tbl.rows[2].cells[0], 'DENOMINAÇÃO DO PROGRAMA', den, 9)
        add_label_value(tbl.rows[2].cells[2], 'DATA DA ELABORAÇÃO',
                        f'{dat}    FOLHA: {str(folha_num).zfill(2)}/{str(total_folhas).zfill(2)}', 9)

        # Bordas brancas
        for row in tbl.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                borders = parse_xml(
                    f'<w:tcBorders {nsdecls("w")}>'
                    f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{WHT}"/>'
                    f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{WHT}"/>'
                    f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{WHT}"/>'
                    f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{WHT}"/>'
                    f'</w:tcBorders>'
                )
                tcPr.append(borders)

        # ─── ÁREA DE CONTEÚDO ──────────────────────────────
        doc.add_paragraph('')  # espaço

        if etapa_num == 0:
            # ETAPA 00 — descrição geral
            p = doc.add_paragraph()
            run = p.add_run(f'Rotina que inicialmente será executada diariamente, às {hora} da manhã.')
            run.font.size = Pt(10)
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')

            p = doc.add_paragraph()
            run = p.add_run('OBJETIVO:')
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')

            objetivo = etapa_data.get('objetivo', dados.get('den', ''))
            p = doc.add_paragraph()
            run = p.add_run(f'Automatizar a disponibilização do arquivo na pasta FTP (Externo):')
            run.font.size = Pt(10)
            run.font.name = 'Arial Narrow'

            # SFTP details
            if dados.get('ip'):
                doc.add_paragraph('')
                sftp_lines = [
                    f'Protocolo: {dados.get("proto", "SFTP")}',
                    f'Host: IP {dados.get("ip", "")}',
                    f'porta: {dados.get("porta", "")} user: {dados.get("user", "")} Senha: {dados.get("senha", "")}',
                    f'Pasta: {dados.get("pasta", "")}',
                ]
                for line in sftp_lines:
                    p = doc.add_paragraph()
                    run = p.add_run(line)
                    run.font.size = Pt(10)
                    run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            p = doc.add_paragraph()
            run = p.add_run('IMPORTANTE:')
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            p = doc.add_paragraph()
            run = p.add_run('COMANDOS NECESSÁRIOS ANTES DE INICIAR O SQLPLUS:')
            run.font.size = Pt(10)
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            for cmd in ['export LANG=pt_BR.ISO-8859-1', 'export NLS_LANG=BRAZILIAN PORTUGUESE_BRAZIL.WE8MSWIN1252']:
                p = doc.add_paragraph()
                run = p.add_run(cmd)
                run.font.size = Pt(10)
                run.font.name = 'Arial Narrow'

        elif etapa_data.get('tipo') == 'SQLPLUS':
            p = doc.add_paragraph()
            run = p.add_run('EXECUTAR SCRIPT SQL:')
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            p = doc.add_paragraph()
            run = p.add_run(etapa_data.get('inp', ''))
            run.font.size = Pt(10)
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            obs = etapa_data.get('obs', '')
            if obs:
                p = doc.add_paragraph()
                run = p.add_run(f'Observação: {obs}')
                run.font.size = Pt(10)
                run.font.name = 'Arial Narrow'

            if etapa_data.get('out'):
                doc.add_paragraph('')
                doc.add_paragraph('')
                p = doc.add_paragraph()
                run = p.add_run('ARQUIVO DE SAÍDA GERADO PELO SCRIPT SQL:')
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.name = 'Arial Narrow'

                doc.add_paragraph('')
                p = doc.add_paragraph()
                run = p.add_run(etapa_data.get('out', ''))
                run.font.size = Pt(10)
                run.font.name = 'Arial Narrow'

                # Se o arquivo tem variável &x, explicar
                if '&' in etapa_data.get('out', ''):
                    doc.add_paragraph('')
                    p = doc.add_paragraph()
                    run = p.add_run("Onde: &x = 'DD-MM-YYYY'")
                    run.font.size = Pt(10)
                    run.font.name = 'Arial Narrow'

        elif etapa_data.get('tipo') == 'COMPACTAR':
            p = doc.add_paragraph()
            run = p.add_run('COMPACTAR ARQUIVO:')
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            p = doc.add_paragraph()
            run = p.add_run(etapa_data.get('out', etapa_data.get('inp', '') + '.zip'))
            run.font.size = Pt(10)
            run.font.name = 'Arial Narrow'

        elif etapa_data.get('tipo') == 'TRANSFERIR':
            p = doc.add_paragraph()
            run = p.add_run('TRANSFERIR ARQUIVO:')
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            p = doc.add_paragraph()
            run = p.add_run(etapa_data.get('inp', ''))
            run.font.size = Pt(10)
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            doc.add_paragraph('')
            p = doc.add_paragraph()
            run = p.add_run('Disponibilizar os arquivos compactados, na pasta (CP):')
            run.font.size = Pt(10)
            run.font.name = 'Arial Narrow'

            doc.add_paragraph('')
            p = doc.add_paragraph()
            run = p.add_run('Destino do arquivo gerado (servidor SFTP):')
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.name = 'Arial Narrow'

            sftp_items = [
                f'- IP: {dados.get("ip", "")}',
                f'- Porta: {dados.get("porta", "")}',
                f'- Protocolo: {dados.get("proto", "SFTP")}',
                f'- Pasta: {dados.get("pasta", "")}',
                'Credenciais:',
                f'- Usuário: {dados.get("user", "")}',
                f'- Senha: {dados.get("senha", "")}',
            ]
            for item in sftp_items:
                p = doc.add_paragraph()
                run = p.add_run(item)
                run.font.size = Pt(10)
                run.font.name = 'Arial Narrow'

        # ─── RODAPÉ PADRÃO (campos de classificação + relatórios) ─
        doc.add_paragraph('')

        def set_table_borders(tbl):
            """Aplica bordas pretas em TODAS as células da tabela."""
            bdr_xml = (
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            )
            for row in tbl.rows:
                for cell in row.cells:
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(bdr_xml))

        def set_cell_text_black(cell, text, size=7, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
            """Texto preto para as tabelas de rodapé."""
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = align
            run = p.add_run(str(text))
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.name = 'Arial Narrow'
            run.font.color.rgb = RGBColor.from_string('000000')
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)

        # Tabela "CAMPOS DE CLASSIFICAÇÃO" — 12 colunas espelhadas + 3 linhas vazias
        NCOLS_CLASS = 12
        NROWS_EMPTY = 1
        tbl_class = doc.add_table(rows=2 + NROWS_EMPTY, cols=NCOLS_CLASS)
        tbl_class.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Linha título — mescla todas as colunas
        for ci in range(1, NCOLS_CLASS):
            tbl_class.rows[0].cells[0].merge(tbl_class.rows[0].cells[ci])
        set_cell_text_black(tbl_class.rows[0].cells[0], 'CAMPOS DE CLASSIFICAÇÃO', 9, True, WD_ALIGN_PARAGRAPH.CENTER)

        # Linha cabeçalho — espelhada
        headers_class = ['NUM', 'NOME DO CAMPO', 'POS.REL.', 'DIM', 'FORM', 'SEQ.',
                         'NUM', 'NOME DO CAMPO', 'POS.REL.', 'DIM', 'FORM', 'SEQ.']
        for i, h in enumerate(headers_class):
            set_cell_text_black(tbl_class.rows[1].cells[i], h, 7, False, WD_ALIGN_PARAGRAPH.LEFT)

        # Linhas vazias (para preenchimento manual)
        for r in range(2, 2 + NROWS_EMPTY):
            for c in range(NCOLS_CLASS):
                set_cell_text_black(tbl_class.rows[r].cells[c], '', 7)

        set_table_borders(tbl_class)

        doc.add_paragraph('')

        # Tabela "RELATÓRIOS" — 9 colunas + 3 linhas vazias
        NCOLS_REL = 9
        tbl_rel = doc.add_table(rows=2 + NROWS_EMPTY, cols=NCOLS_REL)
        tbl_rel.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Linha título — mescla todas as colunas
        for ci in range(1, NCOLS_REL):
            tbl_rel.rows[0].cells[0].merge(tbl_rel.rows[0].cells[ci])
        set_cell_text_black(tbl_rel.rows[0].cells[0], 'RELATÓRIOS', 9, True, WD_ALIGN_PARAGRAPH.CENTER)

        # Linha cabeçalho
        headers_rel = ['DSNAME/ETAPA', 'CS', 'SQ', 'CAMPO DE CONTROLE', 'CADEIA ESP.',
                       'CÓD. FORMULÁRIO', 'CÓPIAS', 'IMPRESSORA', 'FCB']
        for i, h in enumerate(headers_rel):
            set_cell_text_black(tbl_rel.rows[1].cells[i], h, 7, False, WD_ALIGN_PARAGRAPH.LEFT)

        # Linhas vazias
        for r in range(2, 2 + NROWS_EMPTY):
            for c in range(NCOLS_REL):
                set_cell_text_black(tbl_rel.rows[r].cells[c], '', 7)

        set_table_borders(tbl_rel)

        # Quebra de página (exceto na última etapa)
        if etapa_num < len(etapas):
            doc.add_page_break()

    # ─── GERAR TODAS AS PÁGINAS ─────────────────────────────
    # Etapa 00 — descrição geral
    criar_pagina_etapa(0, {}, 1)

    # Etapas 01, 02, 03...
    for i, et in enumerate(etapas):
        criar_pagina_etapa(et['num'], et, i + 2)

    # Salvar em bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════
# GERADOR DE SVG (FLUXOGRAMA) — mantém lógica existente
# ═══════════════════════════════════════════════════════════════════

def escape_svg(s):
    return str(s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def wrap_text(text, max_chars):
    words = str(text or '').split(' ')
    lines, cur = [], ''
    for w in words:
        if len((cur + ' ' + w).strip()) > max_chars:
            if cur:
                lines.append(cur)
            cur = w
        else:
            cur = (cur + ' ' + w).strip()
    if cur:
        lines.append(cur)
    return lines


def gerar_svg(dados: dict) -> list:
    """Gera lista de SVGs (um por folha) com paginação automática."""
    cod = dados.get('cod', '???')
    sis = dados.get('sis', '')
    sub = dados.get('sub', '')
    den = dados.get('den', '')
    amb = dados.get('amb', 'UNIX')
    dat = dados.get('dat', '')
    hora = dados.get('hora', '??h??')
    ip = dados.get('ip', '')
    porta = dados.get('porta', '')
    pasta = dados.get('pasta', '')
    user = dados.get('user', '')
    etapas = dados.get('etapas', [])
    tabelas = dados.get('tabelas', [])

    W = 700
    BLU, GRY, LGY, BRD = '#1F4E79', '#333', '#888', '#333'
    FNT = 'Arial Narrow,Arial,sans-serif'
    ex = escape_svg

    # ─── Estimativa de altura para decidir paginação ──────────
    # Altura max do SVG que cabe numa A4 no Word (27.7cm úteis × 700/18)
    MAX_SVG_H = 1077
    HEADER_H  = 116   # header PRODAM (96px) + gap (20px)
    FOOTER_H  = 34    # linha + texto + margem
    FIM_H     = 74    # seta + pill FIM
    SFTP_H    = 162   # seta + cilindro SFTP
    TABELAS_H = 16    # linha de tabelas Oracle
    # Por etapa: seta(28) + oval_in(~58) + box(56) + oval_out(~58) + gap(10)
    ET_H      = 210   # estimativa conservadora

    tem_sftp = bool(ip)
    n_et = len(etapas)

    # Estima altura total se tudo ficasse em 1 folha
    est_total = HEADER_H + (ET_H * n_et) + FIM_H + FOOTER_H + TABELAS_H
    if tem_sftp:
        est_total += SFTP_H

    # Se cabe em 1 página A4 → tudo numa folha; senão → max 2 por folha
    MAX_ET = n_et if (n_et > 0 and est_total <= MAX_SVG_H) else 2

    # Agrupa etapas em folhas
    folhas_etapas = []
    for i in range(0, len(etapas), MAX_ET):
        folhas_etapas.append(etapas[i:i + MAX_ET])
    if not folhas_etapas:
        folhas_etapas = [[]]

    tem_sftp = bool(ip)
    ultima_cheia = len(folhas_etapas[-1]) >= MAX_ET
    total_folhas = len(folhas_etapas) + (1 if tem_sftp and ultima_cheia else 0)

    def draw_header(els, Y, fnum):
        C1, C2, C3 = 120, 380, 200
        RH1, RH2, RH3 = 36, 24, 36
        HH = RH1 + RH2 + RH3
        bw = 1.5
        els.append(f'<rect x="0" y="{Y}" width="{W}" height="{HH}" fill="{BLU}"/>')
        els.append(f'<rect x="0" y="{Y}" width="{W}" height="{HH}" fill="none" stroke="#fff" stroke-width="{bw}"/>')
        els.append(f'<line x1="{C1}" y1="{Y}" x2="{C1}" y2="{Y+RH1}" stroke="#fff" stroke-width="{bw}"/>')
        els.append(f'<line x1="{C1+C2}" y1="{Y}" x2="{C1+C2}" y2="{Y+RH1}" stroke="#fff" stroke-width="{bw}"/>')
        els.append(f'<line x1="0" y1="{Y+RH1}" x2="{W}" y2="{Y+RH1}" stroke="#fff" stroke-width="{bw}"/>')
        els.append(f'<text x="{C1//2}" y="{Y+RH1//2+5}" font-size="14" font-weight="bold" text-anchor="middle" fill="#fff" font-family="{FNT}">PRODAM</text>')
        els.append(f'<text x="{C1+C2//2}" y="{Y+RH1//2+5}" font-size="12" font-weight="bold" text-anchor="middle" fill="#fff" font-family="{FNT}">FLUXOGRAMA DE ROTINA DE OPERAÇÃO</text>')
        els.append(f'<text x="{C1+C2+C3//2}" y="{Y+RH1//2-2}" font-size="8" text-anchor="middle" fill="#9BB8D8" font-family="{FNT}">CÓDIGO DA ROTINA:</text>')
        els.append(f'<text x="{C1+C2+C3//2}" y="{Y+RH1//2+12}" font-size="12" font-weight="bold" text-anchor="middle" fill="#fff" font-family="{FNT}">{ex(cod)} ({ex(amb)})</text>')
        y2 = Y + RH1
        els.append(f'<line x1="{W//2}" y1="{y2}" x2="{W//2}" y2="{y2+RH2}" stroke="#fff" stroke-width="{bw}"/>')
        els.append(f'<line x1="0" y1="{y2+RH2}" x2="{W}" y2="{y2+RH2}" stroke="#fff" stroke-width="{bw}"/>')
        els.append(f'<text x="10" y="{y2+RH2//2+4}" font-size="10" fill="#fff" font-family="{FNT}">SISTEMA: <tspan font-weight="bold">{ex(sis)}</tspan></text>')
        els.append(f'<text x="{W//2+10}" y="{y2+RH2//2+4}" font-size="10" fill="#fff" font-family="{FNT}">SUBSISTEMA: <tspan font-weight="bold">{ex(sub)}</tspan></text>')
        y3 = y2 + RH2
        els.append(f'<line x1="{W-C3}" y1="{y3}" x2="{W-C3}" y2="{y3+RH3}" stroke="#fff" stroke-width="{bw}"/>')
        els.append(f'<text x="10" y="{y3+12}" font-size="8" fill="#9BB8D8" font-family="{FNT}">DENOMINAÇÃO DO PROGRAMA</text>')
        els.append(f'<text x="10" y="{y3+27}" font-size="10" font-weight="bold" fill="#fff" font-family="{FNT}">{ex(den)}</text>')
        els.append(f'<text x="{W-C3+10}" y="{y3+12}" font-size="8" fill="#9BB8D8" font-family="{FNT}">DATA DA ELABORAÇÃO:</text>')
        els.append(f'<text x="{W-C3+10}" y="{y3+27}" font-size="10" font-weight="bold" fill="#fff" font-family="{FNT}">{ex(dat)}    FOLHA: {str(fnum).zfill(2)}/{str(total_folhas).zfill(2)}</text>')
        return Y + HH + 20

    def draw_oval(els, Y, nome):
        fw = min(W - 100, len(nome) * 7 + 50)
        fh = 30
        els.append(f'<ellipse cx="{W//2}" cy="{Y+fh//2}" rx="{fw//2}" ry="{fh//2}" fill="none" stroke="{BRD}" stroke-width="1.2"/>')
        els.append(f'<text x="{W//2}" y="{Y+fh//2+4}" font-size="10" text-anchor="middle" fill="{GRY}" font-family="{FNT}">{ex(nome)}</text>')
        return Y + fh

    def draw_arrow(els, Y, length=24):
        mid = Y + length - 8
        els.append(f'<line x1="{W//2}" y1="{Y}" x2="{W//2}" y2="{mid}" stroke="{GRY}" stroke-width="1.2"/>')
        els.append(f'<polygon points="{W//2},{Y+length} {W//2-5},{mid} {W//2+5},{mid}" fill="{GRY}"/>')
        return Y + length

    def draw_footer(els, Y, fnum):
        els.append(f'<line x1="0" y1="{Y}" x2="{W}" y2="{Y}" stroke="{LGY}" stroke-width="0.5"/>')
        els.append(f'<text x="10" y="{Y+12}" font-size="8" fill="{LGY}" font-family="{FNT}">PRODAM / {ex(sis)} / {ex(cod)}</text>')
        els.append(f'<text x="{W-10}" y="{Y+12}" font-size="8" text-anchor="end" fill="{LGY}" font-family="{FNT}">FOLHA {str(fnum).zfill(2)}/{str(total_folhas).zfill(2)}</text>')
        return Y + 20

    def draw_sftp(els, Y):
        CW_s, CRH_s = 260, 120
        CX_s = (W - CW_s) // 2
        CR = 16
        els.append(f'<ellipse cx="{CX_s+CW_s//2}" cy="{Y+CR}" rx="{CW_s//2}" ry="{CR}" fill="none" stroke="{BRD}" stroke-width="1.5"/>')
        els.append(f'<rect x="{CX_s}" y="{Y+CR}" width="{CW_s}" height="{CRH_s-CR*2}" fill="#fff" stroke="{BRD}" stroke-width="1.5"/>')
        els.append(f'<ellipse cx="{CX_s+CW_s//2}" cy="{Y+CRH_s-CR}" rx="{CW_s//2}" ry="{CR}" fill="none" stroke="{BRD}" stroke-width="1.5"/>')
        els.append(f'<text x="{CX_s+CW_s//2}" y="{Y+38}" font-size="12" font-weight="bold" text-anchor="middle" fill="{GRY}" font-family="{FNT}">SFTP — Servidor Externo</text>')
        els.append(f'<text x="{CX_s+CW_s//2}" y="{Y+55}" font-size="10" text-anchor="middle" fill="{GRY}" font-family="{FNT}">IP: {ex(ip)}  |  Porta: {ex(porta)}</text>')
        els.append(f'<text x="{CX_s+CW_s//2}" y="{Y+70}" font-size="10" text-anchor="middle" fill="{GRY}" font-family="{FNT}">Usuário: {ex(user)}</text>')
        els.append(f'<text x="{CX_s+CW_s//2}" y="{Y+85}" font-size="9" text-anchor="middle" fill="{LGY}" font-family="{FNT}">{ex(pasta)}</text>')
        return Y + CRH_s + 10

    def draw_fim(els, Y):
        Y = draw_arrow(els, Y)
        Y += 4
        FW, FX = 160, (W - 160) // 2
        els.append(f'<rect x="{FX}" y="{Y}" width="{FW}" height="32" rx="16" fill="{BLU}"/>')
        els.append(f'<text x="{W//2}" y="{Y+21}" font-size="12" font-weight="bold" text-anchor="middle" fill="#fff" font-family="{FNT}">FIM</text>')
        return Y + 46

    def make_svg(els, Y):
        return f'<svg viewBox="0 0 {W} {Y}" xmlns="http://www.w3.org/2000/svg" width="{W}" style="background:#fff;max-width:100%">\n' + '\n'.join(els) + '\n</svg>'

    # ─── GERA FOLHAS ─────────────────────────────────────────────
    svgs = []
    prev_out = None

    for fi, folha_ets in enumerate(folhas_etapas):
        fnum = fi + 1
        els = []
        Y = draw_header(els, 0, fnum)

        for et in folha_ets:
            EW, EX = 200, (W - 200) // 2
            EH = 56

            Y = draw_arrow(els, Y)
            Y += 4

            inp = et.get('inp', '')
            if inp and inp != prev_out:
                Y = draw_oval(els, Y, inp)
                Y += 4
                Y = draw_arrow(els, Y, 20)
                Y += 4

            if et['tipo'] == 'SQLPLUS':
                RW, RH = 90, 48
                RX = EX - RW - 30
                RY = Y + (EH - RH) // 2
                els.append(f'<rect x="{RX}" y="{RY}" width="{RW}" height="{RH}" rx="10" fill="none" stroke="{BRD}" stroke-width="1.2"/>')
                els.append(f'<text x="{RX+RW//2}" y="{RY+18}" font-size="11" font-weight="bold" text-anchor="middle" fill="{GRY}" font-family="{FNT}">Rotina</text>')
                els.append(f'<text x="{RX+RW//2}" y="{RY+34}" font-size="11" font-weight="bold" text-anchor="middle" fill="{BLU}" font-family="{FNT}">{ex(cod)}</text>')
                els.append(f'<line x1="{RX+RW}" y1="{RY+RH//2}" x2="{EX}" y2="{Y+EH//2}" stroke="{BRD}" stroke-width="1.2"/>')
                els.append(f'<polygon points="{EX},{Y+EH//2} {EX-7},{Y+EH//2-4} {EX-7},{Y+EH//2+4}" fill="{BRD}"/>')
                CW, CRH = 90, 56
                CX = EX + EW + 30
                CY = Y + (EH - CRH) // 2
                CR = 10
                els.append(f'<ellipse cx="{CX+CW//2}" cy="{CY+CR}" rx="{CW//2}" ry="{CR}" fill="none" stroke="{BRD}" stroke-width="1.2"/>')
                els.append(f'<rect x="{CX}" y="{CY+CR}" width="{CW}" height="{CRH-CR*2}" fill="#fff" stroke="{BRD}" stroke-width="1.2"/>')
                els.append(f'<ellipse cx="{CX+CW//2}" cy="{CY+CRH-CR}" rx="{CW//2}" ry="{CR}" fill="none" stroke="{BRD}" stroke-width="1.2"/>')
                els.append(f'<text x="{CX+CW//2}" y="{CY+CRH//2}" font-size="10" text-anchor="middle" fill="{GRY}" font-family="{FNT}">Tabelas</text>')
                els.append(f'<text x="{CX+CW//2}" y="{CY+CRH//2+14}" font-size="10" text-anchor="middle" fill="{GRY}" font-family="{FNT}">Oracle</text>')
                els.append(f'<line x1="{CX}" y1="{CY+CRH//2}" x2="{EX+EW}" y2="{Y+EH//2}" stroke="{BRD}" stroke-width="1.2"/>')
                els.append(f'<polygon points="{EX+EW},{Y+EH//2} {EX+EW+7},{Y+EH//2-4} {EX+EW+7},{Y+EH//2+4}" fill="{BRD}"/>')

            els.append(f'<rect x="{EX}" y="{Y}" width="{EW}" height="{EH}" fill="none" stroke="{BRD}" stroke-width="1.8"/>')
            els.append(f'<line x1="{EX}" y1="{Y+28}" x2="{EX+EW}" y2="{Y+28}" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<text x="{EX+EW//2}" y="{Y+20}" font-size="13" font-weight="bold" text-anchor="middle" fill="{GRY}" font-family="{FNT}">ETAPA {str(et["num"]).zfill(2)}</text>')
            els.append(f'<text x="{EX+EW//2}" y="{Y+47}" font-size="11" text-anchor="middle" fill="{GRY}" font-family="{FNT}">{ex(et.get("titulo","") or et["tipo"])}</text>')
            Y += EH

            out = et.get('out', '')
            if out:
                Y += 4
                Y = draw_arrow(els, Y, 20)
                Y += 4
                Y = draw_oval(els, Y, out)

            prev_out = out
            Y += 10

        # SFTP nesta folha se cabe
        is_last = fi == len(folhas_etapas) - 1
        if tem_sftp and is_last and not ultima_cheia:
            Y = draw_arrow(els, Y, 28)
            Y += 4
            Y = draw_sftp(els, Y)
            Y = draw_fim(els, Y)
        elif is_last and not tem_sftp:
            Y = draw_fim(els, Y)

        if fi == 0 and tabelas:
            els.append(f'<text x="10" y="{Y+8}" font-size="8" fill="{LGY}" font-family="{FNT}">Tabelas Oracle: {ex(", ".join(tabelas))}</text>')
            Y += 16

        Y += 10
        Y = draw_footer(els, Y, fnum)
        svgs.append(make_svg(els, Y + 4))

    # Folha extra para SFTP se a última estava cheia
    if tem_sftp and ultima_cheia:
        els = []
        Y = draw_header(els, 0, total_folhas)
        if prev_out:
            Y = draw_arrow(els, Y)
            Y += 4
            Y = draw_oval(els, Y, prev_out)
            Y += 4
        Y = draw_arrow(els, Y, 28)
        Y += 4
        Y = draw_sftp(els, Y)
        Y = draw_fim(els, Y)
        Y += 10
        Y = draw_footer(els, Y, total_folhas)
        svgs.append(make_svg(els, Y + 4))

    return svgs


# ═══════════════════════════════════════════════════════════════════
# GERADOR DE FLUXOGRAMA EM WORD (.docx)
# ═══════════════════════════════════════════════════════════════════

def gerar_fluxograma_docx(svgs: list, cod: str = 'fluxograma') -> bytes:
    """Converte lista de SVGs do fluxograma em .docx.
    Se o conteudo total couber em uma pagina A4, combina em uma imagem so.
    Caso contrario, usa uma pagina Word por folha SVG.
    """
    import cairosvg
    from docx import Document
    from docx.shared import Cm

    PAGE_W_CM = 18.0   # largura util: 21 - 1.5 - 1.5
    PAGE_H_CM = 27.7   # altura util:  29.7 - 1 - 1
    SVG_W_PX  = 700    # viewBox width padrao dos SVGs gerados
    GAP_PX    = 30     # espaco entre folhas no SVG combinado

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    def svg_height(svg_str):
        m = re.search(r'viewBox="0 0 \d+ (\d+)"', svg_str)
        return int(m.group(1)) if m else 500

    def combine_svgs(svg_list):
        heights = [svg_height(s) for s in svg_list]
        total_h = sum(heights) + (len(svg_list) - 1) * GAP_PX
        els = []
        offset = 0
        for s, h in zip(svg_list, heights):
            inner = re.search(r'<svg[^>]*>(.*)</svg>', s, re.DOTALL)
            if inner:
                els.append(f'<g transform="translate(0,{offset})">{inner.group(1)}</g>')
            offset += h + GAP_PX
        return (
            f'<svg viewBox="0 0 {SVG_W_PX} {total_h}" '
            f'xmlns="http://www.w3.org/2000/svg" width="{SVG_W_PX}" style="background:#fff">\n'
            + '\n'.join(els) + '\n</svg>'
        )

    if len(svgs) == 1:
        png_bytes = cairosvg.svg2png(bytestring=svgs[0].encode('utf-8'), output_width=1400)
        doc.add_picture(io.BytesIO(png_bytes), width=Cm(PAGE_W_CM))
    else:
        heights = [svg_height(s) for s in svgs]
        total_h_px = sum(heights) + (len(svgs) - 1) * GAP_PX
        combined_h_cm = total_h_px * (PAGE_W_CM / SVG_W_PX)

        if combined_h_cm <= PAGE_H_CM:
            # Cabe numa pagina: uma unica imagem combinada
            combined_svg = combine_svgs(svgs)
            png_bytes = cairosvg.svg2png(
                bytestring=combined_svg.encode('utf-8'), output_width=1400
            )
            doc.add_picture(io.BytesIO(png_bytes), width=Cm(PAGE_W_CM))
        else:
            # Nao cabe: uma pagina Word por folha SVG
            for i, svg_str in enumerate(svgs):
                png_bytes = cairosvg.svg2png(
                    bytestring=svg_str.encode('utf-8'), output_width=1400
                )
                doc.add_picture(io.BytesIO(png_bytes), width=Cm(PAGE_W_CM))
                if i < len(svgs) - 1:
                    doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════
# MONTAGEM DAS ETAPAS A PARTIR DO SQL
# ═══════════════════════════════════════════════════════════════════

def montar_etapas(dados_sql: dict, compactar: bool, transferir: bool) -> list:
    """Monta a lista de etapas a partir dos dados extraídos do SQL."""
    etapas = []
    num = 1

    # ETAPA 01 — SQLPLUS
    arquivo_saida = dados_sql.get('arquivo_saida', '')
    obs_sql = descrever_operacao_sql(dados_sql)
    etapas.append({
        'num': num,
        'tipo': 'SQLPLUS',
        'titulo': 'Executar Script SQL',
        'inp': dados_sql.get('nome_script', ''),
        'out': arquivo_saida,
        'obs': obs_sql,
    })
    num += 1

    # ETAPA 02 — COMPACTAR (se selecionado)
    if compactar:
        inp_compact = arquivo_saida
        out_compact = arquivo_saida + '.zip' if arquivo_saida else ''
        etapas.append({
            'num': num,
            'tipo': 'COMPACTAR',
            'titulo': 'Compactar arquivo',
            'inp': inp_compact,
            'out': out_compact,
            'obs': '',
        })
        num += 1

    # ETAPA 03 — TRANSFERIR (se selecionado)
    if transferir:
        inp_transf = etapas[-1].get('out', '') if etapas else ''
        etapas.append({
            'num': num,
            'tipo': 'TRANSFERIR',
            'titulo': 'Transferir arquivo via SFTP',
            'inp': inp_transf,
            'out': '',
            'obs': '',
        })

    return etapas


# ═══════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ═══════════════════════════════════════════════════════════════════

# Estado
if 'resultados' not in st.session_state:
    st.session_state.resultados = []

col_form, col_output = st.columns([1, 1.5])

with col_form:
    st.markdown("### 📂 Upload do(s) Script(s) SQL")

    st.markdown(
        '<div class="info-box">Faça upload de um ou mais <strong>scripts SQL</strong> (.sql). '
        'O programa gera automaticamente a <strong>Característica (.docx)</strong> e o '
        '<strong>Fluxograma (SVG)</strong> para cada script.</div>',
        unsafe_allow_html=True
    )

    sql_files = st.file_uploader(
        "Scripts SQL (.sql / .txt)",
        type=['sql', 'txt'],
        key='sql_upload',
        accept_multiple_files=True,
    )

    st.markdown("---")
    st.markdown("### ⚙ Dados da Rotina")
    st.markdown(
        '<div class="warn-box">Preencha os campos abaixo. Se subir <strong>mais de um script</strong>, '
        'os mesmos dados serão aplicados a todos (você pode ajustar o código da rotina individualmente).</div>',
        unsafe_allow_html=True
    )

    c1, c2 = st.columns(2)
    with c1:
        cod_rotina = st.text_input("Código da Rotina", value="", placeholder="Ex: SH07681B")
    with c2:
        amb = st.selectbox("Ambiente", ['UNIX', 'Windows'], index=0)

    c1, c2 = st.columns(2)
    with c1:
        sistema = st.text_input("Sistema", value="SIGA SAUDE")
    with c2:
        subsistema = st.text_input("Subsistema", value="", placeholder="Ex: SH0768 / SIGA SAÚDE")

    denominacao = st.text_input("Denominação do Programa", value="", placeholder="Ex: Automatização - Relatório...")

    c1, c2 = st.columns(2)
    with c1:
        # Mês em português sem depender de locale
        import datetime as _dt
        _meses = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
                  7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}
        _hoje = _dt.date.today()
        _data_default = f"{_meses[_hoje.month]} {_hoje.year}"
        data_elab = st.text_input("Data da Elaboração", value=_data_default, placeholder="Ex: Junho 2026")
    with c2:
        horario = st.text_input("Horário de Execução", value="00h00", placeholder="Ex: 00h00")

    st.markdown("---")
    st.markdown("### 📦 Etapas adicionais")

    incluir_compactar = st.checkbox("Incluir etapa COMPACTAR (ZIP)", value=True)
    incluir_transferir = st.checkbox("Incluir etapa TRANSFERIR (SFTP)", value=True)

    if incluir_transferir:
        st.markdown("#### 🔗 Dados SFTP")
        c1, c2 = st.columns(2)
        with c1:
            sftp_ip = st.text_input("IP do servidor", placeholder="Ex: 144.22.186.214")
            sftp_proto = st.selectbox("Protocolo", ['SFTP', 'FTP'], index=0)
            sftp_user = st.text_input("Usuário", placeholder="Ex: prodam")
        with c2:
            sftp_porta = st.text_input("Porta", placeholder="Ex: 44322")
            sftp_pasta = st.text_input("Pasta destino", placeholder="Ex: /dados/relatorios")
            sftp_senha = st.text_input("Senha", type="password", placeholder="")

    st.markdown("---")

    if st.button("▶▶ Gerar Característica + Fluxograma", type="primary", use_container_width=True):
        if not sql_files:
            st.error("Faça upload de pelo menos um script SQL.")
        elif not cod_rotina.strip():
            st.error("Preencha o Código da Rotina.")
        else:
            resultados = []

            for idx, sql_file in enumerate(sql_files):
                texto_sql = sql_file.read().decode('utf-8', errors='replace')
                nome_script = sql_file.name
                dados_sql = extrair_dados_sql(texto_sql, nome_script)

                # Se tem mais de um script, adicionar sufixo ao código
                cod_atual = cod_rotina.strip()
                if len(sql_files) > 1:
                    # Usa letra A, B, C... — sempre aplica, independente do ultimo char
                    sufixo = chr(65 + idx) if idx < 26 else str(idx + 1)
                    cod_atual = cod_atual + sufixo

                # Montar etapas
                etapas = montar_etapas(dados_sql, incluir_compactar, incluir_transferir)

                # Dados completos
                dados = {
                    'cod': cod_atual,
                    'sis': sistema,
                    'sub': subsistema,
                    'den': denominacao or f'Rotina {cod_atual}',
                    'amb': amb,
                    'dat': data_elab,
                    'hora': horario,
                    'ip': sftp_ip if incluir_transferir else '',
                    'porta': sftp_porta if incluir_transferir else '',
                    'pasta': sftp_pasta if incluir_transferir else '',
                    'user': sftp_user if incluir_transferir else '',
                    'senha': sftp_senha if incluir_transferir else '',
                    'proto': sftp_proto if incluir_transferir else 'SFTP',
                    'etapas': etapas,
                    'tabelas': dados_sql['tabelas'],
                }

                # Gerar Característica .docx
                docx_bytes = gerar_caracteristica_docx(dados)

                # Gerar Fluxograma SVG
                svgs = gerar_svg(dados)

                # Gerar Fluxograma .docx (SVG → PNG → Word)
                try:
                    fluxo_docx_bytes = gerar_fluxograma_docx(svgs, cod_atual)
                except Exception as e:
                    fluxo_docx_bytes = None
                    st.warning(f"Fluxograma .docx não gerado (cairosvg indisponível): {e}")

                resultados.append({
                    'cod': cod_atual,
                    'nome_script': nome_script,
                    'dados': dados,
                    'dados_sql': dados_sql,
                    'docx_bytes': docx_bytes,
                    'fluxo_docx_bytes': fluxo_docx_bytes,
                    'svgs': svgs,
                })

                st.markdown(
                    f'<div class="success-box">✅ <strong>{nome_script}</strong> — '
                    f'{len(dados_sql["tabelas"])} tabelas, '
                    f'{len(etapas)} etapas, '
                    f'arquivo saída: <code>{dados_sql["arquivo_saida"] or "N/A"}</code></div>',
                    unsafe_allow_html=True
                )

            st.session_state.resultados = resultados

    # Mostra dados extraídos
    if st.session_state.resultados:
        for i, res in enumerate(st.session_state.resultados):
            d = res['dados']
            ds = res['dados_sql']
            with st.expander(f"📋 {res['cod']} — {res['nome_script']}", expanded=False):
                st.markdown(f"""
| Campo | Valor |
|---|---|
| Código | `{d['cod']}` |
| Sistema | {d['sis']} |
| Subsistema | {d['sub']} |
| Denominação | {d['den']} |
| Etapas | {len(d['etapas'])} |
| Tabelas SQL | {', '.join(ds['tabelas']) if ds['tabelas'] else 'N/A'} |
| Arquivo saída | `{ds['arquivo_saida'] or 'N/A'}` |
| Operações | {descrever_operacao_sql(ds)} |
""")

with col_output:
    st.markdown("### 📊 Resultado")

    if st.session_state.resultados:
        for res in st.session_state.resultados:
            cod = res['cod']
            svgs = res['svgs']

            if len(st.session_state.resultados) > 1:
                st.markdown(f"#### 🔹 {cod} — {res['nome_script']}")

            # ─── Downloads ──────────────────────────────
            st.markdown("##### 📥 Downloads")
            dl_cols = st.columns(3)

            with dl_cols[0]:
                st.download_button(
                    label=f"⬇ Característica .docx",
                    data=res['docx_bytes'],
                    file_name=f"Caracteristica_{cod}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_docx_{cod}_{i}",
                    use_container_width=True,
                )

            with dl_cols[1]:
                if res.get('fluxo_docx_bytes'):
                    st.download_button(
                        label=f"⬇ Fluxograma .docx",
                        data=res['fluxo_docx_bytes'],
                        file_name=f"Fluxograma_{cod}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_fluxo_docx_{cod}_{i}",
                        use_container_width=True,
                    )

            with dl_cols[2]:
                if len(svgs) == 1:
                    st.download_button(
                        label=f"⬇ Fluxograma .svg",
                        data=svgs[0].encode('utf-8'),
                        file_name=f"Fluxograma_{cod}.svg",
                        mime="image/svg+xml",
                        key=f"dl_svg_{cod}_{i}",
                        use_container_width=True,
                    )
                else:
                    # Combina múltiplas folhas em um SVG
                    W = 700
                    combined_els = []
                    total_h = 0
                    gap = 30
                    for svg_str in svgs:
                        vb = re.search(r'viewBox="0 0 (\d+) (\d+)"', svg_str)
                        h = int(vb.group(2)) if vb else 500
                        inner = re.search(r'<svg[^>]*>(.*)</svg>', svg_str, re.DOTALL)
                        if inner:
                            combined_els.append(f'<g transform="translate(0,{total_h})">{inner.group(1)}</g>')
                        total_h += h + gap
                    combined_svg = f'<svg viewBox="0 0 {W} {total_h}" xmlns="http://www.w3.org/2000/svg" width="{W}" style="background:#fff;max-width:100%">\n' + '\n'.join(combined_els) + '\n</svg>'
                    st.download_button(
                        label=f"⬇ Fluxograma_{cod} ({len(svgs)} folhas).svg",
                        data=combined_svg.encode('utf-8'),
                        file_name=f"Fluxograma_{cod}_Completo.svg",
                        mime="image/svg+xml",
                        key=f"dl_svg_{cod}_{i}",
                        use_container_width=True,
                    )

            # ─── Preview do fluxograma ──────────────────
            st.markdown("##### 👁 Preview do Fluxograma")
            for i, svg in enumerate(svgs):
                if len(svgs) > 1:
                    st.caption(f"Folha {i+1}/{len(svgs)}")
                st.markdown(svg, unsafe_allow_html=True)

            if len(st.session_state.resultados) > 1:
                st.markdown("---")

    else:
        st.markdown("""
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:400px;color:#999;text-align:center">
            <div style="font-size:48px;margin-bottom:12px;opacity:.3">📊</div>
            <div style="font-size:14px;line-height:1.6">
                Faça upload do(s) script(s) SQL,<br>
                preencha os dados da rotina<br>
                e clique em <strong>Gerar Característica + Fluxograma</strong>
            </div>
        </div>
        """, unsafe_allow_html=True)
