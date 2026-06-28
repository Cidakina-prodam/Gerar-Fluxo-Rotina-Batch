import streamlit as st
import re
import io

# ─── CONFIGURAÇÃO DA PÁGINA ────────────────────────────────────────
st.set_page_config(
    page_title="Gerador de Fluxograma Batch — PRODAM",
    page_icon="📊",
    layout="wide",
)

# ─── CSS CUSTOMIZADO ───────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { max-width: 1200px; padding-top: 1rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .header-bar {
        background: #1F4E79; color: #A8C8E8; padding: 12px 20px;
        border-radius: 6px; margin-bottom: 16px;
        display: flex; align-items: center; justify-content: space-between;
    }
    .header-bar h2 { color: #fff; margin: 0; font-size: 18px; }
    .header-bar .badge { background: #2E75B6; color: #A8C8E8; font-size: 11px; padding: 3px 10px; border-radius: 12px; }
    .etapa-box { border: 1px solid #ddd; border-radius: 6px; padding: 12px; margin-bottom: 8px; background: #fafafa; }
    .etapa-header { font-weight: 700; color: #1F4E79; font-size: 13px; margin-bottom: 6px; }
    .success-box { background: #E8F5E9; border-left: 3px solid #1B5E20; padding: 10px 14px; border-radius: 4px; margin: 8px 0; }
    .info-box { background: #E3F0FA; border-left: 3px solid #1565A0; padding: 10px 14px; border-radius: 4px; margin: 8px 0; }
</style>
""", unsafe_allow_html=True)

# ─── CABEÇALHO ─────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <div>
        <h2>■ PRODAM — Gerador de Fluxograma Batch</h2>
        <div style="font-size:12px;color:#7AABCC;margin-top:2px">Upload dos documentos → geração automática do fluxograma</div>
    </div>
    <span class="badge">SIGA SAÚDE · V2.0</span>
</div>
""", unsafe_allow_html=True)

# ─── FUNÇÕES DE LEITURA DE ARQUIVOS ────────────────────────────────

def ler_doc_binario(data: bytes) -> str:
    """Extrai texto legível de arquivo .doc binário (Word 97-2003)."""
    chars = []
    for b in data:
        if (32 <= b <= 126) or b in (10, 13) or (160 <= b <= 255):
            chars.append(chr(b))
        else:
            chars.append(' ')
    raw = ''.join(chars)
    chunks = re.findall(r'[\x20-\x7E\xA0-\xFF\r\n]{4,}', raw)
    return ' '.join(chunks).replace('  ', ' ').strip()


def ler_docx(data: bytes) -> str:
    """Extrai texto de arquivo .docx usando python-docx."""
    try:
        from docx import Document
        doc = Document(io.BytesIO(data))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        # tabelas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        paragraphs.append(cell.text.strip())
        return '\n'.join(paragraphs)
    except Exception:
        return ler_doc_binario(data)


def ler_arquivo_doc(uploaded_file) -> str:
    """Lê arquivo .doc ou .docx e retorna texto."""
    data = uploaded_file.read()
    nome = uploaded_file.name.lower()
    if nome.endswith('.docx'):
        return ler_docx(data)
    elif nome.endswith('.doc'):
        # tenta python-docx primeiro (às vezes funciona com .doc)
        texto = ''
        try:
            from docx import Document
            doc = Document(io.BytesIO(data))
            texto = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
        except Exception:
            pass
        if not texto.strip():
            texto = ler_doc_binario(data)
        return texto
    else:
        return data.decode('utf-8', errors='replace')


# ─── PARSER DA CARACTERÍSTICA ─────────────────────────────────────

def extrair_dados_caracteristica(texto: str) -> dict:
    """Extrai informações estruturadas do documento de Característica."""
    dados = {
        'cod': '', 'sis': '', 'sub': '', 'den': '', 'amb': 'UNIX',
        'dat': '', 'hora': '',
        'ip': '', 'porta': '', 'pasta': '', 'user': '', 'proto': 'SFTP',
        'etapas': []
    }

    # Código da rotina
    m = re.search(r'(?:CÓDIGO\s*DA\s*ROTINA|CODIGO\s*DA\s*ROTINA)[:\s]*([A-Z0-9]{6,10})', texto, re.IGNORECASE)
    if m:
        dados['cod'] = m.group(1).strip()

    # Sistema
    m = re.search(r'SISTEMA[:\s]+([A-Z][A-Z\s]{3,20}?)(?:\s{2,}|SUB)', texto, re.IGNORECASE)
    if m:
        dados['sis'] = m.group(1).strip()

    # Subsistema
    m = re.search(r'SUBSISTEMA[:\s]+(.+?)(?:\r|\n|CÓD|COD|DENOM)', texto, re.IGNORECASE)
    if m:
        dados['sub'] = m.group(1).strip()[:40]

    # Denominação
    m = re.search(r'DENOMINA[ÇC][ÃA]O\s+DO\s+PROGRAMA[:\s]+(.+?)(?:DATA|FOLHA|\r|\n)', texto, re.IGNORECASE)
    if m:
        dados['den'] = m.group(1).strip()[:100]

    # Data
    m = re.search(r'DATA\s+DA\s+ELABORA[ÇC][ÃA]O[:\s]+(.+?)(?:FOLHA|\r|\n)', texto, re.IGNORECASE)
    if m:
        dados['dat'] = m.group(1).strip()[:20]

    # Horário de execução
    m = re.search(r'(\d{1,2}h\d{2})', texto)
    if m:
        dados['hora'] = m.group(1)

    # Ambiente
    if re.search(r'\(UNIX\)|UNIX', texto, re.IGNORECASE):
        dados['amb'] = 'UNIX'
    elif re.search(r'\(Windows\)|Windows', texto, re.IGNORECASE):
        dados['amb'] = 'Windows'

    # SFTP / transferência
    m = re.search(r'IP[:\s]+(\d+\.\d+\.\d+\.\d+)', texto)
    if m:
        dados['ip'] = m.group(1)
    m = re.search(r'[Pp]orta[:\s]+(\d+)', texto)
    if m:
        dados['porta'] = m.group(1)
    m = re.search(r'[Pp]asta[:\s]+(/[\w\-/]+)', texto)
    if m:
        dados['pasta'] = m.group(1)
    m = re.search(r'[Uu]su[áa]rio[:\s]+(\S+)', texto)
    if m:
        dados['user'] = m.group(1)

    # Etapas
    etapa_matches = list(re.finditer(r'ETAPA\s*(\d{1,2})', texto, re.IGNORECASE))
    for idx, em in enumerate(etapa_matches):
        num = em.group(1).lstrip('0') or '0'
        if num == '0':
            continue  # etapa 00 é configuração

        # Pega o texto entre esta etapa e a próxima (ou fim)
        start = em.end()
        end = etapa_matches[idx + 1].start() if idx + 1 < len(etapa_matches) else min(start + 800, len(texto))
        bloco = texto[start:end]

        etapa = {'num': int(num), 'tipo': 'CUSTOM', 'titulo': '', 'inp': '', 'out': '', 'obs': ''}

        # Detecta tipo — COMPACTAR e TRANSFERIR têm prioridade sobre SQLPLUS
        # porque "SQLPLUS" aparece no cabeçalho de TODAS as etapas como "CÓD. PROGRAMA SQLPLUS"
        if re.search(r'COMPACTAR|ZIP\s+COMPRESS|Compactar\s+arq', bloco, re.IGNORECASE):
            etapa['tipo'] = 'COMPACTAR'
            etapa['titulo'] = 'Compactar arquivo'
        elif re.search(r'TRANSFER|SFTP|DISPONIBILIZAR', bloco, re.IGNORECASE):
            etapa['tipo'] = 'TRANSFERIR'
            etapa['titulo'] = 'Transferir arquivo via SFTP'
        elif re.search(r'SHELL|SCRIPT\s+SH|BASH', bloco, re.IGNORECASE):
            etapa['tipo'] = 'SHELL'
            etapa['titulo'] = 'Executar Script Shell'
        elif re.search(r'E-?MAIL|ENVIO', bloco, re.IGNORECASE):
            etapa['tipo'] = 'EMAIL'
            etapa['titulo'] = 'Envio de e-mail'
        elif re.search(r'EXECUTAR\s+SCRIPT|SELECT|SPOOL|script_\d+', bloco, re.IGNORECASE):
            etapa['tipo'] = 'SQLPLUS'
            etapa['titulo'] = 'Executar Script SQL'
        else:
            # Padrão: se é etapa 01 e nenhum outro tipo foi detectado, assume SQLPLUS
            if int(num) == 1:
                etapa['tipo'] = 'SQLPLUS'
                etapa['titulo'] = 'Executar Script SQL'

        # Arquivos — busca nomes de arquivo com extensão (incluindo & para variáveis PRODAM)
        arqs = re.findall(r'[\w\-\.&]+(?:\.csv|\.sql|\.zip|\.txt|\.dat|\.xml|\.log)(?:\.zip)?', bloco, re.IGNORECASE)
        # Remove duplicatas e fragmentos curtos
        arqs_uniq = list(dict.fromkeys([a for a in arqs if len(a) > 5]))
        for a in arqs_uniq:
            a_lower = a.lower()
            if a_lower.endswith('.sql') and not etapa['inp']:
                etapa['inp'] = a
            elif a_lower.endswith('.csv') and not etapa['out'] and etapa['tipo'] == 'SQLPLUS':
                etapa['out'] = a
            elif a_lower.endswith('.csv.zip') and not etapa['out']:
                etapa['out'] = a
            elif a_lower.endswith('.csv') and not etapa['inp'] and etapa['tipo'] != 'SQLPLUS':
                etapa['inp'] = a
            elif a_lower.endswith('.zip') and not etapa['out']:
                etapa['out'] = a

        # Arquivo com variável &x (padrão PRODAM)
        arqs_var = re.findall(r'[\w\-\.]+&\w+\.[\w\.]+', bloco)
        arqs_var = list(dict.fromkeys(arqs_var))
        for a in arqs_var:
            if etapa['tipo'] == 'SQLPLUS' and not etapa['out']:
                etapa['out'] = a
            elif etapa['tipo'] == 'COMPACTAR':
                if not etapa['inp'] and not a.endswith('.zip'):
                    etapa['inp'] = a
                elif not etapa['out'] and a.endswith('.zip'):
                    etapa['out'] = a
                elif not etapa['out']:
                    etapa['out'] = a + '.zip'
            elif etapa['tipo'] == 'TRANSFERIR':
                if not etapa['inp']:
                    etapa['inp'] = a

        # TRANSFERIR não produz arquivo de saída
        if etapa['tipo'] == 'TRANSFERIR':
            etapa['out'] = ''

        dados['etapas'].append(etapa)

    # Pós-processamento: preenche inp vazio com out da etapa anterior
    for i in range(1, len(dados['etapas'])):
        if not dados['etapas'][i]['inp'] and dados['etapas'][i - 1].get('out'):
            dados['etapas'][i]['inp'] = dados['etapas'][i - 1]['out']

    return dados


def extrair_dados_sql(texto: str) -> dict:
    """Extrai informações do script SQL."""
    dados = {'tabelas': [], 'arquivo_saida': '', 'arquivo_script': ''}

    # Tabelas (FROM e JOIN)
    tabelas = set()
    for m in re.finditer(r'(?:FROM|JOIN)\s+(TB_\w+|VW_\w+|DUAL)', texto, re.IGNORECASE):
        tabelas.add(m.group(1).upper())
    dados['tabelas'] = sorted(tabelas)

    # Arquivo de saída (SPOOL) — ignora SPOOL ON/OFF
    for m in re.finditer(r'SPOOL\s+(\S+)', texto, re.IGNORECASE):
        val = m.group(1).strip()
        if val.upper() not in ('ON', 'OFF', 'OUT'):
            dados['arquivo_saida'] = val
            break

    return dados


# ─── GERADOR DE SVG ────────────────────────────────────────────────

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
    MAX_ET = 2
    BLU, GRY, LGY, BRD = '#1F4E79', '#333', '#888', '#333'
    FNT = 'Arial Narrow,Arial,sans-serif'
    ex = escape_svg

    # Agrupa etapas em folhas (max 2 por folha)
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

# ─── INTERFACE STREAMLIT ───────────────────────────────────────────

col_form, col_output = st.columns([1, 1.5])

with col_form:
    st.markdown("### 📂 Upload dos arquivos")

    st.markdown('<div class="info-box">Faça upload da <strong>Característica</strong> (.doc/.docx) e/ou do <strong>Script SQL</strong> (.sql) — o programa extrai os dados automaticamente.</div>', unsafe_allow_html=True)

    doc_file = st.file_uploader(
        "Característica da Rotina (.doc / .docx / .txt)",
        type=['doc', 'docx', 'txt'],
        key='doc_upload'
    )

    sql_file = st.file_uploader(
        "Script SQL (.sql / .txt)",
        type=['sql', 'txt'],
        key='sql_upload'
    )

    # Estado
    if 'dados' not in st.session_state:
        st.session_state.dados = None
    if 'svgs' not in st.session_state:
        st.session_state.svgs = None

    if st.button("▶▶ Analisar e gerar fluxograma", type="primary", use_container_width=True):
        if not doc_file and not sql_file:
            st.error("Faça upload de pelo menos um arquivo.")
        else:
            dados = {
                'cod': '', 'sis': '', 'sub': '', 'den': '', 'amb': 'UNIX',
                'dat': '', 'hora': '', 'ip': '', 'porta': '', 'pasta': '',
                'user': '', 'proto': 'SFTP', 'etapas': [], 'tabelas': []
            }

            # Lê Característica
            if doc_file:
                with st.spinner("Lendo Característica..."):
                    texto_doc = ler_arquivo_doc(doc_file)
                    dados_doc = extrair_dados_caracteristica(texto_doc)
                    dados.update({k: v for k, v in dados_doc.items() if v})
                    st.markdown(f'<div class="success-box">✅ <strong>{doc_file.name}</strong> — {len(texto_doc)} caracteres extraídos, {len(dados_doc.get("etapas",[]))} etapas identificadas.</div>', unsafe_allow_html=True)

            # Lê SQL
            if sql_file:
                with st.spinner("Lendo Script SQL..."):
                    texto_sql = sql_file.read().decode('utf-8', errors='replace')
                    dados_sql = extrair_dados_sql(texto_sql)
                    dados['tabelas'] = dados_sql['tabelas']
                    if dados_sql['arquivo_saida'] and dados['etapas']:
                        for et in dados['etapas']:
                            if et['tipo'] == 'SQLPLUS' and not et['out']:
                                et['out'] = dados_sql['arquivo_saida']
                    if dados_sql['tabelas'] and dados['etapas']:
                        for et in dados['etapas']:
                            if et['tipo'] == 'SQLPLUS':
                                et['obs'] = 'Tabelas: ' + ', '.join(dados_sql['tabelas'])
                    st.markdown(f'<div class="success-box">✅ <strong>{sql_file.name}</strong> — {len(dados_sql["tabelas"])} tabelas identificadas.</div>', unsafe_allow_html=True)

            if dados['etapas']:
                st.session_state.dados = dados
                st.session_state.svgs = gerar_svg(dados)
            else:
                st.warning("Não foi possível identificar etapas automaticamente. Verifique se o documento segue o padrão da Característica PRODAM.")

    # Mostra dados extraídos
    if st.session_state.dados:
        d = st.session_state.dados
        with st.expander("📋 Dados extraídos", expanded=False):
            st.markdown(f"""
| Campo | Valor |
|---|---|
| Código | `{d['cod']}` |
| Sistema | {d['sis']} |
| Subsistema | {d['sub']} |
| Denominação | {d['den']} |
| Ambiente | {d['amb']} |
| Data | {d['dat']} |
| Horário | {d['hora']} |
| SFTP IP | {d['ip']} |
| Etapas | {len(d['etapas'])} |
| Tabelas SQL | {len(d.get('tabelas', []))} |
""")

with col_output:
    st.markdown("### 📊 Fluxograma de operação")

    if st.session_state.svgs:
        svgs = st.session_state.svgs
        cod = st.session_state.dados.get('cod', 'fluxograma') if st.session_state.dados else 'fluxograma'

        # Mostra cada folha
        for i, svg in enumerate(svgs):
            if len(svgs) > 1:
                st.markdown(f"---")
            st.markdown(svg, unsafe_allow_html=True)

            # Download individual
            st.download_button(
                label=f"⬇ Baixar Folha {str(i+1).zfill(2)} (SVG)",
                data=svg.encode('utf-8'),
                file_name=f"Fluxograma_{cod}_Folha{str(i+1).zfill(2)}.svg",
                mime="image/svg+xml",
                key=f"dl_{i}",
            )

        # Download de todas as folhas combinadas
        if len(svgs) > 1:
            all_svgs = '\n\n'.join(svgs)
            st.download_button(
                label=f"⬇ Baixar TODAS as folhas ({len(svgs)} SVGs)",
                data=all_svgs.encode('utf-8'),
                file_name=f"Fluxograma_{cod}_Completo.svg",
                mime="image/svg+xml",
                key="dl_all",
                use_container_width=True,
            )
    else:
        st.markdown("""
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:400px;color:#999;text-align:center">
            <div style="font-size:48px;margin-bottom:12px;opacity:.3">📊</div>
            <div style="font-size:14px;line-height:1.6">
                Faça upload dos arquivos<br>e clique em <strong>Analisar e gerar fluxograma</strong>
            </div>
        </div>
        """, unsafe_allow_html=True)
