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


def gerar_svg(dados: dict) -> str:
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

    W, PAD = 700, 50
    BLU, BLU2, GRY, LGY, BRD = '#1F4E79', '#2E75B6', '#444', '#999', '#555'
    Y = 0
    els = []
    ex = escape_svg

    # Cabeçalho institucional
    HH = 90
    els.append(f'<rect x="0" y="0" width="{W}" height="{HH}" fill="{BLU}"/>')
    els.append(f'<rect x="{PAD-10}" y="8" width="{W-PAD*2+20}" height="{HH-16}" fill="none" stroke="{BLU2}" stroke-width="1"/>')
    els.append(f'<text x="{W//2}" y="30" font-size="11" font-weight="bold" text-anchor="middle" fill="#9BB8D8" font-family="Arial Narrow,Arial,sans-serif" letter-spacing="2">PRODAM — SIGA SAÚDE</text>')
    els.append(f'<text x="{W//2}" y="52" font-size="14" font-weight="bold" text-anchor="middle" fill="#fff" font-family="Arial Narrow,Arial,sans-serif" letter-spacing="2">FLUXOGRAMA DE ROTINA DE OPERAÇÃO</text>')
    els.append(f'<text x="{W//2}" y="68" font-size="10" text-anchor="middle" fill="#D4A017" font-family="Arial Narrow,Arial,sans-serif">CÓDIGO: {ex(cod)} ({ex(amb)})  |  {ex(sis)}  |  SUBSISTEMA: {ex(sub)}</text>')
    els.append(f'<text x="{W//2}" y="82" font-size="9" text-anchor="middle" fill="#7AABCC" font-family="Arial Narrow,Arial,sans-serif">{ex(den)}  |  {ex(dat)}</text>')
    Y = HH + 20

    # Scheduler
    SW, SH, SX = 280, 36, (W - 280) // 2
    els.append(f'<rect x="{SX}" y="{Y}" width="{SW}" height="{SH}" rx="18" fill="none" stroke="{BLU2}" stroke-width="1.5"/>')
    els.append(f'<text x="{W//2}" y="{Y+14}" font-size="9" text-anchor="middle" fill="{LGY}" font-family="Arial Narrow,Arial,sans-serif">SCHEDULER — EXECUÇÃO DIÁRIA</text>')
    els.append(f'<text x="{W//2}" y="{Y+28}" font-size="13" font-weight="bold" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">{ex(hora)}</text>')
    Y += SH

    TIPOS_STYLE = {
        'SQLPLUS':    {'stroke': '#1565C0'},
        'COMPACTAR':  {'stroke': '#E65100'},
        'TRANSFERIR': {'stroke': '#1B5E20'},
        'SHELL':      {'stroke': '#6A1B9A'},
        'EMAIL':      {'stroke': '#880E4F'},
        'CUSTOM':     {'stroke': '#555'},
    }

    for et in etapas:
        st = TIPOS_STYLE.get(et['tipo'], TIPOS_STYLE['CUSTOM'])
        EW, EX = 220, (W - 220) // 2

        # Seta
        els.append(f'<line x1="{W//2}" y1="{Y}" x2="{W//2}" y2="{Y+22}" stroke="{GRY}" stroke-width="1.2"/>')
        els.append(f'<polygon points="{W//2},{Y+28} {W//2-5},{Y+20} {W//2+5},{Y+20}" fill="{GRY}"/>')
        Y += 28

        # Arquivo entrada (oval)
        if et.get('inp'):
            fw = min(W - PAD * 2, len(et['inp']) * 6.5 + 40)
            fh = 26
            els.append(f'<ellipse cx="{W//2}" cy="{Y+fh//2}" rx="{fw//2}" ry="{fh//2}" fill="none" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<text x="{W//2}" y="{Y+fh//2+4}" font-size="9.5" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">{ex(et["inp"])}</text>')
            Y += fh + 4
            els.append(f'<line x1="{W//2}" y1="{Y}" x2="{W//2}" y2="{Y+14}" stroke="{GRY}" stroke-width="1"/>')
            els.append(f'<polygon points="{W//2},{Y+18} {W//2-4},{Y+12} {W//2+4},{Y+12}" fill="{GRY}"/>')
            Y += 18

        # Elementos laterais para SQLPLUS
        EH1, EH2, EH = 26, 24, 50
        if et['tipo'] == 'SQLPLUS':
            # Rotina (rect arredondado, esquerda)
            RW, RH = 80, 44
            RX, RY = EX - RW - 20, Y + (EH - RH) // 2
            els.append(f'<rect x="{RX}" y="{RY}" width="{RW}" height="{RH}" rx="8" fill="none" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<text x="{RX+RW//2}" y="{RY+16}" font-size="10" font-weight="bold" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">Rotina</text>')
            els.append(f'<text x="{RX+RW//2}" y="{RY+30}" font-size="10" font-weight="bold" text-anchor="middle" fill="{BLU}" font-family="Arial Narrow,Arial,sans-serif">{ex(cod)}</text>')
            els.append(f'<line x1="{RX+RW}" y1="{RY+RH//2}" x2="{EX}" y2="{Y+EH//2}" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<polygon points="{EX},{Y+EH//2} {EX-6},{Y+EH//2-4} {EX-6},{Y+EH//2+4}" fill="{BRD}"/>')

            # Cilindro Tabelas Oracle (direita)
            CX, CY, CW, CRH = EX + EW + 20, Y + (EH - 44) // 2, 80, 44
            CRY_EL = 8
            els.append(f'<ellipse cx="{CX+CW//2}" cy="{CY+CRY_EL}" rx="{CW//2}" ry="{CRY_EL}" fill="none" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<rect x="{CX}" y="{CY+CRY_EL}" width="{CW}" height="{CRH-CRY_EL*2}" fill="none" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<ellipse cx="{CX+CW//2}" cy="{CY+CRH-CRY_EL}" rx="{CW//2}" ry="{CRY_EL}" fill="none" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<text x="{CX+CW//2}" y="{CY+CRH//2}" font-size="9.5" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">Tabelas</text>')
            els.append(f'<text x="{CX+CW//2}" y="{CY+CRH//2+13}" font-size="9.5" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">Oracle</text>')
            els.append(f'<line x1="{CX}" y1="{CY+CRH//2}" x2="{EX+EW}" y2="{Y+EH//2}" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<polygon points="{EX+EW},{Y+EH//2} {EX+EW+6},{Y+EH//2-4} {EX+EW+6},{Y+EH//2+4}" fill="{BRD}"/>')

        # Caixa da etapa
        els.append(f'<rect x="{EX}" y="{Y}" width="{EW}" height="{EH}" fill="none" stroke="{st["stroke"]}" stroke-width="1.8"/>')
        els.append(f'<rect x="{EX}" y="{Y}" width="{EW}" height="{EH1}" fill="{st["stroke"]}" opacity="0.08"/>')
        els.append(f'<line x1="{EX}" y1="{Y+EH1}" x2="{EX+EW}" y2="{Y+EH1}" stroke="{st["stroke"]}" stroke-width="1"/>')
        els.append(f'<text x="{EX+EW//2}" y="{Y+EH1-7}" font-size="12" font-weight="bold" text-anchor="middle" fill="{st["stroke"]}" font-family="Arial Narrow,Arial,sans-serif">ETAPA {str(et["num"]).zfill(2)}</text>')
        els.append(f'<text x="{EX+EW//2}" y="{Y+EH1+EH2-7}" font-size="11" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">{ex(et.get("titulo","") or et["tipo"])}</text>')
        Y += EH

        # Observações
        obs = et.get('obs', '')
        if obs:
            obs_lines = wrap_text(obs, 60)
            OBH = len(obs_lines) * 13 + 6
            els.append(f'<rect x="{EX}" y="{Y}" width="{EW}" height="{OBH}" fill="none" stroke="{st["stroke"]}" stroke-width="0.7" stroke-dasharray="3,2" opacity="0.5"/>')
            for li, l in enumerate(obs_lines):
                els.append(f'<text x="{EX+EW//2}" y="{Y+12+li*13}" font-size="9" text-anchor="middle" fill="{LGY}" font-family="Arial Narrow,Arial,sans-serif">{ex(l)}</text>')
            Y += OBH

        # Arquivo saída (oval)
        if et.get('out'):
            els.append(f'<line x1="{W//2}" y1="{Y}" x2="{W//2}" y2="{Y+14}" stroke="{GRY}" stroke-width="1"/>')
            els.append(f'<polygon points="{W//2},{Y+18} {W//2-4},{Y+12} {W//2+4},{Y+12}" fill="{GRY}"/>')
            Y += 18
            fw = min(W - PAD * 2, len(et['out']) * 6.5 + 40)
            fh = 26
            els.append(f'<ellipse cx="{W//2}" cy="{Y+fh//2}" rx="{fw//2}" ry="{fh//2}" fill="none" stroke="{BRD}" stroke-width="1"/>')
            els.append(f'<text x="{W//2}" y="{Y+fh//2+4}" font-size="9.5" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">{ex(et["out"])}</text>')
            Y += fh

        Y += 14

    # Bloco SFTP
    if ip:
        els.append(f'<line x1="{W//2}" y1="{Y}" x2="{W//2}" y2="{Y+22}" stroke="#1B5E20" stroke-width="1.2"/>')
        els.append(f'<polygon points="{W//2},{Y+28} {W//2-5},{Y+20} {W//2+5},{Y+20}" fill="#1B5E20"/>')
        Y += 28
        SW2 = 320
        SX2 = (W - SW2) // 2
        sfH = 82
        CRY = 12
        els.append(f'<ellipse cx="{SX2+SW2//2}" cy="{Y+CRY}" rx="{SW2//2}" ry="{CRY}" fill="none" stroke="#1B5E20" stroke-width="1.5"/>')
        els.append(f'<rect x="{SX2}" y="{Y+CRY}" width="{SW2}" height="{sfH-CRY*2}" fill="none" stroke="#1B5E20" stroke-width="1.5"/>')
        els.append(f'<ellipse cx="{SX2+SW2//2}" cy="{Y+sfH-CRY}" rx="{SW2//2}" ry="{CRY}" fill="none" stroke="#1B5E20" stroke-width="1.5"/>')
        els.append(f'<text x="{SX2+SW2//2}" y="{Y+28}" font-size="11" font-weight="bold" text-anchor="middle" fill="#1B5E20" font-family="Arial Narrow,Arial,sans-serif">SFTP — Servidor Externo</text>')
        els.append(f'<text x="{SX2+SW2//2}" y="{Y+43}" font-size="9.5" text-anchor="middle" fill="{GRY}" font-family="Arial Narrow,Arial,sans-serif">IP: {ex(ip)}  |  Porta: {ex(porta)}  |  Usuário: {ex(user)}</text>')
        els.append(f'<text x="{SX2+SW2//2}" y="{Y+57}" font-size="9" text-anchor="middle" fill="{LGY}" font-family="Arial Narrow,Arial,sans-serif">{ex(pasta)}</text>')
        Y += sfH + 14

    # Fim
    els.append(f'<line x1="{W//2}" y1="{Y}" x2="{W//2}" y2="{Y+18}" stroke="{GRY}" stroke-width="1.2"/>')
    els.append(f'<polygon points="{W//2},{Y+24} {W//2-5},{Y+16} {W//2+5},{Y+16}" fill="{GRY}"/>')
    Y += 24
    FW, FX, FH = 160, (W - 160) // 2, 32
    els.append(f'<rect x="{FX}" y="{Y}" width="{FW}" height="{FH}" rx="16" fill="{BLU}"/>')
    els.append(f'<text x="{W//2}" y="{Y+21}" font-size="12" font-weight="bold" text-anchor="middle" fill="#fff" font-family="Arial Narrow,Arial,sans-serif">FIM — RC=0</text>')
    Y += FH + 14

    # Rodapé
    els.append(f'<line x1="{PAD}" y1="{Y}" x2="{W-PAD}" y2="{Y}" stroke="{LGY}" stroke-width="0.5"/>')
    els.append(f'<text x="{PAD}" y="{Y+12}" font-size="8" fill="{LGY}" font-family="Arial Narrow,Arial,sans-serif">PRODAM / {ex(sis)} / {ex(cod)}</text>')
    els.append(f'<text x="{W-PAD}" y="{Y+12}" font-size="8" text-anchor="end" fill="{LGY}" font-family="Arial Narrow,Arial,sans-serif">FOLHA 01/01</text>')
    Y += 24

    # Tabelas SQL (se houver)
    if tabelas:
        els.append(f'<text x="{PAD}" y="{Y+12}" font-size="8" fill="{LGY}" font-family="Arial Narrow,Arial,sans-serif">Tabelas Oracle: {ex(", ".join(tabelas))}</text>')
        Y += 20

    svg = f'<svg viewBox="0 0 {W} {Y}" xmlns="http://www.w3.org/2000/svg" width="{W}" style="background:#fff;max-width:100%">'
    svg += '\n'.join(els)
    svg += '</svg>'
    return svg


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
    if 'svg' not in st.session_state:
        st.session_state.svg = None

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
                        # Atualiza arquivo de saída da etapa SQLPLUS
                        for et in dados['etapas']:
                            if et['tipo'] == 'SQLPLUS' and not et['out']:
                                et['out'] = dados_sql['arquivo_saida']
                    # Adiciona obs com tabelas na etapa SQLPLUS
                    if dados_sql['tabelas'] and dados['etapas']:
                        for et in dados['etapas']:
                            if et['tipo'] == 'SQLPLUS':
                                et['obs'] = 'Tabelas: ' + ', '.join(dados_sql['tabelas'])
                    st.markdown(f'<div class="success-box">✅ <strong>{sql_file.name}</strong> — {len(dados_sql["tabelas"])} tabelas identificadas.</div>', unsafe_allow_html=True)

            if dados['etapas']:
                st.session_state.dados = dados
                st.session_state.svg = gerar_svg(dados)
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

    if st.session_state.svg:
        # Mostra SVG
        st.markdown(st.session_state.svg, unsafe_allow_html=True)

        # Botão de download
        svg_bytes = st.session_state.svg.encode('utf-8')
        cod = st.session_state.dados.get('cod', 'fluxograma') if st.session_state.dados else 'fluxograma'
        st.download_button(
            label="⬇ Baixar SVG",
            data=svg_bytes,
            file_name=f"Fluxograma_{cod}.svg",
            mime="image/svg+xml",
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
