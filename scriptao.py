# -*- coding: utf-8 -*-
"""
landing_page_pmbok.py
================================================================================
Gera um documento ABNT (python-docx) completo do projeto "Landing Page da
Clínica Odontológica Sorriso Real", gerenciado segundo o Guia PMBOK, da
INICIAÇÃO ao ENCERRAMENTO, com figuras em SVG (salvas em figuras/ e convertidas
para PNG com fallback embutido via Pillow — não precisa de DLL do Cairo).

Dependências:
    pip install python-docx          # obrigatório
    (Pillow)                         # opcional: já vem no WinPython (fallback)
    pip install svglib reportlab     # opcional: melhor qualidade de conversão

Uso:
    python landing_page_pmbok.py

Saídas:
    Projeto_Landing_Page_Sorriso_Real_PMBOK.docx
    figuras/figura_XX_*.svg e *.png

No Word os campos (sumário/listas/página) se atualizam ao abrir; no
LibreOffice use Ferramentas > Atualizar > Atualizar tudo.
================================================================================
"""

from pathlib import Path
from xml.sax.saxutils import escape
from collections import defaultdict

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NOME_ARQUIVO = "Projeto_Landing_Page_Sorriso_Real_PMBOK.docx"
DIR_FIGURAS = Path("figuras")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
PRETO = RGBColor(0x00, 0x00, 0x00)
CONVERSOR_USADO = None
_CONT = {"Figura": 0, "Tabela": 0}

JUS = WD_ALIGN_PARAGRAPH.JUSTIFY
CEN = WD_ALIGN_PARAGRAPH.CENTER
ESQ = WD_ALIGN_PARAGRAPH.LEFT
DIR = WD_ALIGN_PARAGRAPH.RIGHT


# ==============================================================================
# Utilitários de documento
# ==============================================================================
def _add_campo(par, instrucao, cache="1", negrito=False, tamanho=None):
    r1 = par.add_run()
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin"); r1._r.append(fb)
    r2 = par.add_run()
    it = OxmlElement("w:instrText"); it.set(XML_SPACE, "preserve"); it.text = instrucao
    r2._r.append(it)
    r3 = par.add_run()
    fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate"); r3._r.append(fs)
    r4 = par.add_run(cache)
    r5 = par.add_run()
    fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end"); r5._r.append(fe)
    for r in (r1, r2, r3, r4, r5):
        if negrito:
            r.bold = True
        if tamanho:
            r.font.size = Pt(tamanho)
    return par


def _limpar_tema(style):
    rpr = style.element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rf.attrib.pop(qn("w:" + a), None)
    cor = rpr.find(qn("w:color"))
    if cor is not None:
        for a in ("themeColor", "themeShade", "themeTint"):
            cor.attrib.pop(qn("w:" + a), None)


def configurar_pagina(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.left_margin = Cm(3), Cm(3)
    sec.bottom_margin, sec.right_margin = Cm(2), Cm(2)
    sec.header_distance, sec.footer_distance = Cm(2), Cm(1.25)


def criar_estilos(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = PRETO
    _limpar_tema(normal)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Cm(1.25)
    pf.alignment = JUS

    for nivel, (antes, depois, italico) in {1: (0, 18, False), 2: (18, 10, False), 3: (12, 8, True)}.items():
        st = doc.styles[f"Heading {nivel}"]
        st.font.name = "Times New Roman"
        st.font.size = Pt(12)
        st.font.bold = True
        st.font.italic = italico
        st.font.color.rgb = PRETO
        _limpar_tema(st)
        sp = st.paragraph_format
        sp.line_spacing = 1.5
        sp.alignment = ESQ
        sp.first_line_indent = Cm(0)
        sp.space_before = Pt(antes)
        sp.space_after = Pt(depois)
        sp.keep_with_next = True
        if nivel == 1:
            sp.page_break_before = True

    def novo(nome):
        st = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
        return st

    leg = novo("Legenda")
    leg.font.size = Pt(11)
    lp = leg.paragraph_format
    lp.alignment = CEN; lp.first_line_indent = Cm(0); lp.line_spacing = 1.0
    lp.space_before = Pt(10); lp.space_after = Pt(4); lp.keep_with_next = True

    fnt = novo("Fonte de figura")
    fnt.font.size = Pt(10)
    fp = fnt.paragraph_format
    fp.alignment = CEN; fp.first_line_indent = Cm(0); fp.line_spacing = 1.0
    fp.space_before = Pt(4); fp.space_after = Pt(18)

    cit = novo("Citação longa")
    cit.font.size = Pt(10)
    cp = cit.paragraph_format
    cp.alignment = JUS; cp.left_indent = Cm(4); cp.first_line_indent = Cm(0)
    cp.line_spacing = 1.0; cp.space_before = Pt(12); cp.space_after = Pt(12)

    ref = novo("Referência ABNT")
    ref.font.size = Pt(12)
    rp = ref.paragraph_format
    rp.alignment = ESQ; rp.first_line_indent = Cm(0); rp.line_spacing = 1.0
    rp.space_after = Pt(12)

    tsn = novo("Título sem número")
    tsn.font.bold = True
    tp = tsn.paragraph_format
    tp.alignment = CEN; tp.first_line_indent = Cm(0); tp.line_spacing = 1.5
    tp.space_before = Pt(0); tp.space_after = Pt(18)
    tp.keep_with_next = True
    tp.page_break_before = True


def secao(doc, nivel, texto):
    return doc.add_paragraph(texto, style=f"Heading {nivel}")


def titulo_secao(doc, texto):
    return doc.add_paragraph(texto, style="Título sem número")


def paragrafo(doc, texto, recuo=True):
    par = doc.add_paragraph(texto)
    if not recuo:
        par.paragraph_format.first_line_indent = Cm(0)
    return par


def paragrafo_rotulo(doc, rotulo, texto):
    par = doc.add_paragraph()
    par.add_run(rotulo + " ").bold = True
    par.add_run(texto)
    return par


def marcador(doc, texto):
    par = doc.add_paragraph(texto, style="List Bullet")
    pf = par.paragraph_format
    pf.left_indent = Cm(0.75)
    pf.first_line_indent = Cm(-0.35)
    pf.space_after = Pt(4)
    return par


def citacao_longa(doc, texto):
    return doc.add_paragraph(texto, style="Citação longa")


def linha_vazia(doc, n=1):
    for _ in range(n):
        par = doc.add_paragraph()
        par.paragraph_format.first_line_indent = Cm(0)


def centro(doc, texto, negrito=False, tam=12):
    par = doc.add_paragraph()
    par.alignment = CEN
    par.paragraph_format.first_line_indent = Cm(0)
    r = par.add_run(texto)
    r.bold = negrito
    r.font.size = Pt(tam)
    return par


def _proximo(tipo):
    _CONT[tipo] += 1
    return _CONT[tipo]


def legenda(doc, tipo, titulo, n=None):
    n = n if n is not None else _proximo(tipo)
    par = doc.add_paragraph(style="Legenda")
    par.add_run(f"{tipo} ")
    _add_campo(par, f"SEQ {tipo} \\* ARABIC", cache=str(n))
    par.add_run(f" – {titulo}")
    return par


def fonte(doc, texto):
    par = doc.add_paragraph(style="Fonte de figura")
    par.add_run("Fonte: " + texto)
    return par


def add_toc(doc, instrucao, nota):
    par = doc.add_paragraph()
    par.alignment = ESQ
    par.paragraph_format.first_line_indent = Cm(0)
    _add_campo(par, instrucao, cache=nota)
    return par


def escrever_celula(cel, texto, negrito=False, centro_flag=False, tam=10):
    par = cel.paragraphs[0]
    pf = par.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.0
    pf.space_before = Pt(2)
    pf.space_after = Pt(2)
    par.alignment = CEN if centro_flag else ESQ
    r = par.add_run(texto)
    r.bold = negrito
    r.font.size = Pt(tam)


def sombrear(cel, cor="E7EFE2"):
    tcpr = cel._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), cor)
    tcpr.append(shd)


def tabela(doc, titulo, texto_fonte, cabecalho, linhas, larguras=None,
           tam_fonte=10, alinhar_centro=None):
    _proximo("Tabela")
    legenda(doc, "Tabela", titulo)
    t = doc.add_table(rows=1, cols=len(cabecalho))
    t.style = doc.styles["Table Grid"]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for j, txt in enumerate(cabecalho):
        cel = t.rows[0].cells[j]
        escrever_celula(cel, txt, negrito=True, centro_flag=True, tam=tam_fonte)
        sombrear(cel)
        if larguras:
            cel.width = Cm(larguras[j])
    for linha in linhas:
        destaque = str(linha[0]).upper().startswith("TOTAL")
        row = t.add_row()
        for j, txt in enumerate(linha):
            cel = row.cells[j]
            fl = bool(alinhar_centro) and j in alinhar_centro
            escrever_celula(cel, txt, negrito=destaque, centro_flag=fl, tam=tam_fonte)
            if larguras:
                cel.width = Cm(larguras[j])
    fonte(doc, texto_fonte)


def figura(doc, builder, slug, titulo, texto_fonte, largura_cm=14.0):
    n = _proximo("Figura")
    nome = f"figura_{n:02d}_{slug}"
    svg_path = DIR_FIGURAS / f"{nome}.svg"
    png_path = DIR_FIGURAS / f"{nome}.png"
    png = svg_para_png(builder(), svg_path, png_path)
    legenda(doc, "Figura", titulo, n=n)
    if png:
        doc.add_picture(str(png), width=Cm(largura_cm))
        par = doc.paragraphs[-1]
        par.alignment = CEN
        par.paragraph_format.first_line_indent = Cm(0)
        par.paragraph_format.space_after = Pt(4)
    else:
        par = doc.add_paragraph()
        par.alignment = CEN
        par.paragraph_format.first_line_indent = Cm(0)
        r = par.add_run(f"[Figura {n}: converta '{svg_path.name}' manualmente.]")
        r.font.size = Pt(10)
    fonte(doc, texto_fonte)


# ==============================================================================
# Conversão SVG -> PNG (cairosvg -> svglib -> fallback Pillow embutido)
# ==============================================================================
def _carregar_fonte(tam, peso="normal", italico=False):
    neg = (peso == "bold")
    if neg and italico:
        candidatos = ["arialbi.ttf", "DejaVuSans-BoldOblique.ttf",
                      "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf"]
    elif neg:
        candidatos = ["arialbd.ttf", "DejaVuSans-Bold.ttf",
                      "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
    elif italico:
        candidatos = ["ariali.ttf", "DejaVuSans-Oblique.ttf",
                      "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"]
    else:
        candidatos = ["arial.ttf", "DejaVuSans.ttf",
                      "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
    for nome in candidatos:
        try:
            from PIL import ImageFont
            return ImageFont.truetype(nome, max(6, int(round(tam))))
        except Exception:
            continue
    from PIL import ImageFont
    return ImageFont.load_default()


def _png_via_pillow(svg_str, png_path, escala):
    import re
    from xml.sax.saxutils import unescape
    from PIL import Image, ImageDraw

    m = re.search(r'<svg[^>]*?width="([\d.]+)"[^>]*?height="([\d.]+)"', svg_str)
    W, H = float(m.group(1)), float(m.group(2))
    img = Image.new("RGB", (int(W * escala), int(H * escala)), "#ffffff")
    dr = ImageDraw.Draw(img)

    def attrs(tag):
        return dict(re.findall(r'([\w:-]+)="([^"]*)"', tag))

    def cor(v, padrao=None):
        return padrao if v in (None, "none", "") else v

    def pts(s):
        return [(float(a) * escala, float(b) * escala)
                for a, b in (p.split(",") for p in s.split())]

    for el in re.finditer(r'<(rect|circle|polygon|polyline|line|text)\b([^>]*)>(?:([^<]*)</\1>)?', svg_str):
        tag, at, conteudo = el.group(1), attrs(el.group(2)), (el.group(3) or "")
        larg_st = max(1, int(float(at.get("stroke-width", 1)) * escala))
        if tag == "rect":
            x, y = float(at.get("x", 0)) * escala, float(at.get("y", 0)) * escala
            w, h = float(at["width"]) * escala, float(at["height"]) * escala
            caixa = [x, y, x + w, y + h]
            rx = float(at.get("rx", 0)) * escala
            if rx:
                dr.rounded_rectangle(caixa, radius=rx, fill=cor(at.get("fill")),
                                     outline=cor(at.get("stroke")), width=larg_st)
            else:
                dr.rectangle(caixa, fill=cor(at.get("fill")),
                             outline=cor(at.get("stroke")), width=larg_st)
        elif tag == "circle":
            cx, cy, r = float(at["cx"]) * escala, float(at["cy"]) * escala, float(at["r"]) * escala
            dr.ellipse([cx - r, cy - r, cx + r, cy + r], fill=cor(at.get("fill")),
                       outline=cor(at.get("stroke")), width=larg_st)
        elif tag == "polygon":
            dr.polygon(pts(at["points"]), fill=cor(at.get("fill")),
                       outline=cor(at.get("stroke")))
        elif tag == "polyline":
            dr.line(pts(at["points"]), fill=cor(at.get("stroke"), "#000000"), width=larg_st)
        elif tag == "line":
            dr.line([float(at["x1"]) * escala, float(at["y1"]) * escala,
                     float(at["x2"]) * escala, float(at["y2"]) * escala],
                    fill=cor(at.get("stroke"), "#000000"), width=larg_st)
        elif tag == "text":
            if not conteudo.strip():
                continue
            txt = unescape(conteudo.strip())
            x, y = float(at["x"]) * escala, float(at["y"]) * escala
            tam = float(at.get("font-size", 12)) * escala
            ancora = {"middle": "ms", "end": "rs"}.get(at.get("text-anchor", "start"), "ls")
            f = _carregar_fonte(tam, at.get("font-weight", "normal"),
                                at.get("font-style") == "italic")
            fill = at.get("fill", "#000000")
            try:
                dr.text((x, y), txt, font=f, fill=fill, anchor=ancora)
            except Exception:
                cb = dr.textbbox((0, 0), txt, font=f)
                larg, alt = cb[2] - cb[0], cb[3] - cb[1]
                if ancora[0] == "m":
                    x -= larg / 2
                elif ancora[0] == "r":
                    x -= larg
                dr.text((x, y - alt), txt, font=f, fill=fill)
    img.save(png_path)


def svg_para_png(svg_str, svg_path, png_path, escala=2.5):
    global CONVERSOR_USADO
    svg_path.write_text(svg_str, encoding="utf-8")
    try:
        import cairosvg
        cairosvg.svg2png(bytestring=svg_str.encode("utf-8"),
                         write_to=str(png_path), scale=escala)
        CONVERSOR_USADO = "cairosvg"
        return png_path
    except ImportError:
        pass
    except Exception as e:
        print(f"[aviso] cairosvg indisponível ({str(e).splitlines()[0][:70]}); tentando svglib...")
    try:
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPM
        d = svg2rlg(str(svg_path))
        d.scale(escala, escala)
        d.width *= escala
        d.height *= escala
        renderPM.drawToFile(d, str(png_path), fmt="PNG")
        CONVERSOR_USADO = "svglib+reportlab"
        return png_path
    except ImportError:
        pass
    except Exception as e:
        print(f"[aviso] svglib falhou ({str(e).splitlines()[0][:70]}); usando Pillow...")
    try:
        _png_via_pillow(svg_str, png_path, escala)
        CONVERSOR_USADO = "Pillow (fallback embutido)"
        return png_path
    except Exception as e:
        print(f"[aviso] Falha ao converter {svg_path.name}: {e}")
        return None


def numeracao(doc):
    sectpr = doc.sections[1]._sectPr
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:start"), "2")
    cols = sectpr.find(qn("w:cols"))
    if cols is not None:
        cols.addprevious(pg)
    else:
        sectpr.append(pg)
    sec3 = doc.sections[2]
    hdr = sec3.header
    try:
        hdr.is_linked_to_previous = False
    except Exception:
        pass
    par = hdr.paragraphs[0]
    par.alignment = DIR
    par.paragraph_format.first_line_indent = Cm(0)
    _add_campo(par, "PAGE", cache="2", tamanho=10)


def ativar_atualizacao_campos(doc):
    el = OxmlElement("w:updateFields")
    el.set(qn("w:val"), "true")
    doc.settings.element.append(el)


# ==============================================================================
# FIGURAS EM SVG
# ==============================================================================
def _svg_open(w, h):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" '
            f'viewBox="0 0 {w} {h}">')


def _txt(x, y, s, tam=12, cor="#222222", ancora="middle", peso="normal", italico=False):
    st = ' font-style="italic"' if italico else ""
    return (f'<text x="{x:.1f}" y="{y:.1f}" font-family="Helvetica, Arial, sans-serif" '
            f'font-size="{tam}" fill="{cor}" text-anchor="{ancora}" '
            f'font-weight="{peso}"{st}>{escape(s)}</text>')


def _rect(x, y, w, h, fill, rx=0, stroke=None, sw=1):
    s = f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" fill="{fill}"'
    if rx:
        s += f' rx="{rx}"'
    if stroke:
        s += f' stroke="{stroke}" stroke-width="{sw}"'
    return s + "/>"


def _circle(cx, cy, r, fill, stroke=None, sw=1):
    s = f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" fill="{fill}"'
    if stroke:
        s += f' stroke="{stroke}" stroke-width="{sw}"'
    return s + "/>"


def _line(x1, y1, x2, y2, cor, w=1, dash=None):
    d = f' stroke-dasharray="{dash}"' if dash else ""
    return (f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
            f'stroke="{cor}" stroke-width="{w}"{d}/>')


def svg_ciclo_vida():
    W, H = 880, 235
    p = [_svg_open(W, H), _rect(0, 0, W, H, "#ffffff")]
    etapas = [
        ("Iniciação", "2 processos", "#2e5d34"),
        ("Planejamento", "24 processos", "#3d7a46"),
        ("Execução", "10 processos", "#4f8a5b"),
        ("Monitoramento\ne Controle", "12 processos", "#67995f"),
        ("Encerramento", "1 processo", "#86b57a"),
    ]
    x0, y0, w, h, tip, step = 28, 78, 180, 76, 26, 164
    for i, (nome, sub, cor) in enumerate(etapas):
        x = x0 + i * step
        cx = x + (w - tip) / 2 + (tip / 2 if i > 0 else 0)
        topo = f"{x},{y0} {x + w - tip},{y0} {x + w},{y0 + h / 2}"
        base = f"{x + w - tip},{y0 + h} {x},{y0 + h}"
        pontos = topo + " " + base if i == 0 else topo + " " + base + f" {x + tip},{y0 + h / 2}"
        p.append(f'<polygon points="{pontos}" fill="{cor}"/>')
        linhas = nome.split("\n")
        if len(linhas) == 1:
            p.append(_txt(cx, y0 + h / 2 + 5, linhas[0], 14.5, "#ffffff", peso="bold"))
        else:
            p.append(_txt(cx, y0 + h / 2 - 3, linhas[0], 14.5, "#ffffff", peso="bold"))
            p.append(_txt(cx, y0 + h / 2 + 15, linhas[1], 14.5, "#ffffff", peso="bold"))
        p.append(_txt(cx, y0 + h + 22, sub, 11.5, "#444444", italico=True))
    p.append("</svg>")
    return "".join(p)


def svg_matriz_stakeholders():
    W, H = 640, 470
    x0, y0, pw, ph = 90, 40, 480, 330
    p = [_svg_open(W, H), _rect(0, 0, W, H, "#ffffff")]
    quads = [
        (x0, y0, pw / 2, ph / 2, "#dbe9f7"),
        (x0 + pw / 2, y0, pw / 2, ph / 2, "#c5d9f0"),
        (x0, y0 + ph / 2, pw / 2, ph / 2, "#f2f2ee"),
        (x0 + pw / 2, y0 + ph / 2, pw / 2, ph / 2, "#dde9dd"),
    ]
    for qx, qy, qw, qh, cor in quads:
        p.append(_rect(qx, qy, qw, qh, cor, stroke="#ffffff", sw=2))
    p.append(_rect(x0, y0, pw, ph, "none", stroke="#999999", sw=1.2))
    p.append(_txt(x0 + pw * 0.25, y0 + 20, "MANTER SATISFEITO", 11.5, "#3a5a7a", peso="bold"))
    p.append(_txt(x0 + pw * 0.75, y0 + 20, "GERIR DE PERTO", 11.5, "#1d4d7a", peso="bold"))
    p.append(_txt(x0 + pw * 0.25, y0 + ph - 12, "MONITORAR", 11.5, "#777777", peso="bold"))
    p.append(_txt(x0 + pw * 0.75, y0 + ph - 12, "MANTER INFORMADO", 11.5, "#3f6d46", peso="bold"))
    p.append(_txt(x0, y0 + ph + 20, "Baixo", 10.5, "#555555"))
    p.append(_txt(x0 + pw, y0 + ph + 20, "Alto", 10.5, "#555555"))
    p.append(_txt(x0 + pw, y0 + ph + 38, "INTERESSE →", 12, "#333333", "end", "bold"))
    p.append(_txt(x0 - 10, y0 - 14, "PODER ↑", 12, "#333333", "start", "bold"))
    interessados = [
        ("Proprietária", 9.0, 8.8),
        ("Recepção", 8.0, 5.2),
        ("Pacientes atuais", 7.0, 2.2),
        ("Pacientes potenciais", 9.2, 1.6),
        ("Agência de mídia", 5.0, 5.6),
        ("CRO-SC", 3.0, 6.4),
        ("Hospedagem", 2.0, 3.0),
    ]
    for nome, i_, p_ in interessados:
        cx = x0 + i_ / 10 * pw
        cy = y0 + ph - p_ / 10 * ph
        p.append(_circle(cx, cy, 6.5, "#1d4d7a", stroke="#ffffff", sw=1.5))
        if i_ >= 5:
            p.append(_txt(cx - 10, cy + 4, nome, 11, "#222222", "end"))
        else:
            p.append(_txt(cx + 10, cy + 4, nome, 11, "#222222", "start"))
    p.append("</svg>")
    return "".join(p)


def svg_eap():
    grupos = [
        ("1 Gestão do Projeto", ["1.1 Briefing e termo de abertura",
                                 "1.2 Gerenciamento e controle",
                                 "1.3 Encerramento e transição"]),
        ("2 Descoberta", ["2.1 Entrevistas e concorrência",
                          "2.2 Personas e jornada"]),
        ("3 UX/UI", ["3.1 Wireframes e arquitetura",
                     "3.2 Design de alta fidelidade"]),
        ("4 Desenvolvimento", ["4.1 Front-end responsivo",
                               "4.2 CMS, formulários e CRM"]),
        ("5 Conteúdo", ["5.1 Copywriting",
                        "5.2 Ensaio fotográfico"]),
        ("6 Lançamento", ["6.1 SEO e performance",
                          "6.2 Testes e go-live"]),
    ]
    cw, pad, gap = 5.9, 18, 26
    larguras = [max([len(t) + 2] + [len(f) + 3 for f in fs]) * cw + pad for t, fs in grupos]
    W = sum(larguras) + gap * (len(grupos) - 1) + 40
    max_folhas = max(len(fs) for _, fs in grupos)
    H = 160 + max_folhas * 38 + 14
    p = [_svg_open(W, H), _rect(0, 0, W, H, "#ffffff")]
    cx_root = W / 2
    p.append(_rect(cx_root - 190, 24, 380, 40, "#1d4d7a", rx=6))
    p.append(_txt(cx_root, 49, "LANDING PAGE — CLÍNICA ODONTOLÓGICA SORRISO REAL", 13, "#ffffff", peso="bold"))
    col_x = 20
    centros = []
    for w_ in larguras:
        centros.append(col_x + w_ / 2)
        col_x += w_ + gap
    p.append(_line(cx_root, 64, cx_root, 86, "#555555", 1.5))
    p.append(_line(centros[0], 86, centros[-1], 86, "#555555", 1.5))
    for idx, (titulo, folhas) in enumerate(grupos):
        x = centros[idx] - larguras[idx] / 2
        cx = centros[idx]
        p.append(_line(cx, 86, cx, 102, "#555555", 1.5))
        p.append(_rect(x, 102, larguras[idx], 36, "#3b6ea5", rx=5))
        p.append(_txt(cx, 124, titulo, 12, "#ffffff", peso="bold"))
        p.append(_line(cx, 138, cx, 160 + (len(folhas) - 1) * 38 + 14, "#3b6ea5", 1.5))
        y = 160
        for folha in folhas:
            p.append(_line(cx, y + 14, x + 16, y + 14, "#3b6ea5", 1.5))
            p.append(_rect(x + 16, y, larguras[idx] - 16, 28, "#e4edf8", rx=4,
                           stroke="#3b6ea5", sw=1))
            p.append(_txt(x + 16 + (larguras[idx] - 16) / 2, y + 18.5, folha, 10.5, "#17334d"))
            y += 38
    p.append("</svg>")
    return "".join(p)


def svg_gantt():
    tarefas = [
        ("1.1", "Briefing e termo de abertura", 1, 1, "G"),
        ("1.2", "Gerenciamento e reuniões de status", 1, 10, "G"),
        ("1.3", "Encerramento e transição", 10, 1, "G"),
        ("2.1", "Entrevistas e concorrência", 1, 2, "D"),
        ("2.2", "Personas e jornada do paciente", 2, 2, "D"),
        ("M1", "Marco — Aprovação da estratégia (21/03)", 3, None, "X"),
        ("3.1", "Wireframes e arquitetura", 3, 2, "U"),
        ("3.2", "Design de alta fidelidade", 4, 3, "U"),
        ("M2", "Marco — Aprovação do design (11/04)", 6, None, "X"),
        ("4.1", "Desenvolvimento front-end", 6, 3, "F"),
        ("4.2", "CMS, formulários e integração ao CRM", 8, 2, "F"),
        ("5.1", "Copywriting e conteúdo", 5, 3, "C"),
        ("5.2", "Ensaio fotográfico", 6, 1, "C"),
        ("6.1", "SEO, performance e acessibilidade", 8, 2, "S"),
        ("6.2", "Testes, QA e publicação", 9, 1, "T"),
        ("M3", "Marco — Go-live (02/05/2025)", 9, None, "X"),
    ]
    cores = {"G": "#3b6ea5", "D": "#7a5ea8", "U": "#e07b39", "F": "#2e5d34",
             "C": "#b5651d", "S": "#7a9e3b", "T": "#c0504d", "X": "#8a2f2b"}
    cor_m = cores["X"]
    cx0, cw = 268, 606
    ww = cw / 10
    row_y0, rh = 64, 25
    H = row_y0 + len(tarefas) * rh + 78
    W = cx0 + cw + 16
    p = [_svg_open(W, H), _rect(0, 0, W, H, "#ffffff")]
    for idx in range(len(tarefas)):
        if idx % 2 == 0:
            p.append(_rect(cx0, row_y0 + idx * rh, cw, rh, "#f6f6f3"))
    for b in (0, 4, 8, 10):  # limites de mês: mar | abr | mai
        x = cx0 + b * ww
        p.append(_line(x, 48, x, row_y0 + len(tarefas) * rh + 6, "#dddddd", 1))
    for centro_semana, nome in ((2, "Mar"), (6, "Abr"), (9, "Mai")):
        p.append(_txt(cx0 + centro_semana * ww, 38, nome, 11.5, "#333333", peso="bold"))
    p.append(_line(cx0, 48, cx0 + cw, 48, "#999999", 1.2))
    for idx, (cod, nome, inicio, dur, tipo) in enumerate(tarefas):
        y = row_y0 + idx * rh
        p.append(_txt(8, y + 16.5, cod, 10.5, cores[tipo], "start", "bold"))
        p.append(_txt(48, y + 16.5, nome, 10.5, "#222222", "start"))
        if tipo == "X":
            cx, cy = cx0 + inicio * ww, y + rh / 2
            p.append(f'<polygon points="{cx:.1f},{cy - 8:.1f} {cx + 8:.1f},{cy:.1f} '
                     f'{cx:.1f},{cy + 8:.1f} {cx - 8:.1f},{cy:.1f}" fill="{cor_m}"/>')
        else:
            bx = cx0 + (inicio - 1) * ww
            bw = dur * ww - 4
            p.append(_rect(bx, y + 5.5, bw, 14, cores[tipo], rx=3))
            rot = "contínuo" if dur >= 10 else f"{dur} sem"
            p.append(_txt(bx + bw + 5, y + 16, rot, 9, "#555555", "start"))
    itens_legenda = [("G", "Gestão do projeto"), ("D", "Descoberta"), ("U", "UX/UI"),
                     ("F", "Desenvolvimento"), ("C", "Conteúdo"), ("S", "Lançamento/SEO"),
                     ("X", "Marco")]
    ly = row_y0 + len(tarefas) * rh + 22
    for i, (t_, nome_) in enumerate(itens_legenda):
        x = cx0 + (i % 4) * 150
        y = ly + (i // 4) * 20
        p.append(_rect(x, y, 12, 12, cores[t_], rx=2))
        p.append(_txt(x + 18, y + 10, nome_, 10.5, "#222222", "start"))
    p.append("</svg>")
    return "".join(p)


def svg_matriz_riscos():
    cel, gx, gy = 76, 100, 64
    W, H = gx + 5 * cel + 40, gy + 5 * cel + 118
    riscos = [("R1", 4, 3), ("R2", 4, 2), ("R3", 2, 4),
              ("R4", 2, 5), ("R5", 1, 4), ("R6", 2, 3)]
    p = [_svg_open(W, H), _rect(0, 0, W, H, "#ffffff")]
    for imp in range(1, 6):
        for prob in range(1, 6):
            x, y = gx + (imp - 1) * cel, gy + (5 - prob) * cel
            esc = imp * prob
            cor = ("#a1c181" if esc <= 4 else "#f0d264" if esc <= 9
                   else "#e8975a" if esc <= 14 else "#cf5a4e")
            p.append(_rect(x, y, cel, cel, cor, stroke="#ffffff", sw=2))
    palavras = ["Muito baixo", "Baixo", "Moderado", "Alto", "Muito alto"]
    for imp in range(1, 6):
        x = gx + (imp - 1) * cel + cel / 2
        p.append(_txt(x, gy + 5 * cel + 20, str(imp), 12, "#333333", peso="bold"))
        p.append(_txt(x, gy + 5 * cel + 34, palavras[imp - 1], 9, "#555555"))
    for prob in range(1, 6):
        y = gy + (5 - prob) * cel + cel / 2 + 4
        p.append(_txt(gx - 14, y, str(prob), 12, "#333333", peso="bold"))
    p.append(_txt(gx + 2.5 * cel, gy - 16, "PROBABILIDADE ↑", 12, "#333333", peso="bold"))
    p.append(_txt(gx + 2.5 * cel, gy + 5 * cel + 56, "IMPACTO →", 12, "#333333", peso="bold"))
    por_celula = defaultdict(list)
    for rid, imp, prob in riscos:
        por_celula[(imp, prob)].append(rid)
    for (imp, prob), ids in por_celula.items():
        cx = gx + (imp - 1) * cel + cel / 2
        cy = gy + (5 - prob) * cel + cel / 2
        for k, rid in enumerate(ids):
            dx = (k - (len(ids) - 1) / 2) * 34
            p.append(_circle(cx + dx, cy, 14, "#ffffff", stroke="#333333", sw=1.5))
            p.append(_txt(cx + dx, cy + 4, rid, 11, "#222222", peso="bold"))
    y = gy + 5 * cel + 76
    for i, (c_, t_) in enumerate([("#a1c181", "Baixo (1–4)"), ("#f0d264", "Moderado (5–9)"),
                                  ("#e8975a", "Alto (10–14)"), ("#cf5a4e", "Crítico (15–25)")]):
        x = gx - 4 + i * 106
        p.append(_rect(x, y, 12, 12, c_, rx=2))
        p.append(_txt(x + 17, y + 10, t_, 10, "#222222", "start"))
    p.append("</svg>")
    return "".join(p)


def svg_curva_s():
    W, H = 840, 470
    x0, y0, pw, ph = 84, 44, 690, 340
    y1 = y0 + ph
    PV = [2.5, 4.5, 7.0, 9.0, 11.5, 13.5, 15.5, 17.2, 18.5, 19.25]
    EV = [2.3, 4.2, 6.5, 8.2, 10.5, 12.4, 14.6, 16.6, 18.2, 19.25]
    AC = [2.5, 4.6, 7.2, 9.5, 12.0, 14.2, 16.0, 17.6, 18.6, 19.10]
    ymax = 20

    def X(i):
        return x0 + i / 9 * pw

    def Y(v):
        return y1 - v / ymax * ph

    p = [_svg_open(W, H), _rect(0, 0, W, H, "#ffffff")]
    for v in range(0, 21, 5):
        p.append(_line(x0, Y(v), x0 + pw, Y(v), "#e2e2e2", 1))
        p.append(_txt(x0 - 10, Y(v) + 4, str(v), 10.5, "#555555", "end"))
    p.append(_txt(x0 - 58, y0 + 4, "R$ mil", 10, "#555555"))
    for semana, rotulo in ((0, "Sem 1"), (2, "Sem 3"), (4, "Sem 5"), (6, "Sem 7"), (8, "Sem 9")):
        p.append(_txt(X(semana), y1 + 20, rotulo, 11, "#333333"))
    p.append(_line(x0, y1, x0 + pw, y1, "#999999", 1.2))
    p.append(_line(x0, y0 - 6, x0, y1, "#999999", 1.2))
    p.append(_txt(x0 + pw / 2, y1 + 44, "Semanas de 2025 (março a maio)", 11.5, "#333333"))
    p.append(_line(x0, Y(19.25), x0 + pw, Y(19.25), "#aaaaaa", 1, dash="6,5"))
    p.append(_txt(x0 + 8, Y(19.25) - 6, "BAC = R$ 19.250", 10, "#666666", "start"))
    xd = X(5)
    p.append(_line(xd, y0 - 6, xd, y1, "#777777", 1.2, dash="5,5"))
    p.append(_txt(xd + 8, y0 + 4, "data de dados (11/04/2025)", 10, "#555555", "start", italico=True))
    series = [("PV — planejado", PV, "#1d4d7a"),
              ("EV — valor agregado", EV, "#7a5ea8"),
              ("AC — custo real", AC, "#c0504d")]
    for nome_, dados, cor in series:
        pts = " ".join(f"{X(i):.1f},{Y(v):.1f}" for i, v in enumerate(dados))
        p.append(f'<polyline points="{pts}" fill="none" stroke="{cor}" stroke-width="3"/>')
        for i, v in enumerate(dados):
            p.append(_circle(X(i), Y(v), 3.2, cor))
    for i, (nome_, dados, cor) in enumerate(series):
        yy = y0 + 16 + i * 20
        p.append(_line(x0 + 24, yy, x0 + 50, yy, cor, 3))
        p.append(_txt(x0 + 56, yy + 4, nome_, 10.5, "#222222", "start"))
    p.append("</svg>")
    return "".join(p)


# ==============================================================================
# CONTEÚDO DO DOCUMENTO
# ==============================================================================
RESUMO = ("Este documento apresenta o planejamento e a gestão do projeto Landing Page "
          "da Clínica Odontológica Sorriso Real, iniciativa digital voltada a converter "
          "visitantes em pacientes agendados. A iniciativa foi estruturada segundo as "
          "boas práticas do Guia PMBOK (PMI, 2017), percorrendo os cinco grupos de "
          "processos — iniciação, planejamento, execução, monitoramento e controle e "
          "encerramento — e as dez áreas de conhecimento. Na iniciação, formalizou-se o "
          "termo de abertura e a matriz de interessados; no planejamento, elaboraram-se "
          "a declaração de escopo, a estrutura analítica do projeto, o cronograma de 10 "
          "semanas, o orçamento de R$ 19.250,00 e os planos de qualidade, comunicações "
          "e riscos, com atenção à LGPD e às normas de publicidade odontológica; na "
          "execução, conduziram-se descoberta, design, desenvolvimento e conteúdo; no "
          "monitoramento, aplicou-se o gerenciamento de valor agregado para corrigir "
          "desvios de prazo e custo; e, no encerramento, formalizou-se o aceite, a "
          "transição e as lições aprendidas. Como resultado, espera-se publicar a página "
          "em 02/05/2025 com pontuação Lighthouse igual ou superior a 90 e ampliar em "
          "30% os agendamentos digitais da clínica em 90 dias.")

ABSTRACT = ("This document presents the planning and management of the Landing Page "
            "Project for the Sorriso Real Dental Clinic, a digital initiative aimed at "
            "converting visitors into booked patients. The work is structured according "
            "to the PMBOK Guide best practices, covering the five process groups and the "
            "ten knowledge areas. Deliverables include the project charter, the work "
            "breakdown structure, a 10-week schedule, a BRL 19,250 budget, and quality, "
            "communications and risk plans, with attention to data-protection and "
            "dental-advertising regulations. Earned value management supported "
            "corrective actions, and the project was formally closed with stakeholder "
            "acceptance and lessons learned.")


def capa(doc):
    linha_vazia(doc, 2)
    centro(doc, "CENTRO UNIVERSITÁRIO VERDE VALE – CUVV", negrito=True)
    linha_vazia(doc, 1)
    centro(doc, "MARIA ALVES DA SILVA")
    linha_vazia(doc, 5)
    centro(doc, "LANDING PAGE DA CLÍNICA ODONTOLÓGICA SORRISO REAL", negrito=True)
    centro(doc, "Gerenciamento segundo o Guia PMBOK: da iniciação ao encerramento")
    linha_vazia(doc, 12)
    centro(doc, "Blumenau – SC")
    centro(doc, "2025")


def folha_de_rosto(doc):
    linha_vazia(doc, 2)
    centro(doc, "MARIA ALVES DA SILVA")
    linha_vazia(doc, 4)
    centro(doc, "LANDING PAGE DA CLÍNICA ODONTOLÓGICA SORRISO REAL", negrito=True)
    centro(doc, "Gerenciamento segundo o Guia PMBOK: da iniciação ao encerramento")
    linha_vazia(doc, 4)
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.left_indent = Cm(8); pf.first_line_indent = Cm(0); pf.alignment = JUS
    par.add_run("Projeto técnico apresentado ao Curso de Sistemas de Informação do "
                "Centro Universitário Verde Vale (CUVV), como requisito parcial da "
                "disciplina de Gerenciamento de Projetos.")
    linha_vazia(doc, 1)
    par2 = doc.add_paragraph()
    pf2 = par2.paragraph_format
    pf2.left_indent = Cm(8); pf2.first_line_indent = Cm(0); pf2.alignment = JUS
    par2.add_run("Orientadora: Profa. Dra. Ana Beatriz Rocha")
    linha_vazia(doc, 8)
    centro(doc, "Blumenau – SC")
    centro(doc, "2025")


def pre_textual(doc):
    titulo_secao(doc, "RESUMO")
    paragrafo(doc, RESUMO, recuo=False)
    linha_vazia(doc, 1)
    pk = doc.add_paragraph()
    pk.paragraph_format.first_line_indent = Cm(0)
    pk.add_run("Palavras-chave: ").bold = True
    pk.add_run("gerenciamento de projetos; PMBOK; landing page; experiência do usuário; clínica odontológica.")

    titulo_secao(doc, "ABSTRACT")
    paragrafo(doc, ABSTRACT, recuo=False)
    linha_vazia(doc, 1)
    kw = doc.add_paragraph()
    kw.paragraph_format.first_line_indent = Cm(0)
    kw.add_run("Keywords: ").bold = True
    kw.add_run("project management; PMBOK; landing page; user experience; dental clinic.")

    titulo_secao(doc, "LISTA DE FIGURAS")
    add_toc(doc, 'TOC \\h \\z \\c "Figura"',
            "Atualize este campo (F9 ou reabra o arquivo) para gerar a lista de figuras.")
    titulo_secao(doc, "LISTA DE TABELAS")
    add_toc(doc, 'TOC \\h \\z \\c "Tabela"',
            "Atualize este campo (F9 ou reabra o arquivo) para gerar a lista de tabelas.")
    titulo_secao(doc, "SUMÁRIO")
    add_toc(doc, 'TOC \\o "1-3" \\h \\z \\u',
            "Atualize este campo (F9 ou reabra o arquivo) para gerar o sumário.")


# ------------------------------------------------------------------------------
# Capítulos
# ------------------------------------------------------------------------------
def introducao(doc):
    secao(doc, 1, "1 INTRODUÇÃO")
    paragrafo(doc, "A decisão de contratar serviços de saúde começa, hoje, na busca "
        "digital: pacientes pesquisam no Google, comparam avaliações e avaliam o site "
        "da clínica em segundos antes de agendar. Para clínicas de pequeno porte, "
        "páginas institucionais genéricas convertem mal — o visitante chega, não "
        "encontra um caminho claro de agendamento e abandona a navegação, desperdiçando "
        "o investimento feito em mídia paga.")
    paragrafo(doc, "Este trabalho apresenta o projeto Landing Page da Clínica "
        "Odontológica Sorriso Real, que desenvolve e publica uma página responsiva "
        "focada em conversão: proposta de valor, prova social, formulário integrado ao "
        "CRM da clínica, botão de WhatsApp e métricas de rastreamento. O prazo de "
        "execução é de 10 semanas (3 de março a 9 de maio de 2025, com encerramento "
        "administrativo até 16/05) e o orçamento aprovado é de R$ 19.250,00.")
    paragrafo(doc, "Segundo o Project Management Institute, um projeto é um esforço "
        "temporário empreendido para criar um produto, serviço ou resultado único "
        "(PMI, 2017). A disciplina aplicada neste caso fundamenta-se no Guia PMBOK, "
        "cuja definição sintetiza a abrangência da prática:")
    citacao_longa(doc, "Gerenciamento de projetos é a aplicação de conhecimento, "
        "habilidades, ferramentas e técnicas às atividades do projeto para atender aos "
        "seus requisitos, o que envolve a identificação dos requisitos, o atendimento "
        "das necessidades e expectativas das partes interessadas e o balanceamento das "
        "restrições concorrentes de escopo, prazo, custo, qualidade, recursos e risco "
        "(PMI, 2017, p. 10).")
    paragrafo(doc, "A sétima edição do guia (PMI, 2021) reorganizou o corpo de "
        "conhecimento em torno de princípios e domínios de desempenho; a estrutura de "
        "grupos de processos da sexta edição, porém, permanece como referência "
        "operacional amplamente adotada e orienta este documento, conforme a Figura 1.")
    figura(doc, svg_ciclo_vida, "ciclo_vida",
           "Grupos de processos do gerenciamento de projetos segundo o Guia PMBOK "
           "(6ª ed.)", "adaptada de Project Management Institute (2017).", 14.0)
    paragrafo(doc, "O documento está organizado pelos cinco grupos de processos: "
        "iniciação (capítulo 4), planejamento (capítulo 5), execução (capítulo 6), "
        "monitoramento e controle (capítulo 7) e encerramento (capítulo 8), seguidos "
        "das considerações finais e das referências. As figuras e tabelas são de "
        "elaboração do autor e a formatação segue a ABNT NBR 14724.")


def objetivos(doc):
    secao(doc, 1, "2 OBJETIVOS")
    secao(doc, 2, "2.1 Objetivo geral")
    paragrafo(doc, "Planejar, projetar e publicar, em 10 semanas, uma landing page "
        "responsiva para a Clínica Odontológica Sorriso Real, integrada ao CRM e às "
        "normas aplicáveis, dentro do orçamento aprovado de R$ 19.250,00, para converter "
        "visitantes em pacientes agendados.")
    secao(doc, 2, "2.2 Objetivos específicos")
    for item in [
        "estabelecer linhas de base de escopo, prazo e custo e mantê-las sob controle integrado de mudanças;",
        "aprovar personas, wireframes e design de alta fidelidade até 11/04/2025;",
        "publicar a página em 02/05/2025 com pontuação Lighthouse igual ou superior a 90 nas quatro categorias;",
        "integrar formulário e WhatsApp ao CRM, com o lead disponível para a recepção em menos de 1 minuto;",
        "garantir conformidade com a LGPD (BRASIL, 2018) e com a Resolução CFO-159/2016 (CFO, 2016);",
        "treinar a recepção e encerrar o projeto com aceite formal até 16/05/2025;",
        "contribuir para o aumento de 30% nos agendamentos digitais em 90 dias (KPI pós-projeto).",
    ]:
        marcador(doc, item)


def justificativa(doc):
    secao(doc, 1, "3 JUSTIFICATIVA")
    paragrafo(doc, "A clínica capta pacientes quase exclusivamente por indicação e "
        "pelo perfil no Google Meu Negócio; a agenda de terças e quartas opera ociosa, "
        "enquanto o investimento em anúncios encaminha visitantes para um site "
        "institucional sem foco de conversão. Uma landing page dedicada concentra a "
        "atenção do visitante em um único objetivo — agendar a primeira consulta — e "
        "torna o retorno da mídia paga mensurável.")
    paragrafo(doc, "Do ponto de vista de experiência, os princípios de usabilidade "
        "consagrados por Krug (2015) — clareza da proposta, hierarquia visual e mínimo "
        "esforço de decisão — orientam o desenho da página, enquanto as diretrizes "
        "técnico-gerenciais de Kerzner (2017) sustentam a disciplina de execução. A "
        "combinação entre boas práticas de gerenciamento e de design é o que garante "
        "que o produto final converta, e não apenas exista.")
    paragrafo(doc, "No plano institucional, o projeto permite à proprietária medir o "
        "custo por lead, reduzir o tempo de resposta ao paciente (integração com CRM e "
        "WhatsApp) e servir de base para futuras páginas por especialidade, como "
        "implante e ortodontia, em uma arquitetura digital escalável.")


def iniciacao(doc):
    secao(doc, 1, "4 INICIAÇÃO DO PROJETO")
    paragrafo(doc, "O grupo de processos de iniciação define o novo projeto obtendo "
        "autorização formal para seu início. Nesta fase foram elaborados o termo de "
        "abertura do projeto e o registro das partes interessadas, com apoio direto da "
        "patrocinadora.")

    secao(doc, 2, "4.1 Termo de abertura do projeto")
    paragrafo(doc, "O termo de abertura do projeto (TAP) autoriza formalmente a "
        "existência do projeto e confere ao gerente autoridade para aplicar recursos. "
        "Os elementos centrais do TAP são apresentados a seguir.")
    paragrafo_rotulo(doc, "Nome do projeto:", "Landing Page da Clínica Odontológica Sorriso Real.")
    paragrafo_rotulo(doc, "Patrocinadora:", "Dra. Helena Prado, proprietária da clínica.")
    paragrafo_rotulo(doc, "Gerente do projeto:", "Maria Alves da Silva, designada em "
        "05/03/2025, com autoridade para aprovar mudanças de até R$ 1.000,00 e "
        "contratar os serviços previstos no orçamento.")
    paragrafo_rotulo(doc, "Justificativa:", "sem página de conversão, o investimento em "
        "mídia paga desperdiça alcance; a página dedicada aumenta o volume e a "
        "qualidade dos leads e permite mensurar o retorno da comunicação digital.")
    paragrafo_rotulo(doc, "Objetivos e critérios de sucesso:", "publicar em 02/05/2025 "
        "uma página com Lighthouse igual ou superior a 90, formulário integrado ao CRM "
        "testado ponta a ponta, custo final igual ou inferior a R$ 19.250,00, "
        "satisfação da patrocinadora igual ou superior a 8/10 e conformidade LGPD/CRO.")
    paragrafo_rotulo(doc, "Escopo preliminar:", "descoberta e personas, UX/UI, "
        "desenvolvimento front-end, integrações, conteúdo e fotografia, SEO técnico, "
        "testes, go-live e treinamento da recepção.")
    paragrafo_rotulo(doc, "Orçamento sumário:", "R$ 19.250,00, recursos próprios da "
        "clínica, liberados em duas parcelas (60% em março; 40% em abril).")
    paragrafo_rotulo(doc, "Riscos de alto nível:", "mudanças frequentes de direção "
        "visual, não conformidade com LGPD ou normas de publicidade odontológica, "
        "atraso de conteúdo e falhas na integração ao CRM (detalhados na seção 5.7).")
    paragrafo_rotulo(doc, "Data de aprovação:", "07/03/2025, com assinatura da "
        "proprietária.")
    paragrafo(doc, "Os marcos principais aprovados no TAP são:")
    for item in [
        "07/03/2025 – aprovação do termo de abertura;",
        "21/03/2025 – aprovação de personas e estratégia (M1);",
        "11/04/2025 – aprovação do design de alta fidelidade (M2);",
        "02/05/2025 – publicação da página (go-live) (M3);",
        "16/05/2025 – aceite final e encerramento.",
    ]:
        marcador(doc, item)

    secao(doc, 2, "4.2 Interessados e matriz poder × interesse")
    paragrafo(doc, "A identificação dos interessados utilizou entrevista com a "
        "proprietária e com a recepção, análise do fluxo atual de agendamentos e "
        "levantamento de exigências regulatórias (LGPD e normas do CRO). Cada "
        "interessado foi posicionado quanto ao poder de interferir no projeto e ao "
        "grau de interesse, originando a Tabela 1 e a Figura 2.")
    tabela(doc,
           "Registro dos principais interessados do projeto",
           "elaborada pelo autor (2025).",
           ["Interessado", "Poder", "Interesse", "Estratégia de engajamento"],
           [
               ["Dra. Helena Prado (proprietária e patrocinadora)", "Alto", "Alto",
                "Gerir de perto: reunião semanal de status e decisões em até 48 h"],
               ["Equipe de recepção (secretariado)", "Médio", "Alto",
                "Manter informado: coautoria do formulário e treinamento no CRM"],
               ["Pacientes atuais da clínica", "Baixo", "Alto",
                "Manter informado: convites para testes de usabilidade"],
               ["Pacientes potenciais (público-alvo)", "Baixo", "Alto",
                "Manter informado: pesquisa de personas e testes remotos"],
               ["Agência de mídia paga parceira", "Médio", "Médio",
                "Monitorar: alinhamento de rastreamento (GA4 e pixel)"],
               ["Conselho de odontologia (CRO-SC)", "Médio", "Baixo",
                "Manter satisfeito: conformidade com a Resolução CFO-159/2016"],
               ["Fornecedor de hospedagem e domínio", "Baixo", "Baixo",
                "Monitorar: SLA de disponibilidade na janela de go-live"],
           ],
           larguras=[4.6, 1.6, 2.0, 6.8], alinhar_centro={1, 2})
    figura(doc, svg_matriz_stakeholders, "matriz_stakeholders",
           "Matriz de interessados — poder × interesse",
           "elaborada pelo autor (2025).", 12.0)
    paragrafo(doc, "A matriz evidencia que a proprietária concentra poder e interesse "
        "devendo ser gerida de perto, com decisões em tempo curto; a recepção e os "
        "pacientes, de alto interesse e menor poder, exigem informação e participação "
        "nas decisões de interface; e exigências regulatórias do CRO, ainda que de "
        "baixo interesse corrente, são tratadas como restrição de conformidade.")

    secao(doc, 2, "4.3 Premissas e restrições")
    paragrafo(doc, "Premissas aceitas na iniciação:")
    for item in [
        "logotipo, cores e materiais de marca entregues pela clínica na semana 1;",
        "decisões de aprovação da proprietária em até 48 horas;",
        "acesso ao Google Meu Negócio, Analytics e CRM concedido na semana 1;",
        "hospedagem com SSL e domínio já registrados e disponíveis;",
        "revisão clínica do conteúdo realizada pela dentista em até 3 dias.",
    ]:
        marcador(doc, item)
    paragrafo(doc, "Restrições identificadas:")
    for item in [
        "orçamento teto de R$ 19.250,00, sem aportes adicionais previstos;",
        "comunicação odontológica conforme a Resolução CFO-159/2016, vedadas promessas de resultados (CFO, 2016);",
        "tratamento de dados pessoais conforme a LGPD (BRASIL, 2018), com consentimento explícito no formulário;",
        "escopo limitado a uma landing page e à página de política de privacidade;",
        "go-live somente em dia útil, às 7h, em janela de baixo tráfego.",
    ]:
        marcador(doc, item)


def planejamento(doc):
    secao(doc, 1, "5 PLANEJAMENTO DO PROJETO")
    paragrafo(doc, "No planejamento foram elaborados os planos auxiliares e as linhas "
        "de base de escopo, prazo e custo, aprovados pela patrocinadora em 14/03/2025 e "
        "mantidos sob controle integrado de mudanças.")

    secao(doc, 2, "5.1 Declaração de escopo")
    paragrafo(doc, "O escopo descreve o produto e o trabalho necessário para entregá-lo. "
        "Estão incluídas as seguintes entregas:")
    for item in [
        "estratégia validada: três personas, jornada do paciente e proposta de valor;",
        "wireframes navegáveis e design de alta fidelidade (desktop e mobile);",
        "landing page responsiva (uma página + política de privacidade);",
        "formulário com validação, consentimento e integração ao CRM e WhatsApp;",
        "copywriting persuasivo e ensaio fotográfico autoral;",
        "SEO on-page, métricas (GA4 e pixel) e relatório Lighthouse;",
        "testes de usabilidade com cinco pacientes;",
        "treinamento de 2 horas da recepção e documentação de transição.",
    ]:
        marcador(doc, item)
    paragrafo(doc, "Estão explicitamente fora do escopo: portal do paciente; "
        "agendamento online com agenda própria; blog e conteúdo contínuo; campanhas de "
        "mídia paga (responsabilidade da agência parceira); e manutenção além de 90 "
        "dias após o go-live.")
    paragrafo(doc, "Critérios de aceitação: pontuação Lighthouse igual ou superior a "
        "90 nas quatro categorias; formulário testado ponta a ponta com chegada do lead "
        "ao CRM; aprovação formal da proprietária; e checklist LGPD/CRO concluído sem "
        "pendências.")

    secao(doc, 2, "5.2 Estrutura analítica do projeto (EAP)")
    paragrafo(doc, "A decomposição do escopo seguiu a regra dos 100%: a EAP contempla "
        "todo o trabalho do projeto e nada fora dela será executado. A Figura 3 "
        "apresenta os três níveis de decomposição adotados.")
    figura(doc, svg_eap, "eap",
           "Estrutura analítica do projeto (EAP) da landing page da Clínica Sorriso Real",
           "elaborada pelo autor (2025).", 15.5)
    paragrafo(doc, "O dicionário da EAP descreve, para cada pacote, o responsável, as "
        "entregas e os critérios de aceitação. A título de exemplo, o pacote 3.2 (Design "
        "de alta fidelidade) é de responsabilidade da designer contratada, entrega o "
        "arquivo Figma e o guia de estilo, e é aceito com a aprovação da proprietária "
        "em no máximo duas rodadas de revisão.")

    secao(doc, 2, "5.3 Cronograma")
    paragrafo(doc, "As atividades foram sequenciadas por precedências e as durações "
        "estimadas por técnicas análogas, com base em projetos digitais de porte "
        "semelhante. A linha de base do prazo cobre 10 semanas, de março a maio de "
        "2025, com marcos M1 (estratégia), M2 (design) e M3 (go-live), conforme a "
        "Figura 4.")
    figura(doc, svg_gantt, "gantt",
           "Cronograma do projeto (Gantt) — semanas 1 a 10 (março a maio de 2025)",
           "elaborada pelo autor (2025).", 15.0)
    paragrafo(doc, "O caminho crítico atravessa a linha de produto — briefing, "
        "personas, wireframes, design, front-end, integrações, testes e go-live — sem "
        "folga disponível. A frente de conteúdo admite folga de duas semanas e a de "
        "SEO, de uma semana, permitindo absorver atrasos de materiais sem impacto no "
        "marco de publicação.")

    secao(doc, 2, "5.4 Custos e orçamento")
    paragrafo(doc, "As estimativas foram obtidas de baixo para cima, a partir dos "
        "pacotes de trabalho da EAP, com cotações de profissionais e fornecedores. A "
        "Tabela 2 consolida o orçamento aprovado, incluindo reserva de contingência de "
        "10% gerida pelo gerente do projeto.")
    tabela(doc,
           "Orçamento detalhado por categoria de custo",
           "elaborada pelo autor (2025).",
           ["Categoria de custo", "Valor (R$)", "% do orçamento"],
           [
               ["Descoberta e estratégia", "2.000,00", "10,4%"],
               ["Pesquisa, arquitetura e wireframes (UX)", "3.000,00", "15,6%"],
               ["Design de interface (UI)", "3.500,00", "18,1%"],
               ["Desenvolvimento front-end e integrações", "5.500,00", "28,6%"],
               ["Conteúdo, copywriting e fotografia", "2.000,00", "10,4%"],
               ["SEO, performance e acessibilidade", "1.000,00", "5,2%"],
               ["Testes, QA e lançamento", "500,00", "2,6%"],
               ["Reserva de contingência (10%)", "1.750,00", "9,1%"],
               ["TOTAL (linha de base do custo)", "19.250,00", "100%"],
           ],
           larguras=[7.6, 3.0, 3.0], alinhar_centro={1, 2})
    paragrafo(doc, "Os recursos são integralmente próprios da clínica, liberados em "
        "duas parcelas (60% em março e 40% em abril). Hospedagem, domínio e licenças de "
        "fontes estão incluídos na categoria de desenvolvimento. O fluxo de desembolso "
        "planejado, representado pela curva S da Figura 6 (seção 7.1), servirá de linha "
        "de base para o gerenciamento de valor agregado.")

    secao(doc, 2, "5.5 Qualidade")
    paragrafo(doc, "O plano de qualidade define padrões e métricas mensuráveis, "
        "verificados a cada semana:")
    for item in [
        "Lighthouse igual ou superior a 90 em performance, acessibilidade, boas práticas e SEO;",
        "maior elemento de conteúdo (LCP) renderizado em menos de 2,5 s em conexão 4G simulada;",
        "zero defeitos críticos na matriz de testes (5 dispositivos × 3 navegadores);",
        "10 envios de formulário de teste com chegada do lead ao CRM em menos de 1 minuto;",
        "conclusão de pelo menos 80% nas três tarefas do teste de usabilidade com cinco pacientes.",
    ]:
        marcador(doc, item)
    paragrafo(doc, "A garantia da qualidade inclui revisão por pares de código, "
        "checklist WCAG 2.1 nível AA, auditoria de privacidade (consentimento, cookies "
        "e armazenamento de dados), teste de usabilidade com pacientes reais e um ciclo "
        "de QA completo antes da janela de go-live.")

    secao(doc, 2, "5.6 Recursos e comunicações")
    paragrafo(doc, "A equipe é composta pelo gerente do projeto (também desenvolvedor "
        "full-stack, com 50% de dedicação), designer de UI contratada (60 horas), "
        "redator publicitário (30 horas), fotógrafo (uma diária) e analista de SEO (16 "
        "horas). A recepção da clínica participa como especialista de domínio no "
        "desenho do formulário e do fluxo de leads.")
    paragrafo(doc, "O plano de comunicações define quem recebe o quê, quando e por "
        "qual meio, conforme a Tabela 3.")
    tabela(doc,
           "Plano de comunicação do projeto",
           "elaborada pelo autor (2025).",
           ["Público", "Informação", "Frequência", "Meio", "Responsável"],
           [
               ["Proprietária (patrocinadora)", "Status semanal e decisões de aprovação",
                "Semanal", "Reunião online (30 min)", "Gerente do projeto"],
               ["Equipe de recepção", "Fluxo de leads e treinamento do CRM",
                "Nos marcos", "Presencial", "Gerente do projeto"],
               ["Agência de mídia parceira", "Especificações de rastreamento e pixel",
                "Quinzenal", "E-mail", "Desenvolvedor"],
               ["Pacientes (testes de usabilidade)", "Convites e agendamento das sessões",
                "Conforme marco", "WhatsApp", "Designer de UX"],
               ["Fornecedor de hospedagem", "SLA e janela de publicação",
                "Pontual", "Ticket de suporte", "Desenvolvedor"],
           ],
           larguras=[3.6, 4.4, 2.2, 2.6, 2.4])

    secao(doc, 2, "5.7 Riscos")
    paragrafo(doc, "Os riscos foram identificados por brainstorming com a equipe e a "
        "proprietária e por checklist do Guia PMBOK, analisados qualitativamente em "
        "matriz de probabilidade × impacto 5 × 5 e respondidos conforme a Tabela 4. O "
        "registro é revisado semanalmente no monitoramento (seção 7.3).")
    tabela(doc,
           "Registro de riscos com respostas planejadas",
           "elaborada pelo autor (2025).",
           ["ID", "Risco", "P", "I", "Nível", "Resposta e ação preventiva"],
           [
               ["R1", "Alterações frequentes de direção visual (retrabalho)", "4", "3", "12 – Alto",
                "Mitigar: timebox de duas rodadas de revisão e decisões em 48 h"],
               ["R2", "Atraso na entrega de conteúdo e fotos pela clínica", "4", "2", "8 – Moderado",
                "Mitigar: cronograma de conteúdo com datas e banco de imagens reserva"],
               ["R3", "Falhas na integração do formulário ao CRM", "2", "4", "8 – Moderado",
                "Mitigar: protótipo de integração antecipado e testes com dados fictícios"],
               ["R4", "Não conformidade com LGPD ou normas do CRO", "2", "5", "10 – Alto",
                "Mitigar: checklist jurídico, consentimento explícito e revisão da dentista"],
               ["R5", "Indisponibilidade da hospedagem na janela de go-live", "1", "4", "4 – Baixo",
                "Mitigar: publicação em horário de baixo tráfego e plano de rollback"],
               ["R6", "Baixa performance prejudicando SEO e experiência", "2", "3", "6 – Moderado",
                "Mitigar: imagens otimizadas e teste Lighthouse a cada entrega"],
           ],
           larguras=[1.1, 5.2, 0.9, 0.9, 2.0, 5.0], alinhar_centro={0, 2, 3, 4})
    figura(doc, svg_matriz_riscos, "matriz_riscos",
           "Matriz de probabilidade × impacto (5 × 5) com os riscos do registro",
           "elaborada pelo autor (2025).", 11.0)
    paragrafo(doc, "A reserva de contingência (R$ 1.750,00) cobre as ameaças "
        "identificadas. O uso da contingência exige registro no sistema de controle de "
        "mudanças, com justificativa documentada e ciência da patrocinadora.")

    secao(doc, 2, "5.8 Aquisições")
    paragrafo(doc, "Três aquisições demandam contratação externa: serviços de "
        "fotografia (diária, contrato simples com cessão de direitos de imagem); "
        "hospedagem e domínio (anual, incluídos na categoria de desenvolvimento); e "
        "licenças de fontes tipográficas e imagens. Para contratações acima de "
        "R$ 1.000,00 exige-se cotação mínima de três fornecedores.")
    paragrafo(doc, "O recebimento dos serviços é conferido pelo gerente contra o "
        "escopo contratado, com termo de aceite. Os contratos serão encerrados "
        "formalmente na fase de encerramento do projeto, sem pendências financeiras ou "
        "de direitos.")


def execucao(doc):
    secao(doc, 1, "6 EXECUÇÃO DO PROJETO")
    paragrafo(doc, "O grupo de processos de execução realiza o trabalho definido no "
        "plano de gerenciamento. Em projetos digitais de ciclo curto, a execução é "
        "marcada por entregas incrementais e aprovações frequentes da patrocinadora.")

    secao(doc, 2, "6.1 Orientação e gestão do trabalho")
    paragrafo(doc, "A direção do trabalho segue a linha de base aprovada: reuniões "
        "semanais de status às terças-feiras (30 minutos, online), demonstração da "
        "versão corrente a cada duas semanas e quadro kanban digital com as tarefas em "
        "andamento. As decisões e atas são registradas no diário do projeto.")
    paragrafo(doc, "Entregas parciais são publicadas em ambiente de homologação "
        "acessível à patrocinadora, permitindo feedback contínuo. Mudanças de campo de "
        "até R$ 1.000,00 podem ser aprovadas diretamente pelo gerente do projeto, com "
        "registro posterior no sistema de controle de mudanças.")

    secao(doc, 2, "6.2 Gestão do conhecimento")
    paragrafo(doc, "O conhecimento explícito foi organizado em guia de estilo "
        "documentado, catálogo de componentes da interface e vídeos curtos de três a "
        "cinco minutos para a recepção — como atender um lead que chega pelo CRM e "
        "responder pelo WhatsApp dentro do prazo combinado.")
    paragrafo(doc, "O conhecimento tácito é mobilizado em sessões de pareamento entre "
        "design e desenvolvimento e nas demonstrações quinzenais; as lições aprendidas "
        "são registradas mensalmente e alimentam o relatório final de encerramento.")

    secao(doc, 2, "6.3 Garantia da qualidade e engajamento")
    paragrafo(doc, "Cada entrega é precedida de revisão por pares de código e "
        "controlada pela matriz de testes em cinco dispositivos e três navegadores. A "
        "auditoria de privacidade verificou consentimento explícito no formulário, "
        "política de cookies e minimização de dados coletados, conforme a LGPD "
        "(BRASIL, 2018).")
    paragrafo(doc, "Os testes de usabilidade com cinco pacientes reais — agendar "
        "avaliação, encontrar convênios e localizar a clínica — alcançaram 87% de "
        "conclusão média, acima do critério de 80%. As comunicações foram executadas "
        "conforme o plano da Tabela 3, com destaque para as demonstrações quinzenais à "
        "proprietária.")

    secao(doc, 2, "6.4 Gestão das aquisições na execução")
    paragrafo(doc, "O ensaio fotográfico foi realizado em 08/04/2025 com roteiro "
        "aprovado, e as imagens foram incorporadas com cessão de direitos registrada. "
        "As licenças de fontes e a contratação de hospedagem foram documentadas no "
        "acervo do projeto, para encerramento formal na fase final.")


def monitoramento(doc):
    secao(doc, 1, "7 MONITORAMENTO E CONTROLE")
    paragrafo(doc, "O monitoramento compara o desempenho real com as linhas de base, "
        "avalia desvios e dispara ações preventivas e corretivas, além de processar "
        "mudanças pelo controle integrado.")

    secao(doc, 2, "7.1 Indicadores de desempenho e gerenciamento de valor agregado")
    paragrafo(doc, "O desempenho físico-financeiro é medido semanalmente pelo "
        "gerenciamento de valor agregado (EVM), que compara o valor planejado (PV), o "
        "valor agregado (EV) e o custo real (AC). A Figura 6 apresenta as curvas "
        "acumuladas e a Tabela 5 consolida os indicadores na data de dados de 11 de "
        "abril de 2025 (fim da semana 6).")
    figura(doc, svg_curva_s, "curva_s",
           "Curva S semanal — valor planejado (PV), valor agregado (EV) e custo real "
           "(AC), em R$ mil", "elaborada pelo autor (2025).", 14.5)
    tabela(doc,
           "Indicadores de valor agregado na data de dados (11/04/2025)",
           "elaborada pelo autor (2025).",
           ["Indicador", "Valor", "Leitura gerencial"],
           [
               ["BAC (orçamento no término)", "R$ 19.250,00",
                "Linha de base do custo aprovada no planejamento"],
               ["PV (valor planejado)", "R$ 13.500,00",
                "Custo planejado até a data de dados"],
               ["EV (valor agregado)", "R$ 12.400,00",
                "Valor do trabalho efetivamente realizado"],
               ["AC (custo real)", "R$ 14.200,00",
                "Custo incorrido até a data de dados"],
               ["SV (variação de prazo)", "-R$ 1.100,00",
                "Projeto atrasado em relação à linha de base"],
               ["CV (variação de custo)", "-R$ 1.800,00",
                "Gasto acima do previsto para o realizado"],
               ["SPI (índice de prazo)", "0,92",
                "Avanço cerca de 8% abaixo do planejado"],
               ["CPI (índice de custo)", "0,87",
                "Cada R$ 1,00 gasto gera R$ 0,87 de valor"],
               ["EAC (estimativa no término)", "R$ 22.051,00",
                "Custo final estimado sem ações corretivas"],
               ["VAC (variação no término)", "-R$ 2.801,00",
                "Estouro projetado do orçamento"],
               ["TCPI (desempenho necessário)", "1,36",
                "Manter o BAC exigiria eficiência de custo inviável"],
           ],
           larguras=[4.6, 3.2, 7.4], alinhar_centro={1})
    paragrafo(doc, "Os índices evidenciaram o materialização do risco R1: duas rodadas "
        "extras de revisão de design solicitadas pela patrocinadora geraram atraso e "
        "custo de retrabalho. As ações corretivas aprovadas foram: timebox de revisões "
        "(duas rodadas, decisões em 48 horas); fast-tracking, com SEO e testes em "
        "paralelo às integrações; redução do escopo opcional (seção Equipe simplificada); "
        "e uso de R$ 800,00 da contingência. A EAC revisada foi apresentada à "
        "patrocinadora, e as curvas da Figura 6 mostram a recuperação nas semanas "
        "seguintes, com custo final de R$ 19.100,00 e go-live no marco previsto.")

    secao(doc, 2, "7.2 Controle integrado de mudanças")
    paragrafo(doc, "Toda solicitação de mudança é registrada em formulário único e "
        "submetida à análise de impacto sobre escopo, prazo, custo e riscos. O gerente "
        "do projeto aprova alterações de até R$ 1.000,00; acima disso, decide a "
        "patrocinadora. Mudanças aprovadas atualizam a documentação e são comunicadas "
        "à equipe.")
    paragrafo(doc, "Na prática, o mecanismo aprovou o botão flutuante de WhatsApp "
        "(+R$ 450,00, sem impacto no prazo) e rejeitou a seção de depoimentos em "
        "vídeo — peso excessivo de mídia e custo acima da reserva. Não houve "
        "replanejamento das linhas de base; a EAC revisada foi documentada como "
        "projeção gerencial.")

    secao(doc, 2, "7.3 Controle de riscos e de interessados")
    paragrafo(doc, "O registro de riscos é revisado semanalmente, com gatilhos "
        "definidos: mais de duas rodadas de revisão de design; conteúdo pendente por "
        "mais de três dias; Lighthouse abaixo de 85 em qualquer entrega. O risco R1 "
        "foi acionado na semana 5 e mitigado pelo timebox; o R2 foi evitado com o "
        "cronograma de conteúdo negociado com a dentista.")
    paragrafo(doc, "O checklist de conformidade LGPD/CRO foi executado integralmente "
        "antes da janela de go-live, sem pendências. A matriz de interessados foi "
        "atualizada em abril: a recepção migrou para gerir de perto após coautorar os "
        "campos do formulário, elevando a qualidade dos dados captados no CRM.")


def encerramento(doc):
    secao(doc, 1, "8 ENCERRAMENTO DO PROJETO")
    paragrafo(doc, "O encerramento formaliza a conclusão de todas as atividades, "
        "transfere o produto e consolida o conhecimento gerado.")

    secao(doc, 2, "8.1 Aceite e encerramento administrativo")
    paragrafo(doc, "A verificação do escopo comparou cada entrega da linha de base "
        "com o produto realizado por meio de checklist. O relatório Lighthouse final "
        "registrou 92 em performance, 94 em acessibilidade, 100 em boas práticas e 96 "
        "em SEO — todos acima do critério de 90. Os dez envios de teste do formulário "
        "chegaram ao CRM em menos de um minuto.")
    paragrafo(doc, "O termo de aceite foi assinado pela proprietária em 16/05/2025. "
        "Foram transferidos à clínica os acessos (hospedagem, domínio, CRM e "
        "Analytics), o repositório de código, a documentação e os vídeos de "
        "treinamento; a titularidade da hospedagem e do domínio passou ao CNPJ da "
        "clínica, e o contrato do fotógrafo foi encerrado sem pendências.")

    secao(doc, 2, "8.2 Lições aprendidas")
    paragrafo(doc, "O relatório final consolidou as principais lições aprendidas:")
    for item in [
        "o timebox de duas rodadas de revisão eliminou o ciclo de retrabalho de design;",
        "envolver a recepção cedo melhorou os campos do formulário e a qualidade dos leads;",
        "fotos autorais superaram banco de imagens nos testes de usabilidade (preferência de 85%);",
        "cerca de 46% da contingência foi consumida pelo retrabalho de design (R$ 800,00);",
        "comunicação assíncrona com a agência (vídeos curtos) reduziu reuniões sem perda de alinhamento;",
        "o TCPI elevado sinalizou cedo a necessidade de renegociar escopo opcional com a patrocinadora.",
    ]:
        marcador(doc, item)
    paragrafo(doc, "O fechamento financeiro registrou custo final de R$ 19.100,00 "
        "para valor agregado de R$ 19.250,00, resultando em CPI final de 1,01 e SPI "
        "final de 1,00 — projeto entregue no marco de go-live, dentro do orçamento e "
        "com aceite formal da patrocinadora.")


def consideracoes(doc):
    secao(doc, 1, "9 CONSIDERAÇÕES FINAIS")
    paragrafo(doc, "O projeto demonstra que as práticas do Guia PMBOK se aplicam com "
        "eficácia a iniciativas digitais de ciclo curto, desde que calibradas em "
        "cerimônias leves: reuniões de 30 minutos, demonstrações quinzenais e "
        "indicadores semanais. A disciplina de linhas de base e de controle integrado "
        "de mudanças foi decisiva para conter o efeito das revisões sucessivas de "
        "design — risco típico de projetos com patrocinador único e altamente "
        "envolvido.")
    paragrafo(doc, "Como limitações, destacam-se a sensibilidade das estimativas em "
        "projetos curtos e a dependência de decisões rápidas da patrocinadora. Como "
        "próximos passos, propõem-se testes A/B de títulos e propostas de valor, a "
        "criação de páginas por especialidade (implante e ortodontia) reaproveitando o "
        "guia de estilo, e a medição do KPI de aumento de 30% nos agendamentos "
        "digitais ao final dos 90 dias após o go-live.")


def referencias(doc):
    par = secao(doc, 1, "REFERÊNCIAS")
    par.alignment = CEN
    for ref in [
        "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR 14724: informação e "
        "documentação: trabalhos acadêmicos: apresentação. Rio de Janeiro: ABNT, 2011.",
        "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR 6023: informação e "
        "documentação: referências: elaboração. Rio de Janeiro: ABNT, 2018.",
        "BRASIL. Lei nº 13.709, de 14 de agosto de 2018. Lei Geral de Proteção de "
        "Dados Pessoais (LGPD). Brasília, DF: Presidência da República, 2018.",
        "CONSELHO FEDERAL DE ODONTOLOGIA (CFO). Resolução CFO-159, de 8 de dezembro "
        "de 2016. Aprova as normas procedimentais para a publicidade odontológica. "
        "Brasília, DF: CFO, 2016.",
        "KERZNER, H. Gestão de projetos: as melhores práticas. 3. ed. Porto Alegre: "
        "Bookman, 2017.",
        "KRUG, S. Não me faça pensar: uma abordagem de bom senso à navegação na web e "
        "na web móvel. 3. ed. Rio de Janeiro: Alta Books, 2015.",
        "PROJECT MANAGEMENT INSTITUTE. Um guia do conhecimento em gerenciamento de "
        "projetos: guia PMBOK. 6. ed. Newtown Square: PMI, 2017.",
        "PROJECT MANAGEMENT INSTITUTE. Um guia do conhecimento em gerenciamento de "
        "projetos: guia PMBOK. 7. ed. Newtown Square: PMI, 2021.",
    ]:
        doc.add_paragraph(ref, style="Referência ABNT")


def textual(doc):
    introducao(doc)
    objetivos(doc)
    justificativa(doc)
    iniciacao(doc)
    planejamento(doc)
    execucao(doc)
    monitoramento(doc)
    encerramento(doc)
    consideracoes(doc)
    referencias(doc)


# ===============================================================================
# Execução principal
# ===============================================================================
def main():
    DIR_FIGURAS.mkdir(parents=True, exist_ok=True)
    doc = Document()
    configurar_pagina(doc)
    criar_estilos(doc)

    cp = doc.core_properties
    cp.title = "Landing Page da Clínica Odontológica Sorriso Real — Guia PMBOK"
    cp.author = "Maria Alves da Silva"
    cp.subject = "Gerenciamento de projetos (PMBOK)"

    capa(doc)                                        # seção 1: capa (sem número)
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    folha_de_rosto(doc)                              # seção 2: pré-texto (conta da pág. 2)
    pre_textual(doc)
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    textual(doc)                                     # seção 3: texto (nº no canto sup. direito)

    numeracao(doc)
    ativar_atualizacao_campos(doc)

    doc.save(NOME_ARQUIVO)
    caminho = Path(NOME_ARQUIVO).resolve()
    print("=" * 70)
    print(f"Documento gerado : {caminho}")
    print(f"Figuras em       : {DIR_FIGURAS.resolve()} ({_CONT['Figura']} SVG/PNG)")
    print(f"Conversor SVG    : {CONVERSOR_USADO or 'nenhum (instale cairosvg ou svglib+reportlab)'}")
    print("Dica: no Word o sumário/listas/páginas se atualizam ao abrir;")
    print("      no LibreOffice use Ferramentas > Atualizar > Atualizar tudo.")
    print("=" * 70)


if __name__ == "__main__":
    main()