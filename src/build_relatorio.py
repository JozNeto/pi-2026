# -*- coding: utf-8 -*-
"""Gera o Relatório Parcial (V1) em .docx no padrão ABNT (estrutura de TCC).

Todos os números do texto vêm dos JSONs em reports/ (gerados por src/eda.py,
src/forecast.py e src/anomaly.py), evitando erros de transcrição.

Uso: python -m src.build_relatorio
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "reports" / "figures"
OUT = ROOT / "reports" / "final" / "Relatorio_Parcial_V1_SolarSync.docx"

E = json.loads((ROOT / "reports" / "eda_resultados.json").read_text(encoding="utf-8"))
F = json.loads((ROOT / "reports" / "forecast_resultados.json").read_text(encoding="utf-8"))
A = json.loads((ROOT / "reports" / "anomalia_resultados.json").read_text(encoding="utf-8"))

FONTE = "Times New Roman"


# ------------------------------------------------------------------ formatação numérica pt-BR
def n(x, nd=0):
    s = f"{x:,.{nd}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x, nd=1):
    return n(x * 100, nd) + "%"


def dia_br(iso):
    y, m, d = iso.split("-")
    return f"{d}/{m}"


# ------------------------------------------------------------------ documento e estilos
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin, sec.left_margin, sec.bottom_margin, sec.right_margin = Cm(3), Cm(3), Cm(2), Cm(2)
sec.header_distance = Cm(1.25)

ST = doc.styles


def fonte(style, size=12, bold=False, italic=False):
    style.font.name = FONTE
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONTE)


def par_fmt(style, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first=0.0, left=0.0, before=0, after=0, line=1.5, keep_next=False, page_break=False):
    pf = style.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first)
    pf.left_indent = Cm(left)
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    pf.line_spacing = line
    pf.keep_with_next = keep_next
    pf.page_break_before = page_break
    pf.widow_control = True


normal = ST["Normal"]
fonte(normal, 12)
par_fmt(normal, first=1.25)

for nome, size, bold, ital, before, after, brk in (
    ("Heading 1", 12, True, False, 0, 18, True),
    ("Heading 2", 12, False, False, 18, 18, False),
    ("Heading 3", 12, True, False, 18, 18, False),
):
    s = ST[nome]
    fonte(s, size, bold, ital)
    par_fmt(s, align=WD_ALIGN_PARAGRAPH.LEFT, first=0, before=before, after=after, keep_next=True, page_break=brk)


def novo_estilo(nome, base="Normal", size=12, bold=False, italic=False, **kw):
    s = ST.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
    s.base_style = ST[base]
    fonte(s, size, bold, italic)
    par_fmt(s, **kw)
    return s


novo_estilo("TituloPre", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, first=0, before=0, after=18, keep_next=True, page_break=True)
novo_estilo("Resumo", size=12, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
novo_estilo("Simples", size=12, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.LEFT)
novo_estilo("Centro", size=12, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
novo_estilo("Alinea", first=-0.63, left=1.9, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
novo_estilo("LegendaABNT", size=10, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER, before=12, after=4, keep_next=True)
novo_estilo("FonteABNT", size=10, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=12)
novo_estilo("FiguraABNT", size=12, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER, keep_next=True)
novo_estilo("CelulaABNT", size=10, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.LEFT)
novo_estilo("Referencia", size=12, first=0, line=1.0, align=WD_ALIGN_PARAGRAPH.LEFT, after=12)
novo_estilo("SiglaLinha", size=12, first=0, line=1.5, align=WD_ALIGN_PARAGRAPH.LEFT)
novo_estilo("Natureza", size=12, first=0, left=8.0, line=1.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

for nome, ind in (("toc 1", 0), ("toc 2", 0.6), ("toc 3", 1.2), ("table of figures", 0)):
    s = ST.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
    s.base_style = ST["Normal"]
    fonte(s, 12)
    par_fmt(s, align=WD_ALIGN_PARAGRAPH.LEFT, first=0, left=ind, line=1.0, after=6)
    s.paragraph_format.tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)


# ------------------------------------------------------------------ utilitários de conteúdo
TAG = re.compile(r"(<b>.*?</b>|<i>.*?</i>)")


def runs(p, texto, size=None, hl=False):
    for seg in TAG.split(texto):
        if not seg:
            continue
        b, i = seg.startswith("<b>"), seg.startswith("<i>")
        if b or i:
            seg = seg[3:-4]
        r = p.add_run(seg)
        r.bold, r.italic = b or None, i or None
        if size:
            r.font.size = Pt(size)
        if hl:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW


def P(texto, style="Normal", align=None, hl=False):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    runs(p, texto, hl=hl)
    return p


def campo(p, instr, cache="", dirty=False, bold=False, size=None):
    def mk(tipo):
        r = p.add_run()
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), tipo)
        if dirty and tipo == "begin":
            fc.set(qn("w:dirty"), "true")
        r._r.append(fc)
        return r
    mk("begin")
    r = p.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r._r.append(it)
    mk("separate")
    rc = p.add_run(cache)
    rc.bold = bold or None
    if size:
        rc.font.size = Pt(size)
    mk("end")


def H(nivel, texto):
    p = doc.add_heading("", level=nivel)
    runs(p, texto)
    return p


def A_(itens):
    """Alíneas ABNT (a), b), ...)."""
    for k, it in enumerate(itens):
        fim = "." if k == len(itens) - 1 else ";"
        P(f"{chr(97 + k)}) {it}{fim}", style="Alinea")


CONT = {"Figura": 0, "Tabela": 0}


def legenda(tipo, texto):
    CONT[tipo] += 1
    p = doc.add_paragraph(style="LegendaABNT")
    p.add_run(f"{tipo} ")
    campo(p, f"SEQ {tipo} \\* ARABIC", str(CONT[tipo]))
    p.add_run(f" – {texto}")
    return CONT[tipo]


def fonte_txt(texto="Elaborado pelos autores (2026)."):
    P(f"Fonte: {texto}", style="FonteABNT")


def figura(arq, titulo, largura=15.5, fonte="Elaborado pelos autores (2026)."):
    legenda("Figura", titulo)
    p = doc.add_paragraph(style="FiguraABNT")
    p.add_run().add_picture(str(FIG / arq), width=Cm(largura))
    fonte_txt(fonte)


def borda(cell, **lados):
    tcPr = cell._tc.get_or_add_tcPr()
    b = tcPr.find(qn("w:tcBorders"))
    if b is None:
        b = OxmlElement("w:tcBorders")
        tcPr.append(b)
    for lado, sz in lados.items():
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        b.append(e)


def tabela(titulo, cab, linhas, larg, alinh=None, fonte="Elaborado pelos autores (2026).", nota=None):
    """Tabela no padrão IBGE/ABNT: sem linhas verticais; traços superior, sob o cabeçalho e inferior."""
    legenda("Tabela", titulo)
    t = doc.add_table(rows=1 + len(linhas), cols=len(cab))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    alinh = alinh or ["l"] + ["c"] * (len(cab) - 1)
    mapa = {"l": WD_ALIGN_PARAGRAPH.LEFT, "c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT}
    for i, linha in enumerate([cab] + linhas):
        row = t.rows[i]
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        trPr.append(cs)
        if i == 0:
            th = OxmlElement("w:tblHeader")
            trPr.append(th)
        for j, v in enumerate(linha):
            c = row.cells[j]
            c.width = Cm(larg[j])
            p = c.paragraphs[0]
            p.style = ST["CelulaABNT"]
            p.alignment = mapa[alinh[j]] if i > 0 else WD_ALIGN_PARAGRAPH.CENTER
            runs(p, ("<b>" + v + "</b>") if i == 0 else v, size=10)
            if i == 0:
                borda(c, top=8, bottom=4)
            if i == len(linhas):
                borda(c, bottom=8)
    if nota:
        P(nota, style="FonteABNT", align=WD_ALIGN_PARAGRAPH.LEFT)
    fonte_txt(fonte)


def quebra():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ================================================================== PRÉ-TEXTUAIS
AUTORES = ["Adriano Josias da Silva", "Cristiane T. Biazotto", "Jose Araujo Neto", "Mayckon Oliveira Campos", "Renan Silva Pinheiro"]
TITULO = "SolarSync: previsão de geração e detecção de anomalias em sistemas fotovoltaicos residenciais por meio de aprendizagem de máquina"

# ---- Capa
P("UNIVERSIDADE VIRTUAL DO ESTADO DE SÃO PAULO", "Centro", hl=False).runs[0].bold = True
P("[Curso] — [Polo(s)]", "Centro", hl=True)
for _ in range(6):
    P("", "Centro")
for a in AUTORES:
    P(a.upper(), "Centro")
for _ in range(7):
    P("", "Centro")
p = P(f"<b>{TITULO.upper()}</b>", "Centro")
for _ in range(2):
    P("", "Centro")
P("Relatório Parcial (V1) — Projeto Integrador IV", "Centro")
for _ in range(9):
    P("", "Centro")
P("[Cidade]", "Centro", hl=True)
P("2026", "Centro")

# ---- Folha de rosto
quebra()
for a in AUTORES:
    P(a.upper(), "Centro")
for _ in range(9):
    P("", "Centro")
P(f"<b>{TITULO.upper()}</b>", "Centro")
for _ in range(3):
    P("", "Centro")
P("Relatório Parcial apresentado ao Projeto Integrador IV da Universidade Virtual do Estado de São Paulo (Univesp), "
  "como parte das atividades avaliativas da disciplina, correspondente à Quinzena 4 do Plano de Ação.", "Natureza")
P("", "Natureza")
P("Orientadora: Profa. Hellen Patricia Lemos Cordovil", "Natureza")
for _ in range(7):
    P("", "Centro")
P("[Cidade]", "Centro", hl=True)
P("2026", "Centro")

# ---- Resumo
q = E["qualidade"]
h1 = F["horizonte"]["1 h"]
gb1, pe1 = h1["Gradient Boosting"], h1["Persistência"]
RESUMO = (
    f"Este relatório parcial apresenta a primeira versão (V1) do trabalho desenvolvido no Projeto Integrador IV a partir do SolarSync, "
    f"sistema de monitoramento de geração solar fotovoltaica que adquire dados de um inversor híbrido por meio de um ESP32 (RS485/Modbus) "
    f"e os armazena em um banco MongoDB Atlas. Com o objetivo de acrescentar ao sistema uma camada de aprendizagem de máquina, foram "
    f"analisados {n(E['n_registros'])} registros coletados entre 6 e 20 de setembro de 2026 (uma leitura a cada {n(q['intervalo_mediana_s'], 1)} s, "
    f"em mediana) e construídos modelos de previsão de curto prazo da potência fotovoltaica, em resolução de 15 minutos, além de uma linha de base "
    f"de detecção de anomalias. A metodologia compreendeu extração dos dados em modo somente leitura, análise exploratória, reamostragem, "
    f"engenharia de atributos e validação <i>walk-forward</i> diária em {F['config']['n_dias_teste']} dias de teste. Foram comparados os modelos de referência "
    f"(persistência, sazonal ingênuo de 24 horas e climatologia), a regressão Ridge e o <i>Gradient Boosting</i>. Para o horizonte de uma hora, o "
    f"<i>Gradient Boosting</i> obteve erro absoluto médio de {n(gb1['mae'])} W contra {n(pe1['mae'])} W da persistência (redução de "
    f"{pct(gb1['skill_vs_persistencia'], 0)}), superando-a em {F['gb_vence_persistencia_1h_dias']} dos {F['n_dias_avaliados_1h']} dias; para 15 minutos "
    f"a persistência foi o melhor modelo e, para três horas e para o dia seguinte, a climatologia obteve o menor erro. Constatou-se ainda que a potência "
    f"fotovoltaica medida é limitada pelo inversor quando a bateria atinge carga plena, o que restringe a interpretação da série como geração potencial. "
    f"A detecção de anomalias com <i>Isolation Forest</i> sinalizou principalmente picos de carga e teve sobreposição desprezível com a regra atual "
    f"de tensão de saída. Os resultados são preliminares em razão do histórico curto (15 dias) e orientam as etapas seguintes do projeto."
)
doc.add_paragraph("RESUMO", style="TituloPre")
P(RESUMO, "Resumo")
P("", "Resumo")
P("<b>Palavras-chave</b>: Energia solar fotovoltaica. Previsão de séries temporais. Aprendizagem de máquina. Internet das Coisas. Detecção de anomalias.", "Resumo")

# ---- Abstract
ABSTRACT = (
    f"This partial report presents the first version (V1) of the work carried out in Projeto Integrador IV based on SolarSync, a photovoltaic "
    f"monitoring system that acquires data from a hybrid inverter through an ESP32 (RS485/Modbus) and stores them in a MongoDB Atlas database. "
    f"To add a machine learning layer to the system, {E['n_registros']:,} records collected between September 6 and 20, 2026 were analysed, and short-term "
    f"photovoltaic power forecasting models (15-minute resolution) and an anomaly detection baseline were built. The methodology comprised read-only data "
    f"extraction, exploratory analysis, resampling, feature engineering and daily walk-forward validation over {F['config']['n_dias_teste']} test days. Reference models "
    f"(persistence, 24-hour seasonal naive and climatology), Ridge regression and Gradient Boosting were compared. For the one-hour horizon, Gradient Boosting "
    f"achieved a mean absolute error of {n(gb1['mae'])} W against {n(pe1['mae'])} W for persistence ({pct(gb1['skill_vs_persistencia'], 0)} lower), outperforming it on "
    f"{F['gb_vence_persistencia_1h_dias']} of {F['n_dias_avaliados_1h']} days; at 15 minutes persistence was the best model, and at three hours and day-ahead climatology "
    f"had the lowest error. The analysis also showed that measured photovoltaic power is curtailed by the inverter when the battery is fully charged, which limits "
    f"its interpretation as potential generation. Anomaly detection with Isolation Forest mostly flagged load peaks and barely overlapped with the current "
    f"output-voltage rule. Results are preliminary due to the short history (15 days) and guide the next stages of the project."
)
doc.add_paragraph("ABSTRACT", style="TituloPre")
P(ABSTRACT, "Resumo")
P("", "Resumo")
P("<b>Keywords</b>: Photovoltaic solar energy. Time series forecasting. Machine learning. Internet of Things. Anomaly detection.", "Resumo")

# ---- Listas (campos atualizados pelo Word)
doc.add_paragraph("LISTA DE ILUSTRAÇÕES", style="TituloPre")
campo(doc.add_paragraph(style="Simples"), 'TOC \\h \\z \\c "Figura"', "(atualizar campo)", dirty=True)
doc.add_paragraph("LISTA DE TABELAS", style="TituloPre")
campo(doc.add_paragraph(style="Simples"), 'TOC \\h \\z \\c "Tabela"', "(atualizar campo)", dirty=True)

doc.add_paragraph("LISTA DE ABREVIATURAS E SIGLAS", style="TituloPre")
SIGLAS = [
    ("ABNT", "Associação Brasileira de Normas Técnicas"),
    ("API", "<i>Application Programming Interface</i> (interface de programação de aplicações)"),
    ("CA", "Corrente alternada"),
    ("ESP32", "Microcontrolador com Wi-Fi utilizado na aquisição de dados"),
    ("FV", "Fotovoltaica(o)"),
    ("IoT", "<i>Internet of Things</i> (Internet das Coisas)"),
    ("MAE", "<i>Mean Absolute Error</i> (erro absoluto médio)"),
    ("MPPT", "<i>Maximum Power Point Tracking</i> (rastreamento do ponto de máxima potência)"),
    ("nMAE", "MAE normalizado pela média diurna da potência observada"),
    ("RMSE", "<i>Root Mean Squared Error</i> (raiz do erro quadrático médio)"),
    ("RS485", "Padrão de comunicação serial diferencial"),
    ("SBU / SUB", "Modos de prioridade de fonte de saída reportados pelo inversor"),
    ("SOC", "<i>State of Charge</i> (estado de carga da bateria)"),
    ("Univesp", "Universidade Virtual do Estado de São Paulo"),
]
for sg, sig in SIGLAS:
    p = doc.add_paragraph(style="SiglaLinha")
    p.paragraph_format.tab_stops.add_tab_stop(Cm(3.5))
    p.paragraph_format.left_indent = Cm(3.5)
    p.paragraph_format.first_line_indent = Cm(-3.5)
    runs(p, f"{sg}\t{sig}")

# ---- Sumário
doc.add_paragraph("SUMÁRIO", style="TituloPre")
campo(doc.add_paragraph(style="Simples"), 'TOC \\o "1-3" \\h \\z \\u', "(atualizar campo)", dirty=True)

# ================================================================== TEXTUAL (nova seção; numeração a partir daqui)
sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
sec2.header.is_linked_to_previous = False
hp = sec2.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
campo(hp, "PAGE", "1", size=10)
for r in hp.runs:
    r.font.size = Pt(10)
    r.font.name = FONTE
# seção 1: cabeçalho vazio (páginas pré-textuais contadas, mas sem número visível)
doc.sections[0].header.is_linked_to_previous = False

# ------------------------------------------------------------------ 1 INTRODUÇÃO
H(1, "1 INTRODUÇÃO")
H(2, "1.1 CONTEXTUALIZAÇÃO E MOTIVAÇÃO")
P("A geração distribuída de energia elétrica a partir de fontes renováveis, em especial a fotovoltaica, expandiu-se no Brasil com o marco legal da "
  "microgeração e minigeração distribuída (BRASIL, 2022). Nesse cenário, sistemas residenciais compostos por painéis fotovoltaicos, inversores híbridos "
  "e baterias tornaram-se cada vez mais comuns, e o acompanhamento do seu funcionamento passou a depender, em geral, de aplicativos fornecidos pelos "
  "fabricantes dos equipamentos.")
P("O SolarSync surgiu de uma necessidade prática de monitoramento de um sistema de geração fotovoltaica com inversor híbrido da fabricante Anenji, "
  "instalado na residência de um dos integrantes do grupo. O aplicativo genérico disponibilizado para esse equipamento atualiza os dados a cada "
  "aproximadamente 15 minutos, apresenta um conjunto restrito de grandezas, oferece poucos recursos de análise histórica e não permite personalizar "
  "indicadores ou alertas (GRUPO SOLARSYNC, 2026). Como resposta, foi desenvolvida uma arquitetura própria de aquisição, transmissão, armazenamento "
  "e visualização dos dados do inversor em tempo praticamente real.")
P("A infraestrutura existente atende à maior parte dos requisitos do tema norteador do Projeto Integrador IV — desenvolver análise de dados em escala "
  "utilizando dados capturados por Internet das Coisas (IoT) e aprendizagem de máquina, além de preparar uma interface para visualização dos resultados —, "
  "mas ainda não incorpora aprendizagem de máquina nem análise preditiva. Os alertas atuais, por exemplo, baseiam-se em regras e limites fixos, como a "
  "tensão de saída em corrente alternada (CA) fora da faixa de 212 V a 230 V (GRUPO SOLARSYNC, 2026).")
H(2, "1.2 PROBLEMA DE PESQUISA")
P("Como aplicar aprendizagem de máquina aos dados de telemetria já armazenados pelo SolarSync para (i) prever a geração de energia fotovoltaica em "
  "curto prazo e (ii) detectar automaticamente comportamentos anômalos do sistema, sem depender exclusivamente de limiares fixos definidos manualmente?")
H(2, "1.3 OBJETIVOS")
H(3, "1.3.1 Objetivo geral")
P("Desenvolver, a partir da infraestrutura já existente do SolarSync (ESP32, API Node.js e banco MongoDB Atlas), uma camada de análise preditiva capaz de "
  "prever a geração de energia fotovoltaica em curto prazo e detectar anomalias de funcionamento por meio de aprendizagem de máquina, apresentando os "
  "resultados de forma integrada à interface de visualização utilizada pelo sistema.")
H(3, "1.3.2 Objetivos específicos")
A_([
    "caracterizar a base de telemetria disponível quanto à qualidade, à cobertura temporal e aos padrões de operação;",
    "construir e validar modelos de previsão de curto prazo da potência fotovoltaica, comparando-os a modelos de referência;",
    "elaborar uma linha de base não supervisionada de detecção de anomalias e compará-la com a regra atualmente utilizada pelo sistema;",
    "identificar limitações dos dados e do sistema que condicionam a interpretação dos resultados;",
    "definir as etapas seguintes para a solução final, incluindo a apresentação dos resultados no <i>dashboard</i> existente",
])
H(2, "1.4 JUSTIFICATIVA")
P("O trabalho utiliza dados reais, coletados por um sistema IoT em operação, e integra conteúdos de comunicação industrial, sistemas embarcados, bancos "
  "de dados, análise de dados e aprendizagem de máquina. Do ponto de vista prático, prever a geração em horizontes de horas permite planejar o uso "
  "da bateria e do consumo, e detectar comportamentos fora do padrão sem depender de limiares fixos aumenta a capacidade de o sistema alertar o "
  "usuário. Do ponto de vista acadêmico, a base ainda curta oferece a oportunidade de discutir, com rigor, o que é e o que não é possível concluir "
  "com poucos dados — discussão central deste relatório.")
H(2, "1.5 ORGANIZAÇÃO DO TRABALHO")
P("Este relatório está organizado em cinco seções. A seção 2 apresenta o referencial teórico; a seção 3 descreve a metodologia; a seção 4 traz a "
  "análise exploratória, os resultados dos modelos e a discussão; e a seção 5 apresenta as considerações parciais e as próximas etapas. Ao final, "
  "constam as referências e dois apêndices.")

# ------------------------------------------------------------------ 2 REFERENCIAL
H(1, "2 REFERENCIAL TEÓRICO")
H(2, "2.1 GERAÇÃO FOTOVOLTAICA DISTRIBUÍDA E INVERSORES HÍBRIDOS")
P("Um sistema fotovoltaico híbrido combina painéis, um inversor capaz de operar em conjunto com a rede elétrica e com armazenamento, e um banco de baterias. "
  "A potência entregue pelos painéis depende da irradiância e da temperatura; em inversores híbridos, depende também da estratégia de controle. "
  "Em geral, quando a bateria está plenamente carregada e o consumo é inferior à capacidade de geração, o controlador de rastreamento do ponto de "
  "máxima potência (MPPT) reduz a potência extraída dos painéis. Assim, a potência FV <i>medida</i> nem sempre coincide com a potência FV "
  "<i>disponível</i>, característica que é relevante para a interpretação das séries analisadas (seção 4.1.4).")
H(2, "2.2 AQUISIÇÃO DE DADOS: IoT, RS485 E MODBUS")
P("A Internet das Coisas (IoT) refere-se à conexão de dispositivos físicos a redes de dados para coleta e transmissão de informações. No SolarSync, um "
  "microcontrolador ESP32 lê periodicamente registradores do inversor por meio de um conversor MAX485, que adapta a interface serial do ESP32 ao "
  "barramento RS485, padrão de transmissão diferencial adequado a equipamentos industriais (GRUPO SOLARSYNC, 2026). O protocolo Modbus organiza a "
  "comunicação em modelo mestre-escravo, por leitura e escrita de registradores endereçáveis (MODBUS ORGANIZATION, 2012), e a ferramenta QModMaster foi "
  "empregada para identificar os endereços dos registradores relevantes.")
H(2, "2.3 PREVISÃO DE GERAÇÃO FOTOVOLTAICA COMO SÉRIE TEMPORAL")
P("A potência FV é uma série temporal com forte componente diurna determinística (ciclo solar) e uma componente estocástica, associada sobretudo à "
  "nebulosidade. Na previsão de curtíssimo e curto prazo, é prática consolidada comparar modelos mais elaborados a modelos de referência simples, como "
  "a persistência e a climatologia (INMAN; PEDRO; COIMBRA, 2013). A literatura também associa horizontes mais longos ao uso de informações "
  "meteorológicas externas (INMAN; PEDRO; COIMBRA, 2013), fato que condiciona as conclusões deste trabalho, cuja base não contém dados de clima. "
  "Abordagens de aprendizagem de máquina para séries temporais costumam recorrer a atributos construídos a partir de defasagens (<i>lags</i>) e "
  "estatísticas móveis da própria série (HYNDMAN; ATHANASOPOULOS, 2021).")
H(2, "2.4 ALGORITMOS UTILIZADOS")
H(3, "2.4.1 Modelos de referência")
P("<b>Persistência</b>: a previsão para o instante <i>t</i>+<i>h</i> é o valor observado em <i>t</i>. <b>Sazonal ingênuo (24 h)</b>: a previsão é o valor "
  "observado no mesmo horário do dia anterior. <b>Climatologia</b>: a previsão é a média histórica da potência naquele horário do dia, calculada apenas "
  "com dados de treino. Esses modelos não têm parâmetros a ajustar e servem de piso de desempenho para os demais.")
H(3, "2.4.2 Regressão Ridge")
P("A regressão Ridge é uma regressão linear com penalização da norma L2 dos coeficientes, que reduz a variância das estimativas quando há atributos "
  "correlacionados (HOERL; KENNARD, 1970). Os atributos foram padronizados antes do ajuste e o parâmetro de regularização foi fixado em 1,0.")
H(3, "2.4.3 <i>Gradient Boosting</i>")
P("O <i>Gradient Boosting</i> constrói, de forma sequencial, um <i>ensemble</i> aditivo de árvores de decisão rasas, em que cada nova árvore é ajustada "
  "ao gradiente da função de perda do conjunto anterior (FRIEDMAN, 2001; HASTIE; TIBSHIRANI; FRIEDMAN, 2009). Neste trabalho foi utilizada a implementação "
  "baseada em histogramas do <i>scikit-learn</i> (PEDREGOSA et al., 2011), inspirada no LightGBM (KE et al., 2017), que trata valores ausentes de modo "
  "nativo e é eficiente em bases de tamanho reduzido.")
H(2, "2.5 DETECÇÃO DE ANOMALIAS E <i>ISOLATION FOREST</i>")
P("Anomalias são padrões nos dados que não se conformam a um comportamento normal esperado (CHANDOLA; BANERJEE; KUMAR, 2009). Na ausência de rótulos de "
  "falha, são usados métodos não supervisionados. O <i>Isolation Forest</i> isola observações por meio de partições aleatórias sucessivas do espaço de "
  "atributos; como observações atípicas são separadas com menos partições, o comprimento médio do caminho nas árvores fornece um escore de anomalia "
  "(LIU; TING; ZHOU, 2008). O parâmetro de contaminação define a fração esperada de pontos sinalizados.")
H(2, "2.6 VALIDAÇÃO E MÉTRICAS")
P("Em séries temporais, a validação deve preservar a ordem cronológica: o modelo é treinado apenas com o passado e avaliado no futuro, evitando "
  "vazamento de informação (BERGMEIR; BENÍTEZ, 2012). Adota-se aqui a validação <i>walk-forward</i>, em que o conjunto de treino cresce a cada dia de "
  "teste. As métricas utilizadas são o erro absoluto médio (MAE) e a raiz do erro quadrático médio (RMSE), ambos em watts, e o MAE normalizado pela "
  "média da potência observada (nMAE) (HYNDMAN; KOEHLER, 2006). O ganho relativo a um modelo de referência é definido como 1 − MAE<sub>modelo</sub>/MAE<sub>referência</sub>, "
  "positivo quando o modelo supera a referência.".replace("<sub>", "").replace("</sub>", ""))

# ------------------------------------------------------------------ 3 METODOLOGIA
H(1, "3 METODOLOGIA")
H(2, "3.1 CARACTERIZAÇÃO DA PESQUISA")
P("Trata-se de uma pesquisa aplicada, de abordagem quantitativa e caráter exploratório e experimental, realizada sobre dados reais de um sistema fotovoltaico "
  "residencial. O estudo é um estudo de caso único; portanto, seus resultados descrevem o sistema analisado e o período coberto, sem pretensão de "
  "generalização estatística para outras instalações ou estações do ano.")
H(2, "3.2 O SISTEMA SOLARSYNC")
P("A Figura 1 sintetiza a arquitetura. O inversor híbrido é lido por um ESP32 (via MAX485, RS485/Modbus); os dados são enviados por Wi-Fi a uma API "
  "desenvolvida em Node.js, que os armazena no MongoDB Atlas; um <i>dashboard</i> em React e Vite consulta o banco e exibe as informações. A camada de "
  "aprendizagem de máquina, objeto deste trabalho, consome os dados armazenados em modo somente leitura e não altera nenhum componente da "
  "infraestrutura existente.")
figura("fig_arquitetura.png", "Arquitetura do SolarSync e posição da camada de aprendizagem de máquina", 15.0,
       "Elaborado pelos autores (2026), com base em Grupo SolarSync (2026).")
H(2, "3.3 EXTRAÇÃO DOS DADOS")
P(f"Os dados foram extraídos da coleção de telemetria do banco <i>anenji_monitor</i>, no MongoDB Atlas, em 20 de setembro de 2026, com um usuário de "
  f"permissão exclusivamente de leitura. Foram executadas somente operações de consulta (<i>find</i> com projeção dos campos de interesse); nenhuma "
  f"operação de escrita foi realizada, e a credencial de acesso foi mantida em variável de ambiente, sem registro em arquivos. A extração resultou em "
  f"{n(E['n_registros'])} registros, gravados em arquivo local acompanhado de metadados de proveniência (data da extração, contagem e <i>hash</i> SHA-256). "
  f"Cada registro contém instante da leitura, grandezas do arranjo fotovoltaico, da rede, da saída CA, da bateria e da carga, além do modo e do "
  f"estado do inversor (Apêndice A).")
H(2, "3.4 PRÉ-PROCESSAMENTO")
P("O MongoDB armazena datas em tempo universal coordenado (UTC). Os instantes foram convertidos para o horário local (UTC−3, sem horário de verão), "
  "o que foi corroborado pelo perfil diário da geração, que se estende de cerca das 6 h às 18 h locais (seção 4.1.3). Foram verificados duplicidade "
  "de instantes, valores ausentes e intervalos entre leituras. Para a modelagem, as leituras foram agregadas por média em janelas de 15 minutos; "
  "janelas com menos de 10 leituras foram descartadas (tratadas como ausentes) e lacunas de até duas janelas consecutivas foram preenchidas por "
  "interpolação linear. A energia diária foi estimada pela integração da potência ao longo do tempo, limitando o intervalo entre leituras consecutivas a 60 s "
  "(lacunas maiores não contribuem, de modo que a estimativa é conservadora).")
H(2, "3.5 ANÁLISE EXPLORATÓRIA")
P("A análise exploratória compreendeu: (i) avaliação da qualidade da base; (ii) estatísticas descritivas; (iii) perfil horário da potência FV; "
  "(iv) energia diária gerada e consumida; (v) correlações entre grandezas; (vi) relação entre o estado de carga da bateria e a potência FV; e "
  "(vii) identificação de episódios de ausência de rede elétrica.")
H(2, "3.6 MODELAGEM DA PREVISÃO")
H(3, "3.6.1 Definição do problema")
P("O problema consiste em prever a potência FV média em janelas de 15 minutos, <i>h</i> passos à frente, a partir de informações disponíveis no instante de "
  "origem <i>t</i>. Foram considerados três horizontes, com um modelo direto para cada um: 15 minutos (<i>h</i> = 1), 1 hora (<i>h</i> = 4) e 3 horas "
  "(<i>h</i> = 12). Adicionalmente, avaliou-se a previsão para o dia seguinte usando apenas modelos de referência, pois o conjunto de informações disponível "
  "(sem dados meteorológicos) não sustenta um modelo mais elaborado nesse horizonte.")
H(3, "3.6.2 Atributos")
tabela("Atributos utilizados nos modelos de aprendizagem de máquina",
       ["Atributo", "Descrição"],
       [["pv0", "Potência FV média na janela de origem (W)"],
        ["pv_l1 … pv_l4", "Potência FV nas 4 janelas anteriores (defasagens de 15 a 60 min)"],
        ["pv_ma4, pv_ma12, pv_max4", "Média móvel de 1 h e de 3 h e máximo móvel de 1 h da potência FV"],
        ["pvV0", "Tensão do arranjo FV na origem (V)"],
        ["carga0", "Potência da carga na origem (W)"],
        ["soc0", "Estado de carga da bateria na origem (%)"],
        ["pv_dia_ant", "Potência FV no horário-alvo, 24 h antes (W)"],
        ["tod_sin, tod_cos", "Codificação cíclica (seno e cosseno) do horário-alvo"]],
       [5.0, 10.5], ["l", "l"])
H(3, "3.6.3 Modelos e hiperparâmetros")
P("Foram comparados cinco modelos: persistência, sazonal ingênuo (24 h) e climatologia (referências); regressão Ridge (padronização dos atributos, "
  "imputação de ausentes pela mediana e parâmetro de regularização igual a 1,0); e <i>Gradient Boosting</i> (<i>HistGradientBoostingRegressor</i>, com 200 "
  "iterações, taxa de aprendizado de 0,05, profundidade máxima 3, mínimo de 20 amostras por folha, regularização L2 de 1,0 e semente aleatória 42). "
  "Devido ao histórico curto, os hiperparâmetros não foram otimizados: adotaram-se valores conservadores, com o objetivo de reduzir o risco de "
  "sobreajuste. As previsões negativas foram truncadas em zero.")
H(3, "3.6.4 Validação")
P(f"Foi adotada a validação <i>walk-forward</i> diária. Para cada dia de teste D, de 13 a 20 de setembro de 2026 ({F['config']['n_dias_teste']} dias), os modelos "
  f"foram treinados apenas com observações cujo instante-alvo é anterior a D e avaliados nos instantes-alvo de D. O conjunto de treino cresceu de cerca de "
  f"600 janelas (primeiro dia de teste) para cerca de 1.300 (último). O limiar de horário diurno foi definido, em cada dobramento, a partir do treino "
  f"(janelas do dia com média histórica superior a {n(F['config']['limiar_diurno_W'])} W), e as métricas foram calculadas apenas nesses horários, para que os "
  f"valores nulos noturnos não reduzissem artificialmente o erro.")
H(3, "3.6.5 Métricas e importância dos atributos")
P("Foram calculados MAE, RMSE, nMAE e o ganho relativo à persistência, agregados sobre todos os dias de teste, e o MAE por dia (média e desvio-padrão "
  "entre dias). A importância dos atributos do <i>Gradient Boosting</i> (horizonte de 1 h) foi estimada por permutação: o aumento do MAE quando os "
  "valores de um atributo são embaralhados no conjunto de teste, com cinco repetições e média entre os dobramentos.")
H(2, "3.7 DETECÇÃO DE ANOMALIAS")
P(f"Como o estado do inversor é sempre <i>ONLINE</i> e não há registro rotulado de falhas, adotou-se uma abordagem não supervisionada. As leituras foram "
  f"agregadas por média em janelas de 1 minuto ({n(A['n_minutos'])} janelas válidas) e o <i>Isolation Forest</i> (300 árvores, contaminação de {pct(A['contaminacao'], 0)}, "
  f"semente 42) foi treinado com os dias 6 a 15 de setembro sobre 10 variáveis de operação (Apêndice A: tensão e frequência de saída, tensão, potência e SOC da bateria, "
  f"potência e carga percentual do inversor, potência aparente e potência e tensão FV), previamente padronizadas. O limiar de decisão foi fixado no "
  f"quantil de {pct(A['contaminacao'], 0)} dos escores de treino e aplicado, sem reajuste, aos dias 16 a 20 de setembro. Os resultados foram comparados com a regra atual do sistema "
  f"(tensão de saída CA menor que 212 V ou maior que 230 V), quanto à fração de minutos sinalizados e à sobreposição entre os dois critérios, e a "
  f"estabilidade foi avaliada pelo índice de Jaccard entre execuções com cinco sementes distintas. Como não há rótulos, não é possível calcular precisão "
  f"e revocação; a avaliação é, portanto, descritiva.")
H(2, "3.8 FERRAMENTAS E REPRODUTIBILIDADE")
P("O processamento foi realizado em Python, com as bibliotecas <i>pandas</i> (MCKINNEY, 2010), <i>NumPy</i>, <i>SciPy</i>, <i>scikit-learn</i> "
  "(PEDREGOSA et al., 2011) e <i>Matplotlib</i>, e o documento foi gerado por <i>script</i>, com todos os números lidos diretamente dos arquivos de "
  "resultados. Sementes aleatórias foram fixadas e os experimentos são determinísticos. Os <i>scripts</i> de extração, análise, previsão e detecção de anomalias "
  "compõem o repositório do projeto.")

# ------------------------------------------------------------------ 4 RESULTADOS
H(1, "4 RESULTADOS E DISCUSSÃO")
H(2, "4.1 ANÁLISE EXPLORATÓRIA")
H(3, "4.1.1 Qualidade e cobertura da base")
modo = E["modo_inversor"]
tot = sum(modo.values())
P(f"A base contém {n(E['n_registros'])} registros, de 6 de setembro de 2026 às 9 h 42 min até 20 de setembro às 20 h 11 min (horário local), "
  f"cobrindo {q['dias_calendario']} dias de calendário, sendo o primeiro e o último parciais. Não há instantes duplicados nem valores ausentes nas variáveis numéricas. "
  f"O intervalo mediano entre leituras é de {n(q['intervalo_mediana_s'], 1)} s (percentil 95: {n(q['intervalo_p95_s'], 1)} s), e foram identificadas "
  f"{q['lacunas_gt_60s']} lacunas maiores que 60 s, das quais {q['lacunas_gt_600s']} maiores que 10 minutos (a maior com {n(q['intervalo_max_s'] / 60)} minutos). "
  f"O inversor operou no modo SBU em {pct(modo['SBU'] / tot)} das leituras e no modo SUB em {pct(modo['SUB'] / tot)}; apenas {modo.get('GRID', 0)} leituras "
  f"registraram outro modo, e o estado foi <i>ONLINE</i> em todas as leituras. Esses modos são reportados pelo próprio equipamento e, conforme a "
  f"nomenclatura usual de inversores híbridos, representam prioridades de fonte de saída (a confirmar com o responsável pelo sistema).")
tabela("Qualidade da base extraída",
       ["Indicador", "Valor"],
       [["Registros extraídos", n(E["n_registros"])],
        ["Período (horário local)", "06/09/2026 09:42 a 20/09/2026 20:11"],
        ["Instantes duplicados", str(q["timestamps_duplicados"])],
        ["Valores ausentes (variáveis numéricas)", str(q["nulos_total"])],
        ["Intervalo mediano entre leituras", f"{n(q['intervalo_mediana_s'], 1)} s"],
        ["Lacunas > 60 s / > 10 min", f"{q['lacunas_gt_60s']} / {q['lacunas_gt_600s']}"],
        ["Janelas de 15 min válidas", f"{n(E['serie_15min']['linhas_validas'])} de {n(E['serie_15min']['linhas'])}"]],
       [8.0, 7.5], ["l", "c"])
H(3, "4.1.2 Estatística descritiva")
D = E["descritiva"]
rot = [("pvPower", "Potência FV (W)", 0), ("pvVoltage", "Tensão FV (V)", 1), ("pvCurrent", "Corrente FV (A)", 1), ("acOutputVoltage", "Tensão de saída CA (V)", 1),
       ("acOutputFrequency", "Frequência de saída (Hz)", 2), ("batteryVoltage", "Tensão da bateria (V)", 1), ("batterySOC", "SOC da bateria (%)", 0),
       ("loadWatts", "Potência da carga (W)", 0), ("inverterLoadPercent", "Carga do inversor (%)", 0)]
tabela("Estatísticas descritivas das principais variáveis (leituras brutas)",
       ["Variável", "Média", "Desvio-padrão", "Mínimo", "Mediana", "Máximo"],
       [[nm, n(D[k]["mean"], d), n(D[k]["std"], d), n(D[k]["min"], d), n(D[k]["50%"], d), n(D[k]["max"], d)] for k, nm, d in rot],
       [5.0, 2.0, 2.6, 2.0, 2.0, 2.0])
P(f"A potência FV atingiu {n(D['pvPower']['max'])} W, com média de {n(D['pvPower']['mean'])} W (incluindo as horas noturnas, em que é nula). A carga "
  f"média foi de {n(D['loadWatts']['mean'])} W e a mediana de {n(D['loadWatts']['50%'])} W, contra um máximo de {n(D['loadWatts']['max'])} W, o que "
  f"indica um perfil de consumo com base baixa e picos ocasionais elevados — resultado relevante para a detecção de anomalias (seção 4.3). A tensão de saída "
  f"apresentou média de {n(D['acOutputVoltage']['mean'], 1)} V, com mínimo de {n(D['acOutputVoltage']['min'], 1)} V e máximo de {n(D['acOutputVoltage']['max'], 1)} V. "
  f"Entre as correlações, destaca-se a da potência FV com a corrente FV (r = {n(E['correlacao']['pvPower']['pvCurrent'], 2)}), esperada pela relação "
  f"P = V·I, e com a potência da bateria (r = {n(E['correlacao']['pvPower']['batteryPower'], 2)}), pois a geração excedente carrega a bateria; a correlação "
  f"com a potência da carga foi praticamente nula (r = {n(E['correlacao']['pvPower']['loadWatts'], 2)}).")
H(3, "4.1.3 Padrão diário e energia gerada")
raw = pd.read_csv(sorted((ROOT / "data" / "raw").glob("telemetries_*.csv"))[-1], parse_dates=["timestamp"])
raw["timestamp"] -= pd.Timedelta(hours=3)
hh = raw.groupby(raw.timestamp.dt.hour).pvPower
horas = [int(k) for k, v in hh.mean().items() if v > 5]
pico_h, pico_v = int(hh.mean().idxmax()), hh.mean().max()
P(f"A Figura 2 mostra a potência FV ao longo do período e a Figura 3, o perfil horário. A geração ocorre entre {horas[0]} h e {horas[-1] + 1} h locais, "
  f"faixa compatível com o nascer e o pôr do sol na região no período, o que corrobora a conversão de horário adotada. A média da potência é máxima às "
  f"{pico_h} h ({n(pico_v)} W) e decresce ao longo da tarde. A dispersão entre dias é grande (intervalo interquartil amplo), refletindo a "
  f"variabilidade da nebulosidade. O pico de manhã e o declínio à tarde são compatíveis com a orientação dos painéis e com a limitação de potência "
  f"pelo inversor quando a bateria está cheia (seção 4.1.4); a base atual não permite separar os dois efeitos.")
figura("fig_pv_periodo.png", "Potência FV média em janelas de 15 minutos (6 a 20 de setembro de 2026)")
figura("fig_perfil_horario.png", "Perfil horário da potência FV (mediana, intervalo interquartil e média)")
ed = E["energia_diaria_kwh"]
_fora = raw[(raw.acOutputVoltage < 212) | (raw.acOutputVoltage > 230)]
FORA_SUB = float((_fora.inverterMode == "SUB").mean())
FORA_REDE = float((_fora.acInputVoltage >= 100).mean())
completos = [k for k in ed if k != "2026-09-06"]
pvs = [ed[k]["pv"] for k in completos]
kmax = max(completos, key=lambda k: ed[k]["pv"]); kmin = min(completos, key=lambda k: ed[k]["pv"])
P(f"A energia FV diária estimada, considerando os dias completos (7 a 20 de setembro), teve média de {n(sum(pvs) / len(pvs), 1)} kWh, variando de "
  f"{n(ed[kmin]['pv'], 1)} kWh ({dia_br(kmin)}) a {n(ed[kmax]['pv'], 1)} kWh ({dia_br(kmax)}), isto é, o melhor dia gerou cerca de "
  f"{n(ed[kmax]['pv'] / ed[kmin]['pv'], 1)} vezes a energia do pior (Tabela 4 e Figura 4). Essa amplitude reforça que a previsão em horizontes "
  f"superiores a poucas horas depende de informações de nebulosidade.")
tabela("Energia diária estimada (kWh)",
       ["Dia (2026)", "Energia FV", "Energia da carga", "Observação"],
       [[dia_br(k), n(ed[k]["pv"], 2), n(ed[k]["carga"], 2), "parcial (início às 9 h 42)" if k == "2026-09-06" else ("parcial (fim às 20 h 11)" if k == "2026-09-20" else "—")] for k in ed],
       [3.0, 3.2, 3.6, 5.7], ["c", "c", "c", "l"],
       nota="Nota: estimativa por integração da potência (intervalo entre leituras limitado a 60 s). No dia 20/09 a geração FV já estava encerrada no fim do registro.")
figura("fig_energia_diaria.png", "Energia FV diária estimada (kWh)")
H(3, "4.1.4 Estado de carga da bateria e limitação da geração")
ls = E["limitacao_soc"]
dmax = max(ls["por_dia"], key=ls["por_dia"].get)
P(f"Nas janelas diurnas com potência FV superior a 50 W, o SOC da bateria esteve em 99% ou mais em {pct(ls['fracao'])} dos casos "
  f"({ls['linhas_soc_ge_99']} de {ls['linhas_diurnas']}), com grande variação entre dias: de 0% em vários dias até {pct(ls['por_dia'][dmax], 0)} em {dia_br(dmax)}. "
  f"Nessas janelas, a razão mediana entre potência FV e potência da carga foi de {n(ls['razao_pv_carga_mediana'], 2)}, isto é, a potência FV passa a "
  f"acompanhar de perto o consumo. A Figura 5 ilustra o fenômeno em 19/09: a potência FV, que chegava a cerca de 1.900 W, cai abruptamente por volta "
  f"das 11 h 30 min, quando o SOC retorna a 100% após um pico de consumo, e passa a se manter pouco acima da potência da carga, ainda que o período do dia fosse "
  f"de alta irradiância. O comportamento é "
  f"compatível com a redução da potência extraída dos painéis pelo inversor quando não há para onde direcionar a energia excedente (seção 2.1).")
figura("fig_limitacao_soc.png", "Potência FV, potência da carga e SOC em 19/09/2026 (hora local)", 14.5)
P("Esse resultado tem duas consequências para o projeto. Primeiro, a série de potência FV medida <i>censura</i> a geração potencial nos períodos de bateria "
  "cheia: um modelo treinado nela aprende parte da política de controle do inversor, e não apenas a física da geração. Isso explica por que o SOC "
  "figura entre os atributos úteis à previsão (seção 4.2.2) e sugere que, para estimar a geração disponível, serão necessários tratamentos específicos "
  "(seção 5). Segundo, indicadores de eficiência calculados sobre a potência medida devem desconsiderar os períodos de bateria cheia.")
H(3, "4.1.5 Ausência de rede elétrica")
ar = A["ausencia_rede"]
P(f"Foram identificados {ar['episodios']} episódios de ausência de tensão na entrada CA (tensão de entrada inferior a 100 V), totalizando {n(ar['minutos_total'])} minutos, com "
  f"duração mediana de {n(ar['duracao_mediana_min'], 0)} minutos e máxima de {n(ar['duracao_max_min'])} minutos. Esses episódios não são falhas do sistema "
  f"fotovoltaico, mas mudam o regime de operação (o consumo passa a ser atendido exclusivamente por painéis e bateria) e, por isso, devem ser tratados "
  f"como um regime próprio nas etapas seguintes.")

H(2, "4.2 PREVISÃO DA POTÊNCIA FOTOVOLTAICA")
H(3, "4.2.1 Desempenho por horizonte")
linhas = []
for hn in ("15 min", "1 h", "3 h"):
    for m in ("Persistência", "Sazonal ingênuo (24 h)", "Climatologia", "Ridge", "Gradient Boosting"):
        r = F["horizonte"][hn][m]
        linhas.append([hn if m == "Persistência" else "", m, n(r["mae"]), n(r["rmse"]), n(r["nmae"], 2), ("—" if m == "Persistência" else f"{n(r['skill_vs_persistencia'], 2)}")])
tabela("Desempenho dos modelos por horizonte de previsão (validação walk-forward, 8 dias de teste)",
       ["Horizonte", "Modelo", "MAE (W)", "RMSE (W)", "nMAE", "Ganho vs. persistência"],
       linhas, [2.0, 4.6, 2.0, 2.2, 1.7, 3.0], ["c", "l", "c", "c", "c", "c"],
       nota=f"Nota: métricas em horários diurnos, agregadas sobre {F['config']['n_dias_teste']} dias de teste. O nMAE é o MAE dividido pela média diurna da potência observada "
            f"(cerca de {n(F['horizonte']['1 h']['media_y_diurna_W'])} W). Ganho = 1 − MAE/MAE da persistência.")
b = F["horizonte"]["15 min"]
c3 = F["horizonte"]["3 h"]
P(f"No horizonte de <b>15 minutos</b>, a persistência foi o melhor modelo (MAE de {n(b['Persistência']['mae'])} W), seguida da regressão Ridge "
  f"({n(b['Ridge']['mae'])} W) e do <i>Gradient Boosting</i> ({n(b['Gradient Boosting']['mae'])} W), que não a superaram. Esse resultado é esperado: em janelas "
  f"contíguas, o valor mais recente é uma previsão muito informativa, e os modelos aprendidos, com poucas centenas de observações, acrescentam ruído de "
  f"estimação sem ganho compensador.")
P(f"No horizonte de <b>1 hora</b>, o <i>Gradient Boosting</i> apresentou o menor erro (MAE de {n(gb1['mae'])} W e RMSE de {n(gb1['rmse'])} W), com ganho de "
  f"{pct(gb1['skill_vs_persistencia'], 0)} sobre a persistência ({n(pe1['mae'])} W) e de {pct(1 - gb1['mae'] / h1['Climatologia']['mae'], 0)} sobre a climatologia "
  f"({n(h1['Climatologia']['mae'])} W). A regressão Ridge obteve {n(h1['Ridge']['mae'])} W. O <i>Gradient Boosting</i> teve erro médio diário inferior ao da persistência em "
  f"{F['gb_vence_persistencia_1h_dias']} dos {F['n_dias_avaliados_1h']} dias avaliados (Apêndice B), o que sugere que a vantagem não decorre de um único dia; "
  f"ainda assim, com oito dias não é possível afirmar significância estatística. O erro normalizado, de {n(gb1['nmae'], 2)}, mostra que o erro permanece "
  f"elevado em relação à potência diurna média, algo previsível diante da variabilidade das nuvens.")
P(f"No horizonte de <b>3 horas</b>, a climatologia obteve o menor erro ({n(c3['Climatologia']['mae'])} W), à frente do <i>Gradient Boosting</i> "
  f"({n(c3['Gradient Boosting']['mae'])} W) e da Ridge ({n(c3['Ridge']['mae'])} W); todos superaram amplamente a persistência ({n(c3['Persistência']['mae'])} W), que "
  f"deteriora rapidamente com o horizonte. Isso indica que, a três horas, a informação contida nos valores recentes da própria série já é pouco útil e que "
  f"o comportamento diurno médio domina — resultado condizente com a ausência de dados meteorológicos.")
figura("fig_mae_modelos.png", "Erro absoluto médio (MAE) dos modelos por horizonte de previsão")
figura("fig_previsao_1h.png", "Potência FV observada e prevista com horizonte de 1 hora nos três últimos dias de teste", 15.5)
P("A Figura 7 ilustra o comportamento dos modelos. A persistência reproduz a série com defasagem de uma hora (atraso visível nas subidas e descidas), "
  "enquanto o <i>Gradient Boosting</i> antecipa parte das variações. Ambos falham nas quedas abruptas de potência (por exemplo, em 19/09, quando a bateria "
  "atinge carga plena), pois esse evento não é antecipável a partir das variáveis disponíveis.".replace("Figura 7", "Figura 7"))
H(3, "4.2.2 Importância dos atributos")
imp = list(F["importancia_permutacao_1h"].items())
P(f"Pela importância por permutação (Figura 8), os atributos mais relevantes para o <i>Gradient Boosting</i> com horizonte de 1 h foram a potência FV atual "
  f"(<i>pv0</i>, aumento de {n(imp[0][1])} W no MAE ao ser permutada) e a codificação do horário do dia (<i>tod_sin</i> e <i>tod_cos</i>), seguidas da "
  f"média móvel de 3 h e do SOC da bateria. Os demais atributos tiveram contribuição próxima de zero ou negativa, indicando redundância entre defasagens e "
  f"estatísticas móveis. Cabe ressaltar que a importância por permutação é sensível à correlação entre atributos e foi calculada em poucos dias; ela deve "
  f"ser lida como indicativa.")
figura("fig_importancia.png", "Importância por permutação dos atributos (Gradient Boosting, horizonte de 1 h)", 14.0)
H(3, "4.2.3 Previsão para o dia seguinte")
da = F["day_ahead"]
tabela("Previsão para o dia seguinte com modelos de referência",
       ["Modelo", "MAE diurno médio (W)", "Erro médio da energia diária (kWh)"],
       [["Sazonal ingênuo (dia anterior)", n(da["mae_medio"]["mae_naive"]), n(da["erro_energia_kwh_medio"]["naive"], 2)],
        ["Média dos 3 dias anteriores", n(da["mae_medio"]["mae_ma3"]), n(da["erro_energia_kwh_medio"]["ma3"], 2)],
        ["Climatologia (todos os dias anteriores)", n(da["mae_medio"]["mae_clim"]), n(da["erro_energia_kwh_medio"]["clim"], 2)]],
       [6.4, 4.4, 4.7], ["l", "c", "c"],
       nota=f"Nota: média sobre {F['config']['n_dias_teste']} dias de teste; energia diária observada média de {n(da['energia_real_media_kwh'], 2)} kWh.")
P(f"Sem informações meteorológicas, o melhor resultado foi o da climatologia, com MAE diurno de {n(da['mae_medio']['mae_clim'])} W e erro médio de "
  f"{n(da['erro_energia_kwh_medio']['clim'], 2)} kWh na energia diária, o que corresponde a cerca de {pct(da['erro_energia_kwh_medio']['clim'] / da['energia_real_media_kwh'], 0)} "
  f"da energia diária observada média ({n(da['energia_real_media_kwh'], 2)} kWh). O sazonal ingênuo, que repete o dia anterior, foi pior "
  f"(erro de {n(da['erro_energia_kwh_medio']['naive'], 2)} kWh), pois a nebulosidade de um dia é pouco informativa sobre a do seguinte. A Figura 9 compara a energia observada "
  f"com as previsões. Conclui-se que, para o dia seguinte, não há ganho a esperar de modelos baseados apenas na própria série; a incorporação de dados "
  f"meteorológicos é a via natural de melhoria.")
figura("fig_day_ahead.png", "Energia FV diária observada e prevista para o dia seguinte (modelos de referência)")

H(2, "4.3 DETECÇÃO DE ANOMALIAS")
sb = A["sobreposicao_teste"]
rg = A["regra_atual"]
P(f"No período completo ({n(A['n_minutos'])} minutos), a regra atual de tensão de saída (fora de 212–230 V) sinalizou {n(rg['minutos_alertados'])} minutos "
  f"({pct(rg['fracao'])}). Na etapa de teste (16 a 20 de setembro), o <i>Isolation Forest</i>, com limiar definido no treino, sinalizou {n(A['minutos_sinalizados_teste'])} minutos "
  f"({pct(A['fracao_sinalizada_teste'], 2)} dos {n(A['n_teste'])} minutos), em {A['episodios_teste']} episódios, valor próximo da contaminação de "
  f"{pct(A['contaminacao'], 0)} definida, o que indica que o limiar aprendido no treino se transfere de forma razoável ao período seguinte. A estabilidade entre "
  f"sementes foi boa (índice de Jaccard médio de {n(A['jaccard_medio_entre_sementes'], 2)}).")
tabela("Comparação entre a regra atual e o Isolation Forest no período de teste (16 a 20/09)",
       ["Critério", "Minutos sinalizados"],
       [["Regra atual (tensão de saída fora de 212–230 V)", n(sb["regra_total"])],
        ["Isolation Forest (contaminação de 1%)", n(A["minutos_sinalizados_teste"])],
        ["Sinalizados pelos dois critérios", n(sb["regra_e_modelo"])],
        ["Somente pela regra atual", n(sb["so_regra"])],
        ["Somente pelo Isolation Forest", n(sb["so_modelo"])]],
       [10.5, 5.0], ["l", "c"])
z = A["perfil_sinalizados_z_medio"]
eps = A["episodios_lista"]
cmin, cmax = min(e["loadWatts_max"] for e in eps), max(e["loadWatts_max"] for e in eps)
P(f"O ponto central é a <b>sobreposição desprezível</b> entre os critérios: apenas {sb['regra_e_modelo']} minutos foram sinalizados por ambos. O perfil dos minutos "
  f"sinalizados pelo modelo mostra desvios médios de aproximadamente {n(z['loadWatts'], 1)} desvios-padrão na potência da carga (e de "
  f"{n(z['inverterLoadPercent'], 1)} na carga percentual do inversor), e a potência da bateria fortemente negativa ({n(z['batteryPower'], 1)} desvios), ou seja, os episódios "
  f"correspondem a picos de consumo (potência máxima de {n(cmin)} a {n(cmax)} W nos episódios, contra mediana de {n(D['loadWatts']['50%'])} W), atendidos "
  f"com descarga da bateria. Já a tensão de saída, que motiva a regra atual, praticamente não contribuiu para os escores do modelo. A Figura 10 mostra que os "
  f"afundamentos de tensão abaixo de 212 V ocorrem em períodos prolongados (por exemplo, em torno da virada de 17 para 18/09 e ao longo de 20/09), nos quais o modelo não sinaliza nada, "
  f"enquanto os pontos sinalizados situam-se em picos de carga com tensão normal.")
figura("fig_anomalias.png", "Tensão de saída CA e potência da carga no período de teste, com minutos sinalizados pelo Isolation Forest", 15.0)
P("Interpretação: as duas abordagens detectam fenômenos distintos e complementares. A regra atual captura excursões de tensão de saída, que ocorrem "
  f"em longos intervalos — {pct(FORA_SUB, 0)} das leituras fora da faixa ocorreram no modo SUB e {pct(FORA_REDE, 0)} com a rede presente, o que sugere relação com a tensão da rede quando esta alimenta a saída; o <i>Isolation Forest</i>, nas variáveis e na "
  "configuração adotadas, captura combinações incomuns de carga e descarga da bateria. Um pico de consumo não é, por si só, uma falha; portanto, os "
  "minutos sinalizados devem ser entendidos como <i>eventos atípicos</i>, e não como defeitos confirmados. Sem rótulos de falha, não se pode afirmar qual "
  "abordagem é melhor; é possível afirmar apenas que não são equivalentes.")

H(2, "4.4 DISCUSSÃO")
P("Os resultados mostram um quadro consistente com a teoria de previsão solar: (i) no horizonte mais curto, a persistência é difícil de superar; "
  "(ii) em horizonte intermediário (1 h), há espaço para modelos que combinam o valor atual, tendências recentes e o horário do dia; e (iii) em horizontes "
  "maiores (3 h e dia seguinte), sem informação meteorológica, o comportamento médio (climatologia) é o melhor preditor disponível. O ganho do "
  "<i>Gradient Boosting</i> em 1 h é modesto em termos absolutos (o nMAE permanece próximo de 0,5) e assenta em uma base de apenas 15 dias — insuficiente, por exemplo, para "
  "capturar variações sazonais. Deve-se, portanto, ler o resultado como uma demonstração de viabilidade da abordagem e como uma linha de base a ser superada pelas próximas versões, "
  "não como um modelo pronto para operação.")
P("A descoberta de que o inversor limita a potência FV com a bateria cheia é, possivelmente, o achado mais relevante para o projeto: ela mostra que "
  "os dados brutos de telemetria não medem diretamente a grandeza física que se deseja prever e que a modelagem deve considerar o estado de operação do "
  "equipamento. Do mesmo modo, a comparação entre a regra de tensão e o <i>Isolation Forest</i> evidencia que a definição de “anomalia” precisa ser combinada com "
  "o responsável pelo sistema (o que é falha, o que é evento normal) antes de qualquer avaliação quantitativa.")
H(2, "4.5 LIMITAÇÕES E AMEAÇAS À VALIDADE")
A_([
    "<b>Histórico curto</b>: 15 dias, dos quais o primeiro é parcial, sem cobertura de outras estações do ano e com apenas 8 dias de teste, o que limita a generalização e a robustez estatística das comparações",
    "<b>Ausência de dados meteorológicos</b>: irradiância, nebulosidade e temperatura não constam da base, o que limita a previsão para além de poucas horas",
    "<b>Censura da geração pela bateria cheia</b>: a variável-alvo mistura geração disponível e política de controle do inversor",
    "<b>Sem rótulos de anomalia</b>: impossibilita calcular precisão e revocação; a avaliação é descritiva",
    "<b>Hiperparâmetros não otimizados</b>: adotados valores conservadores por causa do tamanho da amostra",
    "<b>Sistema único</b>: os resultados descrevem uma instalação específica",
    "<b>Interpretação dos modos do inversor</b> (SBU/SUB) baseada em nomenclatura usual e ainda não confirmada pelo responsável pelo sistema",
])

# ------------------------------------------------------------------ 5 CONSIDERAÇÕES
H(1, "5 CONSIDERAÇÕES PARCIAIS E PRÓXIMAS ETAPAS")
P(f"Esta primeira versão cumpriu o objetivo de caracterizar a base do SolarSync e de estabelecer linhas de base de previsão e de detecção de anomalias com "
  f"dados reais. Verificou-se que a base tem boa qualidade (sem duplicidades ou ausências, intervalo mediano de {n(q['intervalo_mediana_s'], 1)} s), mas cobertura temporal curta. "
  f"Na previsão da potência FV, o <i>Gradient Boosting</i> superou a persistência no horizonte de 1 hora (MAE de {n(gb1['mae'])} W contra {n(pe1['mae'])} W), "
  f"enquanto a persistência foi imbatível em 15 minutos e a climatologia foi a melhor referência em 3 horas e no dia seguinte. Na detecção de anomalias, o "
  f"<i>Isolation Forest</i> sinalizou picos de carga e mostrou-se complementar, e não equivalente, à regra atual de tensão. Como achado transversal, constatou-se "
  f"que a potência FV medida é limitada pelo inversor quando a bateria está cheia.")
P("As próximas etapas, alinhadas ao Plano de Ação (Quinzenas 5 a 7), são:")
A_([
    "incorporar dados meteorológicos de fonte aberta (irradiância e nebulosidade) como variáveis exógenas, avaliando o ganho nos horizontes de 3 h e do dia seguinte",
    "tratar a censura por bateria cheia, sinalizando os períodos de SOC pleno e avaliando modelos da geração potencial ou modelos separados por regime",
    "ampliar o histórico com a coleta contínua e repetir a validação <i>walk-forward</i> com mais dias, incluindo a otimização de hiperparâmetros com validação aninhada e a estimativa de intervalos de previsão",
    "definir com o responsável pelo sistema o que constitui uma anomalia, rotular eventos conhecidos e comparar variantes (modelos por regime, <i>Local Outlier Factor</i>, <i>autoencoder</i>) com métricas de precisão e revocação",
    "integrar os resultados (previsões e alertas) ao <i>dashboard</i> existente, mediante endpoint de leitura, mantendo o acesso ao banco somente para leitura",
    "consolidar o Relatório Final, a Avaliação Colaborativa e o vídeo de apresentação",
])

# ------------------------------------------------------------------ REFERÊNCIAS
H(1, "REFERÊNCIAS").alignment = WD_ALIGN_PARAGRAPH.CENTER
REFS = [
    "BERGMEIR, C.; BENÍTEZ, J. M. On the use of cross-validation for time series predictor evaluation. <b>Information Sciences</b>, v. 191, p. 192–213, 2012. DOI: 10.1016/j.ins.2011.12.028.",
    "BRASIL. <b>Lei nº 14.300, de 6 de janeiro de 2022</b>. Institui o marco legal da microgeração e minigeração distribuída, o Sistema de Compensação de Energia Elétrica (SCEE) e o Programa de Energia Renovável Social (PERS); altera as Leis nºs 10.848, de 15 de março de 2004, e 9.427, de 26 de dezembro de 1996; e dá outras providências. <b>Diário Oficial da União</b>: seção 1, Brasília, DF, 7 jan. 2022.",
    "CHANDOLA, V.; BANERJEE, A.; KUMAR, V. Anomaly detection: a survey. <b>ACM Computing Surveys</b>, v. 41, n. 3, art. 15, p. 1–58, 2009.",
    "FRIEDMAN, J. H. Greedy function approximation: a gradient boosting machine. <b>The Annals of Statistics</b>, v. 29, n. 5, p. 1189–1232, 2001.",
    "GRUPO SOLARSYNC. <b>Projeto Integrador IV: Sistema de Monitoramento de Geração Solar – SolarSync</b>. [S. l.]: Univesp, 2026. Documento técnico do projeto.",
    "HASTIE, T.; TIBSHIRANI, R.; FRIEDMAN, J. <b>The elements of statistical learning</b>: data mining, inference, and prediction. 2. ed. New York: Springer, 2009.",
    "HOERL, A. E.; KENNARD, R. W. Ridge regression: biased estimation for nonorthogonal problems. <b>Technometrics</b>, v. 12, n. 1, p. 55–67, 1970.",
    "HYNDMAN, R. J.; ATHANASOPOULOS, G. <b>Forecasting</b>: principles and practice. 3. ed. Melbourne: OTexts, 2021. Disponível em: https://otexts.com/fpp3/. Acesso em: 20 set. 2026.",
    "HYNDMAN, R. J.; KOEHLER, A. B. Another look at measures of forecast accuracy. <b>International Journal of Forecasting</b>, v. 22, n. 4, p. 679–688, 2006.",
    "INMAN, R. H.; PEDRO, H. T. C.; COIMBRA, C. F. M. Solar forecasting methods for renewable energy integration. <b>Progress in Energy and Combustion Science</b>, v. 39, n. 6, p. 535–576, 2013.",
    "KE, G. et al. LightGBM: a highly efficient gradient boosting decision tree. In: ADVANCES IN NEURAL INFORMATION PROCESSING SYSTEMS, 30., 2017, Long Beach. <b>Proceedings</b> [...]. Red Hook: Curran Associates, 2017. p. 3146–3154.",
    "LIU, F. T.; TING, K. M.; ZHOU, Z.-H. Isolation forest. In: IEEE INTERNATIONAL CONFERENCE ON DATA MINING, 8., 2008, Pisa. <b>Proceedings</b> [...]. Piscataway: IEEE, 2008. p. 413–422.",
    "MCKINNEY, W. Data structures for statistical computing in Python. In: PYTHON IN SCIENCE CONFERENCE, 9., 2010, Austin. <b>Proceedings</b> [...]. [S. l.: s. n.], 2010. p. 56–61.",
    "MODBUS ORGANIZATION. <b>MODBUS application protocol specification V1.1b3</b>. [S. l.]: Modbus Organization, 2012.",
    "PEDREGOSA, F. et al. Scikit-learn: machine learning in Python. <b>Journal of Machine Learning Research</b>, v. 12, p. 2825–2830, 2011.",
]
for r in REFS:
    P(r, "Referencia")

# ------------------------------------------------------------------ APÊNDICES
ap = doc.add_heading("APÊNDICE A – DICIONÁRIO DE VARIÁVEIS DA TELEMETRIA", level=1)
ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
P("Variáveis da coleção <i>telemetries</i> utilizadas neste trabalho. As descrições foram inferidas dos nomes dos campos, da documentação do sistema e do "
  "comportamento dos dados, e devem ser validadas com o responsável pelo SolarSync.")
tabela("Variáveis utilizadas",
       ["Campo", "Descrição", "Unidade"],
       [["timestamp", "Instante da leitura (armazenado em UTC)", "—"],
        ["pvVoltage / pvCurrent / pvPower", "Tensão, corrente e potência do arranjo fotovoltaico", "V / A / W"],
        ["acInputVoltage / acInputFrequency", "Tensão e frequência da entrada CA (rede)", "V / Hz"],
        ["acOutputVoltage / acOutputFrequency", "Tensão e frequência da saída CA", "V / Hz"],
        ["batteryVoltage / batteryCurrent", "Tensão e corrente da bateria", "V / A"],
        ["batteryPower", "Potência da bateria (positiva na carga; inferido dos dados)", "W"],
        ["batterySOC", "Estado de carga da bateria", "%"],
        ["loadWatts / loadVA", "Potência ativa e aparente da carga", "W / VA"],
        ["inverterLoadPercent", "Carga do inversor em relação à capacidade", "%"],
        ["inverterMode", "Modo de operação reportado (SBU, SUB, GRID)", "—"],
        ["inverterStatus", "Estado do inversor (ONLINE em todas as leituras)", "—"]],
       [5.6, 7.9, 2.0], ["l", "l", "c"])
ap = doc.add_heading("APÊNDICE B – ERRO ABSOLUTO MÉDIO DIÁRIO (HORIZONTE DE 1 HORA)", level=1)
ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
mods = ["Persistência", "Sazonal ingênuo (24 h)", "Climatologia", "Ridge", "Gradient Boosting"]
dias = list(F["horizonte"]["1 h"]["Persistência"]["mae_por_dia"].keys())
lin = [[dia_br(d)] + [n(F["horizonte"]["1 h"][m]["mae_por_dia"][d]) for m in mods] for d in dias]
lin.append(["Média"] + [n(F["horizonte"]["1 h"][m]["mae_diario_media"]) for m in mods])
lin.append(["Desvio-padrão"] + [n(F["horizonte"]["1 h"][m]["mae_diario_dp"]) for m in mods])
tabela("MAE diário (W) por modelo, horizonte de 1 hora",
       ["Dia (2026)", "Persistência", "Sazonal 24 h", "Climatologia", "Ridge", "Gradient Boosting"],
       lin, [2.6, 2.6, 2.6, 2.6, 2.2, 3.0])

# ------------------------------------------------------------------ configurações finais
# títulos de Referências/Apêndices sem numeração já ficam em Heading 1 (aparecem no sumário)
OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUT)
print("Gravado:", OUT)
