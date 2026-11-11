"""Gera a apresentação "Projeto de Modernização Tecnológica - Doce Encanto Confeitaria".

Uso: python3 gerar_apresentacao.py  ->  cria modernizacao_doce_encanto.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

# ---------- Identidade visual ----------
PLUM = RGBColor(0x3D, 0x1F, 0x4F)      # cor principal
CORAL = RGBColor(0xE0, 0x7A, 0x5F)     # destaque
MINT = RGBColor(0x2A, 0x9D, 0x8F)      # positivo
GOLD = RGBColor(0xE9, 0xB5, 0x4A)      # atenção
RED = RGBColor(0xC8, 0x44, 0x4B)       # negativo
BG = RGBColor(0xFA, 0xF7, 0xF2)        # fundo
CARD = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x2B, 0x2B, 0x33)
MUTED = RGBColor(0x6B, 0x65, 0x70)
LINE = RGBColor(0xE6, 0xDF, 0xE8)
FONT = "Arial"
EMPRESA = "Doce Encanto Confeitaria"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.06
    return s


def text(slide, x, y, w, h, runs, size=12, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.1):
    """runs: str ou lista de parágrafos; cada parágrafo é str ou lista de (texto, {estilo})."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = anchor
    paras = runs if isinstance(runs, list) else [runs]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = spacing
        parts = p if isinstance(p, list) else [(p, {})]
        for t, st in parts:
            r = para.add_run()
            r.text = t
            f = r.font
            f.name = FONT
            f.size = Pt(st.get("size", size))
            f.bold = st.get("bold", bold)
            f.italic = st.get("italic", False)
            f.color.rgb = st.get("color", color)
        if isinstance(p, list) and p and "space" in p[0][1]:
            para.space_before = Pt(p[0][1]["space"])
    return tb


def background(slide, color=BG):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def header(slide, num, title, subtitle):
    background(slide)
    rect(slide, 0, 0, 13.333, 1.05, PLUM)
    rect(slide, 0, 1.05, 13.333, 0.06, CORAL)
    badge = rect(slide, 0.45, 0.25, 0.6, 0.6, CORAL, MSO_SHAPE.OVAL)
    text(slide, 0.45, 0.25, 0.6, 0.6, f"{num:02d}", 16, CARD, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(slide, 1.25, 0.14, 10.5, 0.5, title, 24, CARD, True)
    text(slide, 1.25, 0.6, 10.5, 0.35, subtitle, 12, RGBColor(0xE8, 0xD9, 0xEE))
    text(slide, 9.8, 0.38, 3.1, 0.3, EMPRESA, 10, RGBColor(0xE8, 0xD9, 0xEE), align=PP_ALIGN.RIGHT)
    # rodapé
    text(slide, 0.45, 7.1, 8, 0.3, "Projeto de Modernização Tecnológica  •  dados fictícios", 9, MUTED)
    text(slide, 11.9, 7.1, 1, 0.3, f"{num} / 12", 9, MUTED, align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body, accent=CORAL, title_size=13, body_size=10.5, icon=None):
    rect(slide, x, y, w, h, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
    rect(slide, x, y + 0.12, 0.07, h - 0.24, accent)
    tx = x + 0.22
    if icon:
        circ = rect(slide, x + 0.22, y + 0.15, 0.42, 0.42, accent, MSO_SHAPE.OVAL)
        text(slide, x + 0.22, y + 0.15, 0.42, 0.42, icon, 12, CARD, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        text(slide, x + 0.72, y + 0.14, w - 0.85, 0.45, title, title_size, PLUM, True, anchor=MSO_ANCHOR.MIDDLE)
    else:
        text(slide, tx, y + 0.12, w - 0.35, 0.4, title, title_size, PLUM, True)
    text(slide, tx, y + 0.62, w - 0.35, h - 0.7, body, body_size, INK)


def table(slide, x, y, w, data, col_w, row_h=0.36, size=10.5, header_fill=PLUM,
          bold_rows=(), fills=None, aligns=None):
    rows, cols = len(data), len(data[0])
    gt = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(row_h * rows))
    tbl = gt.table
    for j, cw in enumerate(col_w):
        tbl.columns[j].width = Inches(cw)
    for i in range(rows):
        tbl.rows[i].height = Inches(row_h)
        for j in range(cols):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.08)
            c.margin_top = c.margin_bottom = Inches(0.03)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            if i == 0:
                c.fill.fore_color.rgb = header_fill
            elif fills and (i, j) in fills:
                c.fill.fore_color.rgb = fills[(i, j)]
            elif fills and (i, None) in fills:
                c.fill.fore_color.rgb = fills[(i, None)]
            else:
                c.fill.fore_color.rgb = CARD if i % 2 else RGBColor(0xF5, 0xF0, 0xF6)
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = (aligns[j] if aligns else PP_ALIGN.LEFT)
            r = p.add_run()
            r.text = str(data[i][j])
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = i == 0 or i in bold_rows
            r.font.color.rgb = CARD if i == 0 else INK
    return tbl


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t


R, C = PP_ALIGN.RIGHT, PP_ALIGN.CENTER
L = PP_ALIGN.LEFT

# =====================================================================
# SLIDE 1 – Capa
# =====================================================================
s = prs.slides.add_slide(BLANK)
background(s, PLUM)
rect(s, 8.6, 0, 4.733, 7.5, RGBColor(0x4E, 0x2A, 0x63))
for (cx, cy, d, col) in [(9.4, 1.1, 2.6, CORAL), (11.2, 3.4, 1.7, GOLD), (9.0, 4.6, 1.3, MINT),
                         (10.6, 5.6, 0.8, RGBColor(0xE8, 0xD9, 0xEE))]:
    rect(s, cx, cy, d, d, col, MSO_SHAPE.OVAL)
text(s, 9.4, 1.1, 2.6, 2.6, [[("Doce", {"size": 26, "bold": True})], [("Encanto", {"size": 26, "bold": True})],
                             [("confeitaria & café", {"size": 11})]], 26, CARD, True, C, MSO_ANCHOR.MIDDLE, 1.0)
rect(s, 0.8, 1.4, 0.12, 3.1, CORAL)
text(s, 1.15, 1.25, 7.2, 0.4, EMPRESA.upper(), 14, GOLD, True)
text(s, 1.15, 1.7, 7.3, 2.0, "Projeto de Modernização Tecnológica", 40, CARD, True, spacing=1.0)
text(s, 1.15, 3.5, 7.2, 1.0, "Como organização digital, controle financeiro e atendimento online podem "
     "transformar uma confeitaria de bairro em uma empresa mais produtiva, segura e lucrativa.",
     14, RGBColor(0xE8, 0xD9, 0xEE))
rect(s, 1.15, 5.15, 6.6, 1.25, RGBColor(0x4E, 0x2A, 0x63), MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 1.4, 5.25, 6.2, 1.1, [
    [("Aluno(a): ", {"bold": True, "color": GOLD}), ("[Seu nome completo]", {})],
    [("Turma: ", {"bold": True, "color": GOLD}), ("[Sua turma]", {})],
    [("Data: ", {"bold": True, "color": GOLD}), ("novembro de 2026", {})],
], 14, CARD, spacing=1.25)
notes(s, "Capa: substitua [Seu nome completo] e [Sua turma] pelos seus dados antes de entregar.")

# =====================================================================
# SLIDE 2 – Apresentação da empresa
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 2, "Apresentação da empresa", "Quem é a Doce Encanto e como ela funciona hoje")
text(s, 0.45, 1.35, 12.4, 0.75, [[
    ("A Doce Encanto Confeitaria ", {"bold": True, "color": PLUM}),
    ("é uma empresa familiar fundada em 2012 que produz e vende bolos, doces e salgados. Cresceu rápido "
     "por causa da qualidade dos produtos, mas continua sendo administrada com caderno, calculadora e papel — "
     "o que hoje limita o seu crescimento.", {})]], 13)

items = [
    ("Ramo de atividade", "Alimentação — confeitaria artesanal e café (comércio e produção própria).", "R"),
    ("Produtos e serviços", "Bolos de festa sob encomenda, doces finos, salgados, tortas, café da manhã, "
     "coffee break para empresas e entregas por aplicativo.", "P"),
    ("Público-alvo", "Famílias do bairro (25 a 55 anos), noivas e organizadores de festas, empresas da "
     "região que contratam coffee break e cantinas escolares.", "A"),
    ("Funcionários", "12 pessoas: 1 proprietária, 1 gerente, 4 confeiteiros, 3 atendentes, "
     "1 auxiliar de cozinha, 1 entregador e 1 auxiliar administrativo.", "F"),
]
for i, (t, b, ic) in enumerate(items):
    card(s, 0.45 + (i % 2) * 3.55, 2.25 + (i // 2) * 2.25, 3.4, 2.1, t, b, [CORAL, MINT, GOLD, PLUM][i],
         12.5, 11, icon=ic)

rect(s, 7.65, 2.25, 5.25, 4.6, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
text(s, 7.9, 2.35, 4.9, 0.4, "Funcionamento da empresa", 13, PLUM, True)
table(s, 7.9, 2.85, 4.8, [
    ["Aspecto", "Como funciona hoje"],
    ["Horário", "Seg. a sáb., 7h às 20h; dom., 7h às 13h"],
    ["Produção", "Cozinha própria, 2 turnos (5h–13h e 12h–20h)"],
    ["Vendas", "Balcão, telefone, encomendas e 1 app de delivery"],
    ["Movimento", "≈ 250 clientes/dia e 90 encomendas/mês"],
    ["Faturamento", "≈ R$ 78 mil por mês"],
    ["Gestão", "Caderno, planilha de papel e 1 computador antigo"],
], [1.25, 3.55], row_h=0.5, size=10)
notes(s, "Mostrar que a empresa é bem-sucedida em produto, mas fraca em gestão — é aí que entra a tecnologia.")

# =====================================================================
# SLIDE 3 – Diagnóstico dos problemas
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 3, "Diagnóstico dos problemas", "6 problemas de organização e uso da tecnologia — com causas e consequências")
table(s, 0.45, 1.35, 12.43, [
    ["#", "Problema identificado", "Causa", "Consequência para a empresa"],
    ["1", "Controle financeiro manual",
     "Entradas e saídas anotadas em caderno; caixa fechado na calculadora.",
     "A dona não sabe se tem lucro; erros de caixa de ≈ R$ 600/mês; contas pagas com atraso e juros."],
    ["2", "Encomendas registradas em papel",
     "Pedidos anotados em folhas soltas ou recebidos por telefone sem registro.",
     "≈ 8 encomendas por mês esquecidas, atrasadas ou com sabor errado; reembolsos e clientes perdidos."],
    ["3", "Arquivos sem backup e senha única",
     "Tudo fica em 1 computador antigo; todos usam a mesma senha \"doce123\".",
     "Risco de perder notas fiscais e dados de clientes; qualquer pessoa acessa informações sigilosas."],
    ["4", "Estoque sem controle",
     "Ninguém registra entradas, saídas e validade dos ingredientes.",
     "Desperdício de ≈ R$ 3.200/mês em produtos vencidos e compras de urgência mais caras."],
    ["5", "Atendimento lento e só por telefone",
     "Um único telefone fixo; sem WhatsApp Business, catálogo ou redes sociais ativas.",
     "Linha ocupada, respostas em até 4 horas, clientes desistem e compram no concorrente."],
    ["6", "Escalas e tarefas combinadas \"de boca\"",
     "Escala de trabalho em papel no mural; tarefas não ficam registradas.",
     "Faltas sem substituto, horas extras desnecessárias e conflitos entre a equipe."],
], [0.4, 2.65, 4.4, 4.98], row_h=0.74, size=10.5,
    aligns=[C, L, L, L])
notes(s, "Cada problema tem causa (por que acontece) e consequência (o que causa à empresa).")

# =====================================================================
# SLIDE 4 – Prioridades
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 4, "Análise e definição de prioridades", "Problemas classificados por gravidade (impacto financeiro × risco × urgência)")
table(s, 0.45, 1.35, 8.1, [
    ["Ordem", "Problema", "Gravidade", "Prejuízo estimado / mês"],
    ["1º", "Controle financeiro manual", "Crítica", "R$ 600 em erros + juros e decisões no escuro"],
    ["2º", "Arquivos sem backup / senha única", "Crítica", "Risco de perder tudo + multas da LGPD"],
    ["3º", "Encomendas registradas em papel", "Alta", "R$ 2.400 (8 pedidos × R$ 300)"],
    ["4º", "Estoque sem controle", "Alta", "R$ 3.200 em desperdício"],
    ["5º", "Atendimento lento e só por telefone", "Média", "≈ R$ 1.800 em vendas perdidas"],
    ["6º", "Escalas combinadas \"de boca\"", "Média", "≈ R$ 700 em horas extras"],
], [0.75, 2.9, 1.15, 3.3], row_h=0.5, size=10.5, aligns=[C, L, C, L],
    fills={(1, 2): RGBColor(0xF6, 0xD5, 0xD6), (2, 2): RGBColor(0xF6, 0xD5, 0xD6),
           (3, 2): RGBColor(0xFB, 0xE3, 0xD8), (4, 2): RGBColor(0xFB, 0xE3, 0xD8),
           (5, 2): RGBColor(0xFB, 0xF0, 0xD6), (6, 2): RGBColor(0xFB, 0xF0, 0xD6)})
rect(s, 0.45, 5.05, 8.1, 1.8, PLUM, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0.7, 5.15, 7.7, 1.65, [
    [("Prejuízo total estimado: ≈ R$ 8.700 por mês", {"size": 16, "bold": True, "color": GOLD})],
    [("Esse valor equivale a mais de 11% do faturamento. Ou seja: a empresa vende bem, mas perde dinheiro "
      "por desorganização. Resolver os problemas 1 a 4 já recupera cerca de 80% desse prejuízo.",
      {"size": 11.5, "color": CARD, "space": 4})],
], 12, CARD)

card(s, 8.8, 1.35, 4.08, 2.65, "Por que resolver primeiro o 1º e o 2º?",
     "O financeiro é a base de todas as decisões: sem saber quanto entra e sai, a empresa não sabe nem "
     "se pode investir. A segurança vem junto porque uma pane no único computador apagaria notas, "
     "contatos e histórico — um dano que não tem conserto.", RED, 12.5, 10.5)
card(s, 8.8, 4.15, 4.08, 2.7, "Critérios usados",
     [[("Impacto financeiro: ", {"bold": True}), ("quanto o problema custa por mês.", {})],
      [("Risco: ", {"bold": True}), ("se o dano pode ser irreversível ou gerar multa.", {})],
      [("Urgência: ", {"bold": True}), ("se piora a cada dia que passa.", {})],
      [("Os problemas 5 e 6 são importantes, mas o prejuízo é menor e eles ficam mais fáceis de resolver "
        "depois que a base estiver organizada.", {"space": 4})]],
     GOLD, 12.5, 10.5)
notes(s, "Explicar o critério de gravidade e o valor total do prejuízo mensal.")

# =====================================================================
# SLIDE 5 – Ferramentas digitais
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 5, "Ferramentas digitais", "5 ferramentas escolhidas para resolver os problemas do diagnóstico")
tools = [
    ("Sistema de PDV e gestão", "Ex.: Bling, Conta Azul",
     "Registra cada venda no caixa, dá baixa automática no estoque e gera relatórios financeiros.",
     "Ao vender 1 bolo de cenoura, o sistema desconta farinha, ovos e açúcar do estoque e lança R$ 45 "
     "no faturamento do dia.", "Resolve: 1 e 4", CORAL, "$"),
    ("WhatsApp Business", "Atendimento",
     "Catálogo de produtos, mensagens automáticas, etiquetas por status e link direto para pedidos.",
     "Cliente pergunta o preço do bolo de festa às 22h e recebe na hora o catálogo e o link do formulário.",
     "Resolve: 5", MINT, "W"),
    ("Google Workspace", "Docs, Planilhas, Forms, Drive, Agenda",
     "Escritório na nuvem: documentos, planilhas e agenda compartilhados com backup automático.",
     "A gerente lança as contas a pagar na Planilha e a dona acompanha pelo celular em tempo real.",
     "Resolve: 1, 2, 3 e 6", PLUM, "G"),
    ("Trello", "Gestão de tarefas",
     "Quadro visual com colunas (A fazer, Em produção, Pronto, Entregue) e responsáveis.",
     "Cada encomenda vira um cartão com sabor, data e foto; os confeiteiros movem o cartão na produção.",
     "Resolve: 2 e 6", GOLD, "T"),
    ("Canva + Instagram", "Marketing digital",
     "Criação de artes profissionais e divulgação dos produtos para atrair novos clientes.",
     "Toda sexta, a atendente cria um post do \"bolo da semana\" com fotos próprias e agenda a publicação.",
     "Resolve: 5", RED, "C"),
]
w = 2.42
for i, (t, sub, f, ex, res, col, ic) in enumerate(tools):
    x = 0.45 + i * (w + 0.08)
    rect(s, x, 1.35, w, 5.5, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
    rect(s, x, 1.35, w, 0.95, col, MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, 2.0, w, 0.3, col)
    text(s, x + 0.12, 1.42, w - 0.24, 0.55, t, 13, CARD, True, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
    text(s, x + 0.12, 1.95, w - 0.24, 0.3, sub, 9.5, CARD)
    text(s, x + 0.12, 2.42, w - 0.24, 1.7, [[("Função", {"bold": True, "color": PLUM, "size": 10.5})],
                                           [(f, {"space": 2})]], 10.5)
    text(s, x + 0.12, 4.08, w - 0.24, 2.0, [[("Exemplo prático", {"bold": True, "color": PLUM, "size": 10.5})],
                                           [(ex, {"space": 2, "italic": True})]], 10.5)
    rect(s, x + 0.12, 6.33, w - 0.24, 0.38, BG, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + 0.12, 6.33, w - 0.24, 0.38, res, 10, col, True, C, MSO_ANCHOR.MIDDLE)
notes(s, "Os números em \"Resolve\" correspondem aos problemas do slide 3.")

# =====================================================================
# SLIDE 6 – Google Workspace
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 6, "Google Workspace na empresa", "Como cada ferramenta entra na rotina dos funcionários")
apps = [
    ("Documentos", "D", RGBColor(0x42, 0x85, 0xF4),
     "Receitas padronizadas (ficha técnica), manual de atendimento, contratos de encomendas de "
     "casamento e atas de reunião.",
     "Confeiteiro novo abre a ficha \"Bolo Red Velvet\" no tablet e segue as mesmas medidas de sempre."),
    ("Planilhas", "P", RGBColor(0x0F, 0x9D, 0x58),
     "Fluxo de caixa, contas a pagar, controle de estoque com validade e cálculo de custo de cada produto "
     "usando fórmulas (SOMA, SE, PROCV).",
     "A célula fica vermelha quando o leite condensado chega a menos de 10 latas: hora de comprar."),
    ("Formulários", "F", RGBColor(0x67, 0x3A, 0xB7),
     "Formulário de encomendas (sabor, tamanho, data, endereço) e pesquisa de satisfação após "
     "cada entrega.",
     "As respostas caem direto numa Planilha: nenhuma encomenda se perde e a nota média é calculada sozinha."),
    ("Drive", "Dr", RGBColor(0xF4, 0xB4, 0x00),
     "Guarda todos os arquivos da empresa na nuvem, organizados por setor, com permissão por pessoa "
     "e histórico de versões.",
     "Se o computador quebrar, basta entrar em outro aparelho: nada é perdido."),
    ("Agenda", "A", RGBColor(0xDB, 0x44, 0x37),
     "Agenda compartilhada de entregas, escala de turnos, vencimento de boletos e manutenção dos "
     "equipamentos.",
     "O entregador recebe no celular o lembrete \"Sáb. 14h – Bolo 3 kg – Rua das Flores, 120\"."),
]
for i, (n, ic, col, uso, ex) in enumerate(apps):
    y = 1.32 + i * 1.1
    rect(s, 0.45, y, 12.43, 1.0, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
    rect(s, 0.6, y + 0.2, 0.6, 0.6, col, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, 0.6, y + 0.2, 0.6, 0.6, ic, 14, CARD, True, C, MSO_ANCHOR.MIDDLE)
    text(s, 1.3, y + 0.05, 1.75, 0.9, "Google\n" + n, 12.5, PLUM, True, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
    text(s, 3.1, y + 0.05, 5.0, 0.9, [[("Como usar: ", {"bold": True, "color": col}), (uso, {})]], 10.5,
         anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 8.25, y + 0.12, 4.5, 0.76, BG, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, 8.35, y + 0.12, 4.3, 0.76, [[("Na prática: ", {"bold": True, "color": PLUM}),
                                        (ex, {"italic": True})]], 10, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Destacar a integração: Formulário alimenta a Planilha, que fica salva no Drive, e as entregas vão para a Agenda.")

# =====================================================================
# SLIDE 7 – Organização e segurança
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 7, "Organização e segurança das informações", "Estrutura de pastas no Drive + 6 medidas de proteção dos dados")
rect(s, 0.45, 1.35, 5.0, 5.5, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
text(s, 0.65, 1.45, 4.7, 0.4, "Proposta de organização (Drive compartilhado)", 12.5, PLUM, True)
folders = [
    (0, "Doce Encanto – Empresa", PLUM, ""),
    (1, "01_Financeiro", CORAL, "Gerente e dona"),
    (2, "2026 › 11_Novembro", MUTED, ""),
    (1, "02_Encomendas", CORAL, "Atendimento"),
    (1, "03_Estoque_e_Compras", CORAL, "Cozinha"),
    (1, "04_Receitas_Fichas_Tecnicas", CORAL, "Cozinha"),
    (1, "05_RH_Funcionarios", RED, "Só a dona"),
    (1, "06_Clientes (dados pessoais)", RED, "Restrito"),
    (1, "07_Marketing_Fotos", CORAL, "Todos"),
    (1, "08_Documentos_Legais", CORAL, "Dona"),
]
for i, (lvl, nome, col, acesso) in enumerate(folders):
    y = 1.95 + i * 0.36
    x = 0.7 + lvl * 0.35
    rect(s, x, y + 0.07, 0.24, 0.2, col, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + 0.32, y, 3.2, 0.34, nome, 10.5, INK, lvl == 0, anchor=MSO_ANCHOR.MIDDLE)
    if acesso:
        text(s, 3.85, y, 1.45, 0.34, acesso, 9, col if col == RED else MUTED, align=R, anchor=MSO_ANCHOR.MIDDLE)
rect(s, 0.65, 5.65, 4.6, 1.05, BG, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0.75, 5.68, 4.4, 1.0, [[("Padrão de nomes: ", {"bold": True, "color": PLUM}),
                               ("AAAA-MM-DD_tipo_descrição", {"bold": True})],
                              [("Ex.: 2026-11-08_encomenda_casamento-Ana.pdf — os arquivos ficam em ordem de data "
                                "e qualquer pessoa encontra o que precisa.", {"size": 9.5, "space": 2})]], 10.5)

measures = [
    ("1. Senhas individuais + verificação em 2 etapas",
     "Cada funcionário tem login próprio e senha forte; o código no celular impede acesso mesmo se a senha vazar."),
    ("2. Backup automático 3-2-1",
     "3 cópias, em 2 meios diferentes (Drive + HD externo), 1 fora da loja. Um computador quebrado não apaga nada."),
    ("3. Permissões por função",
     "Cada um só vê as pastas do seu setor. Dados de clientes e RH ficam restritos (exigência da LGPD)."),
    ("4. Antivírus e atualizações",
     "Sistemas e antivírus atualizados automaticamente fecham brechas usadas por vírus e sequestro de dados."),
    ("5. Rede Wi-Fi separada",
     "Uma rede para clientes e outra, com senha diferente, para o caixa e os computadores da empresa."),
    ("6. Treinamento contra golpes",
     "A equipe aprende a reconhecer links falsos e golpes do Pix; telas bloqueiam após 5 min sem uso."),
]
for i, (t, b) in enumerate(measures):
    x = 5.65 + (i % 2) * 3.65
    y = 1.35 + (i // 2) * 1.86
    card(s, x, y, 3.55, 1.74, t, b, [MINT, PLUM, RED, GOLD, CORAL, MINT][i], 11.5, 10)
notes(s, "Pastas numeradas mantêm a ordem; as pastas em vermelho têm acesso restrito.")

# =====================================================================
# SLIDE 8 – Ética
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 8, "Ética e responsabilidade digital", "3 situações que podem acontecer na Doce Encanto — e como agir corretamente")
cases = [
    ("Direitos autorais e uso de imagens", "Lei 9.610/98 (Direitos Autorais)", CORAL,
     "Para economizar, uma atendente copia do Instagram fotos de bolos de outra confeitaria e usa no "
     "cardápio da Doce Encanto, como se fossem produtos da casa.",
     "A outra empresa pode processar e pedir indenização; o post é removido; clientes recebem um produto "
     "diferente da foto e a loja perde credibilidade.",
     "Usar fotos próprias (celular + Canva), bancos de imagens gratuitos com licença livre ou pedir "
     "autorização por escrito e dar o crédito ao autor."),
    ("Privacidade de dados", "Lei 13.709/18 (LGPD)", PLUM,
     "Um funcionário exporta a lista de clientes (nome, telefone e endereço) e envia para um amigo que vende "
     "produtos de festa, sem que os clientes tenham autorizado.",
     "Vazamento de dados pessoais: multa de até 2% do faturamento, processo dos clientes e "
     "demissão por justa causa do funcionário.",
     "Coletar só os dados necessários, pedir consentimento no formulário, restringir o acesso à pasta "
     "Clientes e nunca repassar dados a terceiros."),
    ("Divulgação de informações falsas", "Código Penal – calúnia e difamação", RED,
     "Uma mensagem diz que \"a confeitaria concorrente usa ovos estragados\". Um funcionário compartilha "
     "a mensagem no grupo da loja e nos stories da Doce Encanto sem verificar.",
     "A notícia era falsa: a empresa pode responder por difamação, pagar indenização e ter a imagem manchada "
     "por espalhar boatos.",
     "Verificar a fonte antes de compartilhar, não publicar nada sobre concorrentes e criar uma regra: "
     "só a gerente publica nas redes da empresa."),
]
for i, (t, lei, col, sit, cons, cor) in enumerate(cases):
    x = 0.45 + i * 4.19
    rect(s, x, 1.35, 4.05, 5.5, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
    rect(s, x, 1.35, 4.05, 0.12, col)
    text(s, x + 0.2, 1.55, 3.7, 0.4, t, 13.5, PLUM, True)
    text(s, x + 0.2, 1.93, 3.7, 0.3, lei, 9.5, col, True)
    blocks = [("Situação", sit, BG), ("Consequências", cons, RGBColor(0xFB, 0xE9, 0xE6)),
              ("Forma correta de agir", cor, RGBColor(0xE3, 0xF3, 0xF1))]
    yy = 2.35
    for j, (lab, body, fill) in enumerate(blocks):
        rect(s, x + 0.15, yy, 3.75, 1.43, fill, MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, x + 0.25, yy + 0.05, 3.55, 1.35,
             [[(lab, {"bold": True, "color": [PLUM, RED, MINT][j], "size": 10.5})], [(body, {"space": 2})]], 10)
        yy += 1.5
notes(s, "Relacionar cada situação a uma lei e mostrar a atitude correta.")

# =====================================================================
# SLIDE 9 – Análise financeira
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 9, "Análise financeira", "Resultado de um mês típico (outubro/2026) — antes da modernização")
receitas = [("Vendas no balcão", 38000), ("Encomendas de bolos e festas", 22000),
            ("Coffee break para empresas", 8500), ("Delivery por aplicativo", 5000),
            ("Salgados para cantinas escolares", 4500)]
despesas = [("Salários e encargos (12 func.)", 32400), ("Ingredientes e embalagens", 26000),
            ("Aluguel do imóvel", 6500), ("Energia, água e gás", 4800),
            ("Desperdício de produtos vencidos", 3200), ("Taxas de app e maquininha", 2900),
            ("Manutenção e despesas diversas", 2700)]
TR, TD = sum(v for _, v in receitas), sum(v for _, v in despesas)


def brl(v):
    s_ = f"{abs(v):,.0f}".replace(",", ".")
    return ("- R$ " if v < 0 else "R$ ") + s_


rows = [["Fontes de receita", "Valor"]] + [[n, brl(v)] for n, v in receitas] + [["TOTAL DE RECEITAS", brl(TR)]]
table(s, 0.45, 1.35, 4.25, rows, [2.95, 1.3], row_h=0.42, size=10.5, bold_rows=(len(rows) - 1,),
      header_fill=MINT, aligns=[L, R], fills={(len(rows) - 1, None): RGBColor(0xE3, 0xF3, 0xF1)})
rows = [["Despesas mensais", "Valor"]] + [[n, brl(v)] for n, v in despesas] + [["TOTAL DE DESPESAS", brl(TD)]]
table(s, 4.9, 1.35, 4.35, rows, [3.05, 1.3], row_h=0.42, size=10.5, bold_rows=(len(rows) - 1,),
      header_fill=RED, aligns=[L, R], fills={(len(rows) - 1, None): RGBColor(0xFB, 0xE9, 0xE6)})

rect(s, 9.45, 1.35, 3.43, 2.35, RED, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 9.6, 1.45, 3.1, 2.2, [
    [("RESULTADO DO MÊS", {"size": 11, "bold": True})],
    [("Receitas − Despesas", {"size": 10, "space": 2})],
    [(f"{brl(TR)} − {brl(TD)}", {"size": 11, "space": 2})],
    [(brl(TR - TD), {"size": 30, "bold": True, "space": 4})],
    [("PREJUÍZO", {"size": 14, "bold": True})],
], 12, CARD, align=C)
card(s, 9.45, 3.85, 3.43, 3.0, "O que a tabela mostra", "", GOLD, 12, 10.5)
text(s, 9.67, 4.45, 3.1, 2.35, [
    [("A empresa fatura R$ 78 mil, mas gasta R$ 78,5 mil: trabalha muito e ainda perde R$ 500 por mês.", {})],
    [("Só o desperdício (R$ 3.200) já transformaria o prejuízo em lucro. Esse é o principal sinal de que "
      "falta controle — não de que faltam clientes.", {"space": 5})],
], 10.5)
notes(s, f"Total de receitas: {brl(TR)}. Total de despesas: {brl(TD)}. Resultado: {brl(TR - TD)} (prejuízo).")

# =====================================================================
# SLIDE 10 – Investimento
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 10, "Investimento em tecnologia", "Orçamento de implantação — 8 itens com valores e justificativas")
inv = [
    ("Sistema de PDV e gestão (plano anual)", 2400, "Controla caixa, estoque e financeiro (problemas 1 e 4)."),
    ("2 computadores novos", 6000, "Substituem a máquina antiga: 1 no caixa e 1 no escritório."),
    ("Tablet para balcão e cozinha", 1500, "Atendentes lançam pedidos e confeiteiros consultam as fichas."),
    ("Google Workspace (6 contas × 12 meses)", 2520, "E-mail, Drive, Planilhas e Agenda com backup e segurança."),
    ("Impressora térmica + leitor de código", 1200, "Cupom e etiqueta com validade de cada produto."),
    ("HD externo 2 TB + nobreak", 1100, "Segunda cópia do backup e proteção contra queda de energia."),
    ("Roteador com 2 redes Wi-Fi", 600, "Separa a rede de clientes da rede da empresa."),
    ("Treinamento da equipe (16 horas)", 1800, "Garante que os funcionários usem as ferramentas de verdade."),
]
TI = sum(v for _, v, _ in inv)
rows = [["#", "Investimento", "Valor (R$)", "Justificativa"]]
rows += [[str(i + 1), n, brl(v), j] for i, (n, v, j) in enumerate(inv)]
rows += [["", "INVESTIMENTO TOTAL", brl(TI), "Pago à vista com reserva + parcelamento em 6×"]]
table(s, 0.45, 1.35, 9.0, rows, [0.4, 3.1, 1.15, 4.35], row_h=0.5, size=10, bold_rows=(len(rows) - 1,),
      aligns=[C, L, R, L], fills={(len(rows) - 1, None): RGBColor(0xF1, 0xE6, 0xF5)})

# gráfico de pizza do orçamento por categoria
cd = CategoryChartData()
cd.categories = ["Equipamentos", "Softwares", "Segurança e rede", "Treinamento"]
cd.add_series("Investimento", (6000 + 1500 + 1200, 2400 + 2520, 1100 + 600, 1800))
gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(9.6), Inches(1.3), Inches(3.3), Inches(3.4), cd)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(9)
ch.legend.font.name = FONT
ch.has_title = True
ch.chart_title.text_frame.text = "Divisão do investimento"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
ch.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = PLUM
plot = ch.plots[0]
plot.has_data_labels = True
plot.data_labels.number_format = '0%'
plot.data_labels.number_format_is_linked = False
plot.data_labels.show_percentage = True
plot.data_labels.show_value = False
plot.data_labels.font.size = Pt(9)
plot.data_labels.font.bold = True
plot.data_labels.font.color.rgb = CARD
for idx, col in enumerate([PLUM, CORAL, MINT, GOLD]):
    pt = plot.series[0].points[idx]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = col

rect(s, 9.6, 4.85, 3.28, 2.0, PLUM, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 9.75, 4.92, 3.0, 1.9, [
    [("Por que essas escolhas?", {"bold": True, "color": GOLD, "size": 12})],
    [("Cada item ataca um problema do diagnóstico. Foram escolhidas soluções em nuvem e de baixo custo "
      "mensal (≈ R$ 410/mês), sem precisar de técnico fixo.", {"space": 3})],
], 10, CARD)
notes(s, f"Investimento total: {brl(TI)}.")

# =====================================================================
# SLIDE 11 – Antes e depois
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 11, "Comparação antes e depois", "Melhorias esperadas 6 meses após a implantação (projeção com dados fictícios)")
table(s, 0.45, 1.35, 6.9, [
    ["Indicador", "Antes", "Depois", "Melhoria"],
    ["Encomendas perdidas ou erradas/mês", "8", "1", "−88%"],
    ["Desperdício de ingredientes/mês", "R$ 3.200", "R$ 900", "−72%"],
    ["Tempo para fechar o caixa", "60 min", "10 min", "−83%"],
    ["Tempo de resposta ao cliente", "até 4 h", "15 min", "−94%"],
    ["Tempo para registrar uma encomenda", "10 min", "3 min", "−70%"],
    ["Nota de satisfação (Google Forms)", "6,8", "9,0", "+32%"],
    ["Faturamento mensal", "R$ 78.000", "R$ 82.000", "+5%"],
    ["Resultado mensal", "- R$ 500", "R$ 6.500", "+R$ 7.000"],
], [2.95, 1.25, 1.25, 1.45], row_h=0.5, size=10.5, aligns=[L, C, C, C], bold_rows=(8,),
    fills={(i, 3): RGBColor(0xE3, 0xF3, 0xF1) for i in range(1, 9)})
text(s, 0.45, 5.95, 6.9, 0.9, "Como foram estimados: comparando os registros do mês de outubro (caderno e caixa) "
     "com as metas para abril/2027. Os novos custos mensais de software já estão incluídos no resultado \"depois\".",
     9.5, MUTED)

rect(s, 7.55, 1.35, 5.33, 5.5, CARD, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
cd = CategoryChartData()
cd.categories = ["Out/26\n(antes)", "Dez/26", "Jan/27", "Fev/27", "Mar/27", "Abr/27"]
cd.add_series("Receitas (R$ mil)", (78.0, 79.0, 80.0, 80.5, 81.5, 82.0))
cd.add_series("Despesas (R$ mil)", (78.5, 78.0, 77.0, 76.2, 75.8, 75.5))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7.65), Inches(1.45), Inches(5.13), Inches(3.55), cd)
ch = gf.chart
ch.has_title = True
ch.chart_title.text_frame.text = "Receitas × despesas mensais (R$ mil)"
r0 = ch.chart_title.text_frame.paragraphs[0].runs[0].font
r0.size, r0.bold, r0.color.rgb, r0.name = Pt(11), True, PLUM, FONT
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(9)
va = ch.value_axis
va.minimum_scale, va.maximum_scale = 70, 84
va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = LINE
va.tick_labels.font.size = Pt(9)
va.format.line.fill.background()
ch.category_axis.tick_labels.font.size = Pt(9)
ch.category_axis.format.line.color.rgb = LINE
pl = ch.plots[0]
pl.gap_width = 60
pl.overlap = -10
for ser, col in zip(pl.series, [MINT, CORAL]):
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = col

table(s, 7.75, 5.1, 4.93, [
    ["Lucro/prejuízo", "Out", "Dez", "Jan", "Fev", "Mar", "Abr"],
    ["R$ mil", "−0,5", "1,0", "3,0", "4,3", "5,7", "6,5"],
], [1.33, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6], row_h=0.36, size=9.5, aligns=[L] + [C] * 6,
    fills={(1, 1): RGBColor(0xFB, 0xE9, 0xE6)})
text(s, 7.75, 5.9, 4.93, 0.9, "A partir de dezembro as despesas ficam abaixo das receitas: o prejuízo vira lucro "
     "porque a empresa desperdiça menos e perde menos vendas. Novembro é o mês da implantação.", 10, INK)
notes(s, "Gráfico com dados fictícios de projeção. Lucro = receitas − despesas.")

# =====================================================================
# SLIDE 12 – Conclusão
# =====================================================================
s = prs.slides.add_slide(BLANK)
header(s, 12, "Conclusão e avaliação", "O que mais vai fazer diferença, os riscos, como medir e se vale a pena")
card(s, 0.45, 1.35, 6.1, 2.65, "Soluções de maior impacto",
     [[("1. Sistema de PDV e gestão: ", {"bold": True}),
       ("integra caixa, estoque e financeiro; sozinho reduz o desperdício em ≈ R$ 2.300/mês e mostra o lucro real.", {})],
      [("2. Google Forms + Planilhas + Agenda: ", {"bold": True, "space": 4}),
       ("acabam com as encomendas perdidas e organizam as entregas.", {})],
      [("3. Backup e senhas individuais: ", {"bold": True, "space": 4}),
       ("protegem o que a empresa tem de mais valioso — seus dados e seus clientes.", {})]],
     MINT, 13, 10.5)
card(s, 6.78, 1.35, 6.1, 2.65, "Dificuldades que podem surgir e como contornar",
     [[("Resistência da equipe: ", {"bold": True}), ("treinamento prático e um \"padrinho digital\" por turno.", {})],
      [("Curva de aprendizado: ", {"bold": True, "space": 3}), ("implantar em etapas, 1 ferramenta por semana.", {})],
      [("Internet instável: ", {"bold": True, "space": 3}), ("plano de internet de reserva no celular (4G).", {})],
      [("Custo inicial: ", {"bold": True, "space": 3}), ("parcelar equipamentos e começar pelo item mais urgente.", {})]],
     GOLD, 13, 10.5)
card(s, 0.45, 4.15, 6.1, 2.7, "Como medir os resultados",
     [[("Reunião mensal com um painel na Planilha Google, comparando com o mês de outubro:", {})],
      [("• Lucro mensal e desperdício (R$)", {"space": 3})],
      [("• Nº de encomendas erradas ou atrasadas", {})],
      [("• Tempo de resposta no WhatsApp", {})],
      [("• Nota de satisfação no Google Forms (meta ≥ 9)", {})],
      [("• Tempo de fechamento do caixa", {})]],
     PLUM, 13, 10.5)

rect(s, 6.78, 4.15, 6.1, 2.7, PLUM, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 7.0, 4.25, 5.7, 0.4, "O investimento é viável? Sim.", 14, GOLD, True)
kpis = [(brl(TI), "investimento total"), ("+ R$ 7.000", "ganho por mês"), ("≈ 2,5 meses", "para recuperar"),
        ("≈ 390%", "retorno em 12 meses")]
for i, (v, l) in enumerate(kpis):
    x = 7.0 + i * 1.44
    rect(s, x, 4.72, 1.36, 1.0, RGBColor(0x4E, 0x2A, 0x63), MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, 4.76, 1.36, 0.5, v, 12.5, CARD, True, C, MSO_ANCHOR.MIDDLE)
    text(s, x, 5.22, 1.36, 0.45, l, 8.5, RGBColor(0xE8, 0xD9, 0xEE), align=C)
text(s, 7.0, 5.82, 5.7, 1.0, "O resultado passa de − R$ 500 para + R$ 6.500 por mês. Em 12 meses o ganho "
     "(≈ R$ 84 mil) é cerca de 5 vezes o valor investido. A modernização custa menos do que o prejuízo "
     "que a desorganização já causa hoje.", 10.5, CARD)
notes(s, "Retorno em 12 meses = (7.000 × 12 − 17.120) ÷ 17.120 ≈ 390%. Payback = 17.120 ÷ 7.000 ≈ 2,4 meses.")

assert len(prs.slides) == 12
out = __file__.replace("gerar_apresentacao.py", "modernizacao_doce_encanto.pptx")
prs.save(out)
print("ok", out, "receitas", TR, "despesas", TD, "investimento", TI)
