"""Gera exemplo/FICHA_EXEMPLO.xlsx: uma ficha de faca ficticia com o mesmo layout
da aba usada em producao (area de impressao A1:L71, assinaturas nas linhas 66 e 68).

Uso: pip install openpyxl && python exemplo/gerar_exemplo.py
Depois abra e salve uma vez no Excel: arquivo salvo so pelo openpyxl faz o Excel travar ao copiar a area como imagem.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

AQUI = os.path.dirname(os.path.abspath(__file__))
CINZA, CINZA2, AMARELO = PatternFill("solid", fgColor="D9D9D9"), PatternFill("solid", fgColor="BFBFBF"), PatternFill("solid", fgColor="FFFF99")
FINO = Side(style="thin")
CAIXA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
CENTRO, DIREITA = Alignment(horizontal="center", vertical="center", wrap_text=True), Alignment(horizontal="right", vertical="center")

wb = Workbook()
ws = wb.active
ws.title = "FICHA"
for col, w in zip("ABCDEFGHIJKL", [15.6, 10.6, 13, 13, 13, 13, 13, 13, 13, 8.6, 10.6, 13]):
    ws.column_dimensions[col].width = w
for r in range(1, 72):
    ws.row_dimensions[r].height = 5.25
for r, h in {1: 33.6, 2: 99.6, 4: 26.4, 8: 20.25, 9: 20.4, 11: 20.25, 12: 55.8, 65: 18, 66: 40.8,
             67: 14.25, 68: 45, 69: 14.25, 70: 14.25, 71: 14.25}.items():
    ws.row_dimensions[r].height = h


def cel(ref, valor, merge=None, font=None, fill=None, align=CENTRO, borda=False):
    ws[ref] = valor
    if merge:
        ws.merge_cells(merge)
    alvo = ws[merge] if merge else ((ws[ref],),)
    for linha in alvo:
        for c in linha:
            if font: c.font = font
            if fill: c.fill = fill
            c.alignment = align
            if borda: c.border = CAIXA


# cabecalho (empresa e produto ficticios)
cel("A1", "Cartonagem Aurora", "A1:C1", Font(bold=True, size=16, color="1F3B57"))
cel("D1", "FICHA TÉCNICA DE PRODUÇÃO", "D1:L1", Font(size=20))
cel("A2", "FACA", "A2:C2", Font(bold=True, size=36))
cel("D2", "FX2041A", "D2:L2", Font(bold=True, size=72), AMARELO, borda=True)
cel("A4", "Cliente Exemplo S.A.", "A4:C4", Font(size=16), CINZA, Alignment(vertical="center"))
cel("D4", "Caixa cartucho 250 g, 4x0 + verniz UV", "D4:L4", Font(size=16), CINZA)
cel("A8", "Roteiro de fabricação", "A8:B9", Font(size=14), align=Alignment(vertical="center", wrap_text=True))
for col, etapa, x in zip("CDEFGHIJKL", ["IMP V", "IMP F", "TERC", "ACOP", "HOT", "CV", "COL", "DES", "REV", "INK"],
                         ["", "X", "", "", "", "X", "X", "", "", ""]):
    cel(f"{col}8", etapa, font=Font(bold=True), fill=CINZA, borda=True)
    cel(f"{col}9", x, font=Font(bold=True, size=14, color="C00000"), fill=AMARELO if x else None, borda=True)
cel("A11", "OBSERVAÇÕES DE PRODUÇÃO", "A11:L11", Font(bold=True, size=12), CINZA2)
cel("A12", "Ficha fictícia para demonstração. Alterações somente com aprovação da engenharia.", "A12:L12", borda=True)

SECOES = [
    (14, "IMPRESSÃO VERSO", [(16, "CLASSIFICAÇÃO", "-", "TEMPO DE SET UP", "-"),
                             (18, "GRUPO DE MÁQUINAS", "-", "CADÊNCIA (fls / h)", "-"),
                             (20, "MÁQUINAS HABILITADAS", "-", "TABELA DE PERDAS", "-")]),
    (22, "IMPRESSÃO FRENTE", [(24, "CLASSIFICAÇÃO", "Normal", "TEMPO DE SET UP", "00:45"),
                              (26, "GRUPO DE MÁQUINAS", "IMP-A2", "CADÊNCIA (fls / h)", "8.500"),
                              (28, "CORES", "4X0", "TIPO DE VERNIZ", "UV"),
                              (30, "FORMATO IMPRESSÃO", "1020x720", "SETUP POR COR", "00:12"),
                              (32, "MÁQUINAS HABILITADAS", "Offset 6 cores", "TABELA DE PERDAS", "IMP-UV-02")]),
    (34, "CORTE E VINCO", [(36, "CLASSIFICAÇÃO", "Normal", "TEMPO DE SET UP", "03:00"),
                           (38, "GRUPO DE MÁQUINAS", "CV-B1", "CADÊNCIA (fls / h)", "6.000"),
                           (40, "MÁQUINAS HABILITADAS", "Plana automática", "TABELA DE PERDAS", "CV-UV-01")]),
    (42, "COLAGEM", [(44, "TIPO DE COLAGEM", "Lateral", "TEMPO DE SET UP", "01:00"),
                     (46, "GRUPO DE MÁQUINAS", "COL-C1", "CADÊNCIA (unds / h)", "24.000"),
                     (48, "MÁQUINAS HABILITADAS", "Coladeira linear", "TABELA DE PERDAS", "-")]),
    (51, "HOT STAMPING", [(53, "TEMPO DE SET UP", "-", "CADÊNCIA (unds / h)", "-")]),
    (56, "ATIVIDADES SECUNDÁRIAS (cadência - unidades/h) - (% perdas)",
     [(58, "ACOPLAMENTO", "-", "REVISÃO", "-"), (60, "DESTAQUE (ABERTOS)", "-", "INK JET", "-"),
      (62, "RESERVA DE PERDAS", "-", "IMPOSIÇÃO", "6")]),
]
for linha, titulo, campos in SECOES:
    ws.row_dimensions[linha].height = 19.95
    cel(f"A{linha}", titulo, f"A{linha}:L{linha}", Font(bold=True, size=12), CINZA2, Alignment(vertical="center"))
    for r, l1, v1, l2, v2 in campos:
        ws.row_dimensions[r].height = 30
        cel(f"A{r}", l1, f"A{r}:B{r}", Font(size=9), align=DIREITA)
        cel(f"C{r}", v1, f"C{r}:E{r}", borda=True)
        cel(f"G{r}", l2, f"G{r}:I{r}", Font(size=9), align=DIREITA)
        cel(f"J{r}", v2, f"J{r}:L{r}", borda=True)

# bloco de assinaturas: as celulas das linhas 66 e 68 recebem as imagens
cel("A65", "RESPONSÁVEIS", "A65:L65", Font(bold=True, size=12), CINZA2, Alignment(vertical="center"))
for ref, rng, nome in [("A67", "A67:B67", "Impressão Offset"), ("D67", "D67:E67", "Corte e Vinco"),
                       ("G67", "G67:H67", "Colagem/Destaque"), ("J67", "J67:K67", "Eng. Processos"),
                       ("A69", "A69:B69", "Coordenador de Produção"), ("D69", "D69:E69", "PCP"),
                       ("G69", "G69:H69", "Excelência Operacional")]:
    cel(ref, nome, rng, Font(size=9))
    for c in ws[rng][0]:
        c.border = Border(top=FINO)
for rng in ["A66:B66", "D66:E66", "G66:H66", "J66:K66", "A68:B68", "D68:E68", "G68:H68"]:
    ws.merge_cells(rng)
cel("I69", "=TODAY()", "I69:L69", Font(size=9))
ws["I69"].number_format = "dd/mm/yyyy"

ws.print_area = "A1:L71"
ws.page_setup.orientation = "portrait"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
ws.sheet_view.showGridLines = False
saida = os.path.join(AQUI, "FICHA_EXEMPLO.xlsx")
wb.save(saida)
print("gerado:", saida)
