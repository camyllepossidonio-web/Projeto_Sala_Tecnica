import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 1. Criar o arquivo e a aba principal
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Medição de Obra"

# Garantir exibição das linhas de grade no Excel
ws.views.sheetView[0].showGridLines = True

# 2. Definição da Paleta de Cores e Estilos
NAVY_HEADER = "1E3A8A"  # Azul Marinho para o cabeçalho
ZEBRA_FILL = "F8FAFC"  # Cinza suave para alternar linhas
CARD_BG = "F1F5F9"  # Fundo para destaques de grupo

font_title = Font(name="Segoe UI", size=16, bold=True, color="1E293B")
font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="64748B")
font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
font_data = Font(name="Segoe UI", size=10, color="0F172A")
font_bold = Font(name="Segoe UI", size=10, bold=True, color="0F172A")

fill_header = PatternFill(
    start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid"
)
fill_zebra = PatternFill(
    start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid"
)
fill_card = PatternFill(
    start_color=CARD_BG, end_color=CARD_BG, fill_type="solid"
)

thin_side = Side(border_style="thin", color="CBD5E1")
thick_bottom = Side(border_style="medium", color="1E3A8A")
double_bottom = Side(border_style="double", color="1E3A8A")

border_cell = Border(
    left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
)
border_header = Border(
    left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom
)
border_total = Border(top=thin_side, bottom=double_bottom)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")

# 3. Cabeçalho de Identificação da Obra
ws.merge_cells("A1:K1")
ws["A1"] = "RELATÓRIO DE MEDIÇÃO DE OBRA"
ws["A1"].font = font_title

ws.merge_cells("A2:K2")
ws["A2"] = "Acompanhamento Físico-Financeiro dos Serviços de Engenharia"
ws["A2"].font = font_subtitle

info_labels = [
    ("Obra:", "Construção de Edifício Residencial - Bloco A"),
    ("Contratante:", "Construtora Horizonte S/A"),
    ("Medição Nº:", "03 / 2026"),
    ("Período:", "01/10/2026 a 31/10/2026"),
]

for idx, (label, val) in enumerate(info_labels):
  row = 4 + idx // 2
  col_label = 1 if idx % 2 == 0 else 7
  ws.cell(row=row, column=col_label, value=label).font = font_bold
  ws.cell(row=row, column=col_label + 1, value=val).font = font_data

# 4. Tabela Principal de Medição
headers = [
    "Item",
    "Descrição do Serviço",
    "Unid.",
    "Qtd. Prevista",
    "Preço Unit. (R$)",
    "Valor Total (R$)",
    "Qtd. Medida",
    "Valor Medido (R$)",
    "Qtd. Acumulada",
    "Valor Acumulado (R$)",
    "% Executado",
]

header_row = 7
for col_idx, text in enumerate(headers, 1):
  cell = ws.cell(row=header_row, column=col_idx, value=text)
  cell.font = font_header
  cell.fill = fill_header
  cell.alignment = align_center
  cell.border = border_header

# 5. Dados dos Serviços e Inserção das Fórmulas do Excel
servicos = [
    ("1.0", "SERVIÇOS PRELIMINARES", "", "", "", "", ""),
    ("1.1", "Canteiro de obras e mobilização", "M2", 150.0, 120.00, 150.0, 0.0),
    ("1.2", "Limpeza e gabarito do terreno", "M2", 800.0, 25.00, 800.0, 0.0),
    ("2.0", "INFRAESTRUTURA", "", "", "", "", ""),
    ("2.1", "Escavação mecânica para fundações", "M3", 350.0, 45.00, 350.0, 0.0),
    (
        "2.2",
        "Estacas escavadas em concreto armado",
        "M",
        420.0,
        180.00,
        300.0,
        120.0,
    ),
    ("2.3", "Blocos de coroamento e baldrame", "M3", 85.0, 950.00, 40.0, 45.0),
    ("3.0", "SUPRAESTRUTURA", "", "", "", "", ""),
    (
        "3.1",
        "Concreto armado fck=30MPa para pilares",
        "M3",
        110.0,
        850.00,
        0.0,
        55.0,
    ),
    ("3.2", "Formas de madeira plastificada", "M2", 1450.0, 65.00, 0.0, 600.0),
    ("3.3", "Aço CA-50 cortado e dobrado", "KG", 12500.0, 12.50, 0.0, 4500.0),
    ("4.0", "ALVENARIA E FECHAMENTOS", "", "", "", "", ""),
    (
        "4.1",
        "Alvenaria de vedação bloco cerâmico 14x19x29",
        "M2",
        2200.0,
        58.00,
        0.0,
        350.0,
    ),
]

start_data_row = 8
current_row = start_data_row

for item in servicos:
  code, desc, unid, qtd_prev, preco, qtd_ant, qtd_atual = item
  is_group = unid == ""

  ws.cell(row=current_row, column=1, value=code).alignment = align_center
  ws.cell(row=current_row, column=2, value=desc).alignment = align_left
  ws.cell(row=current_row, column=3, value=unid).alignment = align_center

  if is_group:
    # Agrupador/Subtotal
    for c in range(1, 12):
      cell = ws.cell(row=current_row, column=c)
      cell.font = font_bold
      cell.fill = fill_card
  else:
    # Item com Fórmulas do Excel
    ws.cell(row=current_row, column=4, value=qtd_prev).number_format = (
        "#,##0.00"
    )
    ws.cell(row=current_row, column=5, value=preco).number_format = (
        "R$ #,##0.00"
    )

    # Valor Total Previsto = Qtd Prevista * Preço Unit.
    ws.cell(
        row=current_row, column=6, value=f"=D{current_row}*E{current_row}"
    ).number_format = "R$ #,##0.00"

    # Qtd Medida Atual
    ws.cell(row=current_row, column=7, value=qtd_atual).number_format = (
        "#,##0.00"
    )

    # Valor Medido = Qtd Medida * Preço Unit.
    ws.cell(
        row=current_row, column=8, value=f"=G{current_row}*E{current_row}"
    ).number_format = "R$ #,##0.00"

    # Qtd Acumulada = Qtd Anterior + Qtd Atual
    ws.cell(
        row=current_row, column=9, value=f"={qtd_ant}+{qtd_atual}"
    ).number_format = "#,##0.00"

    # Valor Acumulado = Qtd Acumulada * Preço Unit.
    ws.cell(
        row=current_row, column=10, value=f"=I{current_row}*E{current_row}"
    ).number_format = "R$ #,##0.00"

    # % Executado = Qtd Acumulada / Qtd Prevista
    ws.cell(
        row=current_row,
        column=11,
        value=f"=IF(D{current_row}>0, I{current_row}/D{current_row}, 0)",
    ).number_format = "0.0%"

    for c in range(1, 12):
      cell = ws.cell(row=current_row, column=c)
      cell.font = font_data
      cell.border = border_cell
      if current_row % 2 == 1:
        cell.fill = fill_zebra

  current_row += 1

# 6. Linha de Total Geral
tot_row = current_row
ws.cell(row=tot_row, column=2, value="TOTAL GERAL").font = font_bold
ws.cell(row=tot_row, column=2).alignment = align_right

ws.cell(
    row=tot_row, column=6, value=f"=SUM(F{start_data_row}:F{tot_row-1})"
).number_format = "R$ #,##0.00"
ws.cell(
    row=tot_row, column=8, value=f"=SUM(H{start_data_row}:H{tot_row-1})"
).number_format = "R$ #,##0.00"
ws.cell(
    row=tot_row, column=10, value=f"=SUM(J{start_data_row}:J{tot_row-1})"
).number_format = "R$ #,##0.00"
ws.cell(
    row=tot_row, column=11, value=f"=J{tot_row}/F{tot_row}"
).number_format = "0.0%"

for c in range(1, 12):
  cell = ws.cell(row=tot_row, column=c)
  cell.font = font_bold
  cell.border = border_total

# 7. Formatação Condicional (Destacar itens 100% concluídos em verde)
rule_green = CellIsRule(
    operator="greaterThanOrEqual",
    formula=["1"],
    stopIfTrue=True,
    font=Font(color="065F46", bold=True),
    fill=PatternFill(start_color="D1FAE5", fill_type="solid"),
)
ws.conditional_formatting.add(f"K{start_data_row}:K{tot_row-1}", rule_green)

# 8. Ajuste Automático de Largura das Colunas
for col in ws.columns:
  max_len = 0
  col_letter = get_column_letter(col[0].column)
  for cell in col:
    if cell.row < 7:
      continue
    if len(str(cell.value or "")) > max_len:
      max_len = len(str(cell.value or ""))
  ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

ws.column_dimensions["B"].width = 45  # Descrição com maior espaço

# Salvar Arquivo
wb.save("Planilha_de_Medicao_Obra.xlsx")
