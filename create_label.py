from openpyxl import Workbook
from openpyxl.styles import Border, Side

# Создаем рабочий лист
wb = Workbook()
ws = wb.active


# Задаем параметры высот и широт ячеек
row_heights = {1: 13.5, 2: 12.75, 3: 12.75, 4: 13.5, 5: 12.75, 6: 12.75, 7: 12.75, 8: 13.5, 9: 12.75,
              10: 12.75, 11: 12.75, 12: 13.5, 13: 12.75, 14: 12.75, 15: 12.75, 16: 13.5, 17: 6.75}

col_widths = {"A":8.36, "B":8.36, "C": 8.36, "D": 8.36, "E": 9.07, "F": 8.36, "G": 9.65, "H": 8.36, "I": 9.22, "J": 8.36,
              "K": 8.36, "L": 12.22, "M": 11.22, "N": 8.36, "O": 12.94, "P": 6.65, "Q": 8.36, "R": 17.65, "S": 4.07}
# Устанавливаем высоту строк
for row, height in row_heights.items():
    ws.row_dimensions[row].height = height
# Устанавливаем ширину столбцов
for col, width in col_widths.items():
    ws.column_dimensions[col].width = width


# Задаем данные для объединения ячеек
merge_ranges = ["A1:E8", "A9:B12", "C9:E12", "A13:E16", "F1:L4", "M1:O4", "P1:R4", "S1:S16", "F5:I8", "G5:L8", "M5:O8",
                "P5:R8", "F9:I12", "J9:L12", "M9:O12", "P9:R12", "F13:G14", "H13:I14", "J13:K14", "F15:G16", "H15:I16",
                "J15:K16", "L13:M16", "N13:N16", "O13:O16", "P13:R16"]

# Объединяем ячейки
for merge_range in merge_ranges:
    ws.merge_cells(merge_range)

# Создаем стиль границы ячейки
thick_border = Border(left=Side(style="thick"),
                      right=Side(style="thick"),
                      top=Side(style="thick"),
                      bottom=Side(style="thick"))

# Применяем стиль границы ко всем ячейкам в диапазоне
for merge_range in merge_ranges:
    # Получаем координаты первой и последней ячейки в диапазоне
    start_cell, end_cell = merge_range.split(':')
    row_start, col_start = start_cell[1:], start_cell[0]
    row_end, col_end = end_cell[1:], end_cell[0]

    # Применяем стиль границы ко всем ячейкам в диапазоне
    for row in range(int(row_start), int(row_end) + 1):
        for col in range(ord(col_start), ord(col_end) + 1):
            cell = ws[f"{chr(col)}{row}"]
            cell.border = thick_border

wb.save("label.xlsx")