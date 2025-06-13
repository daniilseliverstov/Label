from PIL.ImageChops import offset
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.drawing.image import Image
from sizes import row_heights, col_widths
from datetime import datetime, timedelta

# Создаем рабочий лист
wb = Workbook()
ws = wb.active

# Устанавливаем высоту строк
for row, height in row_heights.items():
    ws.row_dimensions[row].height = height
# Устанавливаем ширину столбцов
for col, width in col_widths.items():
    ws.column_dimensions[col].width = width


# Задаем данные для объединения ячеек
merge_ranges = ["A1:E8", "A9:B12", "C9:E12", "A13:E16", "F1:L4", "M1:O4", "P1:R4", "S1:S16", "F5:I8", "J5:L8", "M5:O8",
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

# Загружаем изображения
kodmi = Image("images/Logo.png")
eac = Image("images/EAC.png")
contacts = Image("images/Contacts.png")


# Задаем размеры изображений
kodmi.width = 323.62
kodmi.height = 108.4615384615385
eac.width = 77.214
eac.height = 61.53856
contacts.width = 193.035
contacts.height = 65.38455

# Вставляем изображения
ws.add_image(kodmi, "A2")
ws.add_image(eac, "A9")
ws.add_image(contacts, "C9")

# Вставляем текст в ячейку
ws["A13"] = "ГОСТ 16371-2014"

# Задаем шрифт, высоту текста
font = Font(name="Times New Roman", size=16, bold=True)
ws["A13"].font = font
# Выравниваем текст по центру
alignment = Alignment(horizontal="center", vertical="center")
ws["A13"].alignment = alignment

ws["F5"] = "Наименование упаковки"
font = Font(name="Times New Roman", size=16, bold=True)
ws["F5"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["F5"].alignment = alignment

ws["J5"] = "Цвет"
font = Font(name="Times New Roman", size=20, bold=True)
ws["J5"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["J5"].alignment = alignment

ws["M5"] = "ЗАКАЗЧИК"
font = Font(name="Times New Roman", size=20, bold=True)
ws["M5"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["M5"].alignment = alignment

ws["P1"] = "ВСЕГО УПАКОВОК"
font = Font(name="Times New Roman", size=14, bold=True)
ws["P1"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["P1"].alignment = alignment

ws["P9"] = "№ УПАКОВКИ"
font = Font(name="Times New Roman", size=14, bold=True)
ws["P9"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["P9"].alignment = alignment

ws["F13"] = "ВЫСОТА"
font = Font(name="Times New Roman", size=14, bold=True)
ws["F13"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["F13"].alignment = alignment

ws["H13"] = "ШИРИНА"
font = Font(name="Times New Roman", size=14, bold=True)
ws["H13"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["H13"].alignment = alignment

ws["J13"] = "ГЛУБИНА"
font = Font(name="Times New Roman", size=14, bold=True)
ws["J13"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["J13"].alignment = alignment

ws["L13"] = "ВЕС"
font = Font(name="Times New Roman", size=14, bold=True)
ws["L13"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["L13"].alignment = alignment

ws["O13"] = "КГ"
font = Font(name="Times New Roman", size=14, bold=True)
ws["O13"].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws["O13"].alignment = alignment

# Получаем дату и форматируем
today = datetime.now()
_date = today + timedelta(days=7)
label_date = _date.strftime("%d.%m.%Y")

ws["S1"] = label_date
font = Font(name="Times New Roman", size=14, bold=True)
ws[""].font = font
alignment = Alignment(horizontal="center", vertical="center")
ws[""].alignment = alignment

wb.save("label.xlsx")
