from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.drawing.image import Image
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from datetime import datetime, timedelta
from sizes import row_heights, col_widths


def set_cell(ws, cell, value, font=None, alignment=None):
    ws[cell] = value
    if font:
        ws[cell].font = font
    if alignment:
        ws[cell].alignment = alignment

def apply_border(ws, merge_ranges, border, row_offset=0):
    for merge_range in merge_ranges:
        start_cell, end_cell = merge_range.split(':')
        start_col_letter, start_row = coordinate_from_string(start_cell)
        end_col_letter, end_row = coordinate_from_string(end_cell)
        start_col = column_index_from_string(start_col_letter)
        end_col = column_index_from_string(end_col_letter)

        # Сдвигаем строки на row_offset
        new_start_cell = f"{start_col_letter}{start_row + row_offset}"
        new_end_cell = f"{end_col_letter}{end_row + row_offset}"
        new_merge_range = f"{new_start_cell}:{new_end_cell}"

        ws.merge_cells(new_merge_range)

        for row in range(start_row + row_offset, end_row + row_offset + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = border

def create_label(ws, start_row):
    """
    Создает одну этикетку, начиная со строки start_row
    """

    # Смещение для строк
    row_offset = start_row - 1

    # Размеры строк для этикетки
    for r, h in row_heights.items():
        ws.row_dimensions[r + row_offset].height = h

    # Размеры столбцов (ставим только один раз в начале, т.к. Столбцы не меняются)

    # Список объединений с учетом смещения
    merge_ranges = [
        "A1:E8", "A9:B12", "C9:E12", "A13:E16", "F1:L4", "M1:O4", "P1:R4", "S1:S16",
        "F5:I8", "J5:L8", "M5:O8", "P5:R8", "F9:I12", "J9:L12", "M9:O12", "P9:R12",
        "F13:G14", "H13:I14", "J13:K14", "F15:G16", "H15:I16", "J15:K16", "L13:M16",
        "N13:N16", "O13:O16", "P13:R16"
    ]

    thick_border = Border(
        left=Side(style="thick"),
        right=Side(style="thick"),
        top=Side(style="thick"),
        bottom=Side(style="thick")
    )

    apply_border(ws, merge_ranges, thick_border, row_offset=row_offset)

    # Вставка изображений
    # Для изображений нужно сдвигать координаты по строкам
    # Но openpyxl Image вставляется по ячейке, нужно сдвинуть строку
    images_info = [
        ("images/Logo.png", "A2", 323.62, 108.4615384615385),
        ("images/EAC.png", "A9", 77.214, 61.53856),
        ("images/Contacts.png", "C9", 193.035, 65.38455),
    ]

    for path, cell, width, height in images_info:
        col_letter, row_num = coordinate_from_string(cell)
        new_row = row_num + row_offset
        new_cell = f"{col_letter}{new_row}"
        img = Image(path)
        img.width = width
        img.height = height
        ws.add_image(img, new_cell)

    center_alignment = Alignment(horizontal="center", vertical="center")

    text_cells = [
        ("A13", "ГОСТ 16371-2014", Font(name="Times New Roman", size=16, bold=True)),
        ("F5", "Наименование упаковки", Font(name="Times New Roman", size=16, bold=True)),
        ("J5", "Цвет", Font(name="Times New Roman", size=20, bold=True)),
        ("M5", "ЗАКАЗЧИК", Font(name="Times New Roman", size=20, bold=True)),
        ("P1", "ВСЕГО УПАКОВОК", Font(name="Times New Roman", size=14, bold=True)),
        ("P9", "№ УПАКОВКИ", Font(name="Times New Roman", size=14, bold=True)),
        ("F13", "ВЫСОТА", Font(name="Times New Roman", size=14, bold=True)),
        ("H13", "ШИРИНА", Font(name="Times New Roman", size=14, bold=True)),
        ("J13", "ГЛУБИНА", Font(name="Times New Roman", size=14, bold=True)),
        ("L13", "ВЕС", Font(name="Times New Roman", size=14, bold=True)),
        ("O13", "КГ", Font(name="Times New Roman", size=14, bold=True)),
    ]

    for cell, text, font in text_cells:
        col_letter, row_num = coordinate_from_string(cell)
        new_row = row_num + row_offset
        new_cell = f"{col_letter}{new_row}"
        set_cell(ws, new_cell, text, font=font, alignment=center_alignment)

    # Дата с поворотом текста
    _date = datetime.now() + timedelta(days=7)
    label_date = _date.strftime("%d.%m.%Y")
    s_cell = f"S{1 + row_offset}"
    ws[s_cell] = label_date
    ws[s_cell].font = Font(name="Times New Roman", size=26, bold=True)
    ws[s_cell].alignment = Alignment(horizontal="center", vertical="center", textRotation=90)

# Основной код
def main(label_count):
    wb = Workbook()
    ws = wb.active

    # Установка ширины столбцов один раз
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Создаем нужное количество этикеток
    for i in range(label_count):
        start_row = 1 + i * 17  # каждая этикетка занимает 17 строк
        create_label(ws, start_row)

    # Сохраняем файл
    wb.save("multylabel.xlsx")

if __name__ == "__main__":
    n = int(input("Введите количество этикеток: "))
    main(n)
