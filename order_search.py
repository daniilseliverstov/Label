import pandas as pd
import re


def main():
    try:
        # Загружаем данные из Excel-файла
        df = pd.read_excel('РАСКРОЙ 2025.xlsx')

        # Просим пользователя ввести номер заказа
        order_number = input("🔍 Введите номер заказа: ").strip()

        # Фильтруем таблицу по указанному номеру заказа
        row = df[df['№ Заказа'].astype(str) == order_number]

        if not row.empty:
            # Берём первую подходящую запись
            found_row = row.iloc[0]

            # Выбираем нужные поля
            store_application_number = found_row['№ магазина / заявка']
            client = found_row['Клиент']
            full_name = found_row['Наименование']

            # Регулярное выражение для выделения наименования и размеров изделия
            match = re.match(r'(.*?)(?:\s*)(\d+)[xхХХ*×](\d+)[xхХХ*×](\d+)', full_name)

            if match:
                item_name = match.group(1).strip()  # Название изделия
                width = int(match.group(2))  # Ширина
                height = int(match.group(3))  # Высота
                depth = int(match.group(4))  # Глубина
            else:
                item_name = full_name
                width = height = depth = None

            carcase_value = found_row['Корпус']

            # Разбиваем строку на отдельные элементы
            parts = carcase_value.split('/')

            # Используем множество для хранения уникальных значений
            unique_words = set()

            # Обрабатываем каждый элемент отдельно
            for part in parts:
                # Извлекаем начало строки до первых цифр с помощью регулярного выражения
                match = re.match(r'^\D+', part.strip())
                if match:
                    word = match.group().strip()
                    unique_words.add(word)  # Добавляем слово в множество

            # Преобразуем множество обратно в список и объединяем обработанные слова обратно в одну строку
            carcase = '/'.join(unique_words)

            extra_component = found_row['Профиль /            Доп. Элементы']

            if extra_component == '' or extra_component == '-':
                extra_component = None




            # Выводим информацию о заказе
            print("\n✅ Найден заказ:")
            print(f"Номер заказа: {order_number}")
            print(f"Магазин / заявка: {store_application_number}")
            print(f"Клиент: {client}")
            print(f"Полное наименование: {full_name}")
            print(f"Наименование: {item_name}")
            print(f"Ширина: {width or 'не определено'}")
            print(f"Высота: {height or 'не определено'}")
            print(f"Глубина: {depth or 'не определено'}")
            print(f"Корпус: {carcase}")
            print(f'Доп. Элемент: {extra_component}')
        else:
            print("❌ Заказ не найден.")
    except Exception as e:
        print(f"Ошибка обработки: {e}")


if __name__ == "__main__":
    main()
