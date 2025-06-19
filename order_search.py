from abc import ABC, abstractmethod
import pandas as pd
import re


class DataLoader(ABC):
    """Абстрактный класс для загрузки данных"""

    @abstractmethod
    def load_data(self, filename):
        pass


class ExcelDataLoader(DataLoader):
    """Класс для загрузки данных из Excel-файлов"""

    def load_data(self, filename):
        try:
            return pd.read_excel(filename)
        except FileNotFoundError:
            raise ValueError(f"Файл '{filename}' не найден.")
        except Exception as e:
            raise RuntimeError(f"Ошибка при загрузке данных: {e}")


class OrderProcessor:
    """Класс для обработки заказов"""

    def __init__(self, data_loader: DataLoader):
        self.data_loader = data_loader

    def process_order(self, order_number):
        df = self.data_loader.load_data('РАСКРОЙ 2025.xlsx')
        filtered_rows = df[df['№ Заказа'].astype(str) == str(order_number)]

        if filtered_rows.empty:
            return f"Заказ №{order_number} не найден."

        first_row = filtered_rows.iloc[0]
        info_extractor = InfoExtractor(first_row)
        extracted_info = info_extractor.extract()
        return extracted_info.format_output()


class InfoExtractor:
    """Класс для извлечения и обработки информации из записей"""

    def __init__(self, row):
        self.row = row

    def extract(self):
        info = {
            'store_application_number': self._extract_store_application(),
            'client': self._extract_client(),
            'full_name': self._extract_full_name(),
            'item_name': self._extract_item_name(),
            'dimensions': self._extract_dimensions(),
            'carcase': self._extract_carcase(),
            'extra_component': self._extract_extra_component(),
            'facade': self._extract_facade(),
            'weight': self._extract_weight(),
        }
        return OrderInfo(**info)

    def _extract_store_application(self):
        return self.row.get('№ магазина / заявка', '')

    def _extract_client(self):
        return self.row.get('Клиент', '')

    def _extract_full_name(self):
        return self.row.get('Наименование', '')

    def _extract_item_name(self):
        match = re.match(r'(.*?)(\d+)[xхХХ*×]', self.row.get('Наименование', ''))
        return match.group(1).strip() if match else ''

    def _extract_dimensions(self):
        dimensions_match = re.search(r'(\d+)\s*[xхХХ*×]\s*(\d+)\s*[xхХХ*×]\s*(\d+)', self.row.get('Наименование', ''))
        return tuple(map(int, dimensions_match.groups())) if dimensions_match else ()

    def _extract_carcase(self):
        raw_carcase = self.row.get('Корпус', '').split('/')
        words = {re.match(r'\D+', p.strip()).group().strip() for p in raw_carcase}
        return '/'.join(words)

    def _extract_extra_component(self):
        component = self.row.get('Профиль /            Доп. Элементы', '')
        return None if component in ['-', ''] else component

    def _extract_facade(self):
        facade = self.row.get('Фасад', '')
        return None if facade in ['-', ''] else facade

    def _extract_weight(self):
        weight = self.row.get('ВЕС, КГ', '')
        return float(weight) if isinstance(weight, (float, int)) else None


class OrderInfo:
    """Класс для представления извлечённой информации о заказе"""

    def __init__(self, **kwargs):
        self.store_application_number = kwargs.get('store_application_number', '')
        self.client = kwargs.get('client', '')
        self.full_name = kwargs.get('full_name', '')
        self.item_name = kwargs.get('item_name', '')
        self.dimensions = kwargs.get('dimensions', ())
        self.carcase = kwargs.get('carcase', '')
        self.extra_component = kwargs.get('extra_component', None)
        self.facade = kwargs.get('facade', None)
        self.weight = kwargs.get('weight', None)

    def format_output(self):
        output = [
            f"✅ Номер заказа: {self.store_application_number}",
            f"✅ Магазин / Заявка: {self.client}",
            f"✅ Полное наименование: {self.full_name}",
            f"✅ Наименование изделия: {self.item_name}"
        ]
        if len(self.dimensions) >= 3:
            output.extend([
                f"✅ Ширина: {self.dimensions[0]} мм",
                f"✅ Высота: {self.dimensions[1]} мм",
                f"✅ Глубина: {self.dimensions[2]} мм"
            ])
        output.append(f"✅ Корпус: {self.carcase}")
        output.append(f"✅ Дополнительный компонент: {self.extra_component or 'нет данных'}")
        output.append(f"✅ Фасад: {self.facade or 'нет данных'}")
        if self.weight is not None:
            output.append(f"✅ Вес: {int(self.weight)} кг")

        return "\n".join(output)


def main():
    loader = ExcelDataLoader()
    processor = OrderProcessor(loader)
    while True:
        order_number = input("🔍 Введите номер заказа (или введите 'q' для выхода): ")
        if order_number.lower() == 'q':
            break
        try:
            result = processor.process_order(order_number)
            print(result)
        except Exception as e:
            print(f"❌ Произошла ошибка: {e}")


if __name__ == "__main__":
    main()