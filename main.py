import json
from openpyxl import Workbook
from openpyxl.styles import Font

class Seat:
    def __init__(self, number, seat_type, comfort_class, reserved=False):
        self.number = number
        self.seat_type = seat_type  # "нижнее"/"верхнее"
        self.comfort_class = comfort_class  # "купейное"/"плацкартное"
        self.reserved = reserved

    def __eq__(self, other):
        # Места считаются одинаковыми, если совпадает класс и тип места
        return (self.comfort_class == other.comfort_class and 
                self.seat_type == other.seat_type)

    def to_dict(self):
        # Преобразование объекта в словарь для сериализации
        return self.__dict__

    @classmethod
    def from_dict(cls, data):
        # Создание объекта из словаря
        return cls(**data)

class Carriage:
    def __init__(self, number, carriage_type):
        self.number = number
        self.carriage_type = carriage_type  # "купейный"/"плацкартный"
        self.seats = []  # Список мест в вагоне

    def add_seat(self, seat):
        # Добавление места в вагон
        self.seats.append(seat)

    def __eq__(self, other):
        # Вагоны равны, если совпадают номера и типы
        return self.number == other.number and self.carriage_type == other.carriage_type

    def to_dict(self):
        # Сериализация вагона и всех его мест
        return {
            'number': self.number,
            'carriage_type': self.carriage_type,
            'seats': [seat.to_dict() for seat in self.seats]
        }

    @classmethod
    def from_dict(cls, data):
        # Создание вагона из словаря
        carriage = cls(data['number'], data['carriage_type'])
        # Восстановление всех мест вагона
        carriage.seats = [Seat.from_dict(seat_data) for seat_data in data['seats']]
        return carriage

class Locomotive:
    def __init__(self, serial_number, power):
        self.serial_number = serial_number
        self.power = power

    def __eq__(self, other):
        # Тепловозы равны по серийному номеру
        return self.serial_number == other.serial_number

    def to_dict(self):
        return self.__dict__

    @classmethod
    def from_dict(cls, data):
        return cls(**data)

class Train:
    def __init__(self, number, route):
        self.number = number  # Номер поезда
        self.route = route    # Маршрут
        self.locomotive = None  # Локомотив
        self.carriages = []     # Список вагонов

    def set_locomotive(self, locomotive):
        # Установка локомотива для поезда
        self.locomotive = locomotive

    def add_carriage(self, carriage):
        # Добавление вагона в состав (если его еще нет)
        if carriage not in self.carriages:
            self.carriages.append(carriage)

    def remove_carriage(self, carriage):
        # Удаление вагона из состава
        if carriage in self.carriages:
            self.carriages.remove(carriage)

    def __eq__(self, other):
        # Поезда равны по номеру
        return self.number == other.number

    def save_to_file(self, filename="train_data.txt"):
        # Сохранение всего состава в JSON файл
        data = {
            'number': self.number,
            'route': self.route,
            'locomotive': self.locomotive.to_dict() if self.locomotive else None,
            'carriages': [carriage.to_dict() for carriage in self.carriages]
        }
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @classmethod
    def load_from_file(cls, filename="train_data.txt"):
        # Загрузка состава из JSON файла
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Создание объекта поезда
        train = cls(data['number'], data['route'])
        # Восстановление локомотива
        if data['locomotive']:
            train.set_locomotive(Locomotive.from_dict(data['locomotive']))
        # Восстановление всех вагонов
        train.carriages = [Carriage.from_dict(carriage_data) for carriage_data in data['carriages']]
        return train

    def create_excel_report(self, filename="train_report.xlsx"):
        # Создание Excel отчета с информацией о составе
        wb = Workbook()
        ws = wb.active
        ws.title = f"Train {self.number}"

        # Заголовки таблицы
        headers = ['Вагон', 'Тип вагона', 'Место', 'Тип места', 'Класс', 'Статус']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)  # Жирный шрифт для заголовков

        # Заполнение данными
        row = 2
        for carriage in self.carriages:
            for seat in carriage.seats:
                # Запись информации о каждом месте
                ws.cell(row=row, column=1, value=carriage.number)
                ws.cell(row=row, column=2, value=carriage.carriage_type)
                ws.cell(row=row, column=3, value=seat.number)
                ws.cell(row=row, column=4, value=seat.seat_type)
                ws.cell(row=row, column=5, value=seat.comfort_class)
                ws.cell(row=row, column=6, value="Забронировано" if seat.reserved else "Свободно")
                row += 1

        # Автоматическая настройка ширины колонок
        for column in ws.columns:
            max_length = 0
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[column[0].column_letter].width = max_length + 2

        wb.save(filename)
        print(f"Отчет сохранен: {filename}")

# Демонстрация работы программы
if __name__ == "__main__":
    # Создание тестовых данных
    seat1 = Seat(1, "нижнее", "купейное")
    seat2 = Seat(2, "верхнее", "купейное")
    seat3 = Seat(1, "нижнее", "плацкартное")

    # Создание и заполнение вагонов
    carriage1 = Carriage(10, "купейный")
    carriage1.add_seat(seat1)
    carriage1.add_seat(seat2)

    carriage2 = Carriage(11, "плацкартный")
    carriage2.add_seat(seat3)

    # Создание локомотива
    locomotive = Locomotive("ТЭП-70-1234", 4000)

    # Формирование состава поезда
    train = Train("045А", "Москва - Санкт-Петербург")
    train.set_locomotive(locomotive)
    train.add_carriage(carriage1)
    train.add_carriage(carriage2)

    # Тестирование методов сравнения
    print("Сравнение мест (одинаковый класс, разный тип):", seat1 == seat2)
    print("Сравнение мест (разный класс):", seat1 == seat3)

    # Сохранение и загрузка данных
    train.save_to_file()
    loaded_train = Train.load_from_file()

    # Генерация Excel отчета
    loaded_train.create_excel_report()