# Defines
import os
from xml.etree import ElementTree

class StreamXMLParser:
    """
    Класс-парсер XML-файлов

    Объекты класса представляют собой сборку парсера для поиска содержимого в XML-фалах, определяемого тегом,
    который задает пользователь. Файлы для отработки сканируются в каталоге, заданным атрибутом объекта src_path,
    результирующие файлы помещаются в каталог, заданный атрибутом объекта dst_path. Если путь для результирующих файлов
    не задан, отработанные файлы складываются в тот же каталог.

    Атрибуты объекта
    ----------------
        src_path (str)
            Путь до каталога, где лежат файлы на отработку
        find_tag (str)
            Имя тега, который будет искаться в XML-файле
        dst_path (str)
            Путь до каталога, куда будут выгружаться обработанные файлы

    Методы класса
    ----------------
        __init__(self, src_path: str, find_tag: str, dst_path: str = '')
            Конструктор: Инициализация
        parse(self)
            Парсинг xml-файлов в каталоге
    """

    def __init__(self, src_path: str, find_tag: str, dst_path: str = ''):
        """
        Конструктор: Инициализация

        Метод инициализирует класс атрибутами путей и именем искомого тега
        :param str src_path: Путь до каталога, где лежат файлы на отработку
        :param str find_tag: Имя тега, который будет искаться в XML-файле
        :param str dst_path: Путь до каталога, куда будут выгружаться обработанные файлы
        """

        self.src_path = src_path
        self.dst_path = src_path if dst_path == '' else dst_path
        self.find_tag = find_tag

    def parse(self):
        """
        Парсинг xml-файлов в каталоге

        Метод проходится по каталогу, определенного с помощью атрибута src_path, генерирует список XMK-файлов
        и отрабатывает его в цикле, производя поиск содержимого тегов по имени, заданному в find_tag. Результирующие
        файлы, содержащие текст ограниченный искомым тегом, записываются в каталог, определенный атрибутом объекта
        dst_path. Имя файла формируется из имя_тега+имя_исходного_файла+.txt
        """

        files = [f for f in os.listdir(self.src_path) if f.endswith('.xml')]

        for xml_file in files:
            xml_root = ElementTree.parse(f'{self.src_path}\\{xml_file}').getroot()
            for element in xml_root.iter(self.find_tag):
                with open(f'{self.dst_path}\\{self.find_tag}_{xml_file}.txt', 'a', encoding='utf-8') as file:
                    file.write(element.text)

# Code
xsp = StreamXMLParser("C:\\Users\\marie\\Downloads", "special_notes")
xsp.parse()