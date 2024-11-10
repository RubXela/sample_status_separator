from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QApplication, QTableView, 
                             QLabel, QFrame, QHeaderView, 
                             QWidget, QPushButton, 
                             QVBoxLayout, QHBoxLayout, 
                             QGroupBox, QTableWidget, 
                             QTableWidgetItem, 
                             QFileDialog, QCheckBox, QMenu)
from PyQt5.QtGui import QIcon, QPixmap
import pandas as pd
import os, re
import json
from docx import Document
import sys
from PyQt5.QtCore import Qt
import io
# from classis.copy_buffer import Copy_TableCell, MyFilter
# from classis.filterwidget import FilterTableWidget, PandasModel
from utils.data_extract import DateExtractor
# from utils.model import create_table_model
 
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

custom_img = resource_path('blan.png')

class MyGUI(QtWidgets.QMainWindow):
    def __init__(self):  
        super().__init__()
        
        self.table_view = QTableView() 
        self.initUI()
        
        # Определяем атрибуты для хранения значений переменных
        self.num_str = 0
        self.pchk = 0
        self.bchk = 0
        self.rp_mp_pp = 0
        self.muns_count = {}
        self.data = None 
        self.sort_order = Qt.AscendingOrder
        self.sort_orders = {3: Qt.AscendingOrder, 8: Qt.AscendingOrder}
        
    # Функция для обновления видимости строк в таблице             
    def filter_rows(self):
        all_checked = all(checkbox.isChecked() for checkbox in [self.chk_pchk_filter, self.chk_bchk_filter, self.chk_rp_mp_pp_filter]) # Проверяем, все ли чекбоксы отмечены
        all_unchecked = all(not checkbox.isChecked() for checkbox in [self.chk_pchk_filter, self.chk_bchk_filter, self.chk_rp_mp_pp_filter]) # Проверяем, все ли чекбоксы не отмечены
        
        for row in range(self.table_widget1.rowCount()):
            item = self.table_widget1.item(row, 8)
            hide_row = False

            if item is not None:
                if (item.text() == 'ПЧК' and not self.chk_pchk_filter.isChecked()) or (item.text() == 'БЧК' and not self.chk_bchk_filter.isChecked()) or (item.text() == 'РПМППП' and not self.chk_rp_mp_pp_filter.isChecked()):
                    hide_row = True

            if all_checked or all_unchecked: # все чекбоксы отмечены или не отмечены
                hide_row = False

            self.table_widget1.setRowHidden(row, hide_row)

    def initUI(self):   
        self.tbl = QTableWidget()
        # Основной вертикальный layout
        layout = QVBoxLayout()
        # Группируем первые два виджета в QHBoxLayout
        hbox1 = QHBoxLayout()
        # Первый виджет "Работа с файлами"
        group_box1 = QGroupBox("Работа с файлами")
        widget1 = QWidget()
        widget1_layout = QVBoxLayout()
        self.file_base_name = QLabel('Имя файла:', self)
        widget1_layout.addWidget(self.file_base_name)
        
        self.label_start_date = QLabel('Начало продажи:')
        widget1_layout.addWidget(self.label_start_date)
        
        self.label_end_date = QLabel('Окончание продажи:')
        widget1_layout.addWidget(self.label_end_date)
        
        button_load = QPushButton('Загрузить файл СЗ', self)
        button_load.setFixedWidth(150)
        button_load.clicked.connect(self.load_file1)
        widget1_layout.addWidget(button_load)

        button_clear = QPushButton('Очистить', self)
        button_clear.setFixedWidth(150)
        button_clear.clicked.connect(self.clear_data)
        widget1_layout.addWidget(button_clear)
        
        """ button_restore = QPushButton('Восстановить', self)
        button_restore.setFixedWidth(150)
        button_restore.clicked.connect(self.restore_data)
        widget1_layout.addWidget(button_restore)"""
        
        widget1.setLayout(widget1_layout)
        group_box1.setLayout(widget1_layout)
        
        # Добавляем кнопки сортировки
        self.sort_button_4 = QPushButton('Сортировать Срок', self)
        self.sort_button_4.clicked.connect(lambda: self.toggle_sort_order(3))
        self.sort_button_4.setFixedWidth(150) 
        widget1_layout.addWidget(self.sort_button_4)

        widget1.setLayout(widget1_layout)

        group_box1.setLayout(widget1_layout)
        
       # Второй виджет "Информация о СЗ"
        group_box2 = QGroupBox("Информация о СЗ")
        widget2 = QWidget()
        widget2_layout = QVBoxLayout()
        
              
        self.label_sum_template = QLabel('Общее количество шаблонов:')
        self.label_pchk = QLabel('Количество значений ПЧК:')
        self.label_bchk = QLabel('Количество значений БЧК:')
        self.label_rp_mp_pp = QLabel('Количество значений РП, МП, ПП:')
        
        # Добавляем метки для отображения информации
        widget2_layout.addWidget(self.label_sum_template)
        widget2_layout.addWidget(self.label_pchk)
        widget2_layout.addWidget(self.label_bchk)
        widget2_layout.addWidget(self.label_rp_mp_pp)
        
        # Добавляем чекбоксы для фильтров
        filter_chkbox_layout = QHBoxLayout()
        self.chk_pchk_filter = QCheckBox('Фильтр ПЧК')
        self.chk_pchk_filter.setChecked(False)
        self.chk_pchk_filter.stateChanged.connect(self.filter_rows)
        self.chk_bchk_filter = QCheckBox('Фильтр БЧК')
        self.chk_bchk_filter.setChecked(False)
        self.chk_bchk_filter.stateChanged.connect(self.filter_rows)
        self.chk_rp_mp_pp_filter = QCheckBox('Фильтр РП,МП,ПП')
        self.chk_rp_mp_pp_filter.setChecked(False)
        self.chk_rp_mp_pp_filter.stateChanged.connect(self.filter_rows)
        
        filter_chkbox_layout.addWidget(self.chk_pchk_filter)        
        filter_chkbox_layout.addWidget(self.chk_bchk_filter)
        filter_chkbox_layout.addWidget(self.chk_rp_mp_pp_filter)
        widget2_layout.addLayout(filter_chkbox_layout)
        
        # widget2.setLayout(widget2_layout)
        group_box2.setLayout(widget2_layout)
        
        hbox1.addWidget(group_box1)
        hbox1.addWidget(group_box2)
       
        # Добавляем группировку для чекбоксов во втором виджете
        layout.addLayout(hbox1)
                
      # Создаем третий виджет с информацией о СЗ
        accordion_group1 = QGroupBox('Обработка СЗ. Для переноса в буфер обмена нажмите Чек-бокс')
        accordion_layout1 = QVBoxLayout()
        self.table_widget1 = QTableWidget()
        accordion_layout1.addWidget(self.table_widget1)
        accordion_group1.setLayout(accordion_layout1)

        # Добавляем третий вертикальный виджет
        layout.addWidget(accordion_group1)
        
        # Добавляем изображение и иконку окна
        pixmap = QPixmap('blan.png')  # Путь к изображению
        icon = QIcon(pixmap)
        self.setWindowIcon(icon)
        
        # Устанавливаем central widget в окно
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Устанавливаем размеры окна
        self.setGeometry(100, 100, 1100, 600)
        self.setWindowTitle(f'Распределение шаблонов по статусам')
        
    def clear_data(self):
        # Удаление таблицы и очистка данных
        table_widget1 = self.findChild(QtWidgets.QTableWidget)
        if table_widget1:
            table_widget1.clearContents()

    def load_from_json(self, filename):
        with open(filename, 'r') as file:
            data = json.load(file)
        for row, row_data in enumerate(data):
            for col, value in enumerate(row_data.values()):
                item = QTableWidgetItem(value)
                self.tbl.setItem(row, col, item)
        return(data)        
    
    def on_restore_click(self):
        file_name = 'data.json'  
        self.load_from_json(file_name)
    
    def get_table_data(self):
        table_data = []
        
        header_labels = [self.tbl.horizontalHeaderItem(i).text() for i in range(self.tbl.columnCount())]
        print('Header labels:', header_labels)  # Отладочный вывод заголовков столбцов
        
        for row in range(self.tbl.rowCount()):
            row_data = {}
            for column, header in enumerate(header_labels):
                item = self.tbl.item(row, column)
                if item is not None:
                    cell_text = item.text()
                    print(f'Row: {row}, Column: {column}, Header: {header}, Cell Text: {cell_text}')  
                    # Отладочный вывод информации о текущей ячейке
                    row_data[header] = cell_text
                else:
                    row_data[header] = ''
            table_data.append(row_data)
        
        print('Table data:', table_data)  
        return table_data
 
    # сохранение данных в JSON файл
    
    """ def update_table(self, data):
            if not isinstance(data, list):
                print("Ошибка: Переданные данные не являются списком")
                return

            # Очистка таблицы перед обновлением 
            self.table_widget1.setRowCount(0)

            for row_num, row_data in enumerate(data):
                if not isinstance(row_data, dict):
                    print(f"Ошибка: Данные в строке {row_num} не являются словарем")
                    continue
                
                print(f"Обновление строки {row_num} с данными: {row_data}")

                if row_num >= self.table_widget1.rowCount():
                    self.table_widget1.insertRow(row_num)  # Добавление новой строки, если необходимо

                for col_num, key in enumerate(row_data):
                    item = QTableWidgetItem(str(row_data[key]))
                    self.table_widget1.setItem(row_num, col_num, item)  # Установка элемента в таблицу

            return


        def save_table_to_json(self, filename, table_data):
            try:
                with open(filename, 'w') as f:
                    json.dump(table_data, f, indent=4)
                    print('файл открыт', table_data)
            except Exception as e:
                print(f"Ошибка при сохранении данных таблицы в формате JSON: {e}")
                                
        # восстановление данных
        def restore_data(self):
            data = self.load_from_json('data.json')  
            self.update_table(data)"""
    
# Метод для обработки события щелчка по чекбоксу
    def on_checkbox_click(self, checkbox):
        table_data = []
        # Определение индекса строки на основе элемента чекбокса
        index = self.table_widget1.indexAt(checkbox.pos())
        if index.isValid():
            row = index.row()
            
            ""# Обновление фона строки и прав доступа к редактированию 
            if checkbox.isChecked():
                for col in range(self.table_widget1.columnCount()):
                    item = self.table_widget1.item(row, col)
                    item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                    item.setBackground(QtGui.QColor(200, 200, 200))  # Изменить цвет фона
            else:
                for col in range(self.table_widget1.columnCount()):
                    item = self.table_widget1.item(row, col)
                    item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
                    item.setBackground(QtGui.QColor(255, 255, 255))  # Восстановить цвет фона
            # Копирование в буфер обмена
            self.copy_to_clipboard(row)  # Копирование названия в буфер
            # Сохранение данных в JSON файл с именем "data.json"
            # Получение данных из таблицы
            # table_data = self.get_table_data()

            # Обновление таблицы с полученными данными
            # self.update_table(table_data)

            # Сохранение обновленных данных таблицы в JSON файл
            # self.save_table_to_json('data.json', table_data)
                 
    def load_file1(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Открыть файл', '', 'Документы Word (*.docx);;Все файлы (*)')
        # print('Выбранный файл:', file_name)
        file_base_name = os.path.basename(file_name)
        # print('Имя файла:', file_base_name)
        gui.updateWindowTitle(file_base_name)
        # Путь к файлу документа
        docx_file = file_name

        # Создаем экземпляр класса DateExtractor
        date_extractor = DateExtractor(docx_file)

        # Вызываем метод extract_dates для извлечения нужных дат
        date_extractor.extract_dates()

        # Получаем извлеченные даты
        start_date = date_extractor.start_date
        end_date = date_extractor.end_date
        # print('Начало продажи:', start_date)
        # print('Окончание продажи:', end_date)
        gui.update_dates(start_date, end_date)
        
        if file_name:
            doc = Document(file_name)
            all_data = []

            # Обработка данных из всех таблиц
            for table in doc.tables:
                data = []
                for row in table.rows:
                   # data.append('', *[cell.text for cell in row.cells])
                   data.append([''] + [cell.text for cell in row.cells])
      # Убираем строки, содержащие "Название в учетной системе"
                data = [row for row in data if "Название в учетной системе" not in row]
                
                all_data.extend(data)  # Объединяем данные из всех таблиц

            header = ['Внесен', 'Название в учетной системе', 'Тип', 
                      'Срок действия в месяцах', 'Визиты',  
                      'Время посещения', 'Гостевые визиты', 
                      'Заморозка', 'Статус', 'Подарочная ФД', 
                      'Подарочная ФД', 'inBody', 
                      'Иные условия(входят в стоимость карты)', 
                      'Стоимость карты', 'Стоимость месяца/визита', 
                      'Стоимость месяца в договоре']
            
            # Создаем датафрейм с правильным количеством столбцов
            self.df = pd.DataFrame(all_data, columns=header)
            self.table_widget1.clear()
            self.table_widget1.setRowCount(len(self.df.index))
            self.table_widget1.setColumnCount(len(self.df.columns))
            # print(f'Датафрейм:{self.df}')
                            
            self.df['Срок действия в месяцах'] = self.df['Срок действия в месяцах'].apply(lambda x: re.sub(r'\W+', '', str(x)))
            self.df['Статус'] = self.df['Статус'].apply(lambda x: re.sub(r'\W+', '', str(x)))
            
            sum_template = self.df['Название в учетной системе'].count()
                       
            # Подсчет уникальных значений 'Срок действия в месяцах'
            muns_count = self.df['Срок действия в месяцах'].value_counts().to_dict()
            # print("Уникальные значения 'Срок действия в месяцах' и их количество:")
            # print(muns_count)
            
            # Подсчет количества значений по статусу
            pchk = self.df['Статус'].value_counts().get('ПЧК', 0)
            bchk = self.df['Статус'].value_counts().get('БЧК', 0)
            rp_mp_pp = self.df['Статус'].value_counts().get('РПМППП', 0)
            
            gui.update_widget_labels(sum_template, pchk, bchk, rp_mp_pp)

            """print('Общее количество шаблонов', sum_template)
            print('Количество значений ПЧК:', pchk)
            print('Количество значений ПБЧК:', bchk)
            print('Количество значений РП,МП,ПП:', rp_mp_pp)
            print('Срок действия в месяцах:')"""
            for value, count in muns_count.items():
                # print(f"Значение: {value}, Количество: {count}")
               
                self.update_widget_labels(sum_template, pchk, bchk, rp_mp_pp)   
                       
            # Проход по всем строкам и столбцам данных в DataFrame
            for i in range(len(self.df.index)):
                checkbox = QCheckBox()
                checkbox.index = i # присвоил индекс строки
                checkbox.setChecked(False)  # Установка начального состояния "выключено"
                # Связывание события нажатия на чекбокс с методом обработки
                checkbox.clicked.connect(lambda state, item=checkbox: self.on_checkbox_click(item))
                
                # Установка чекбокса в ячейку таблицы (строка, первый столбец)
                self.table_widget1.setCellWidget(i, 0, checkbox)
                                                 
                for j in range(len (self.df.columns)):
                    # Создание элемента таблицы с данными из DataFrame
                    item = QTableWidgetItem(str(self.df.iloc[i, j]))
                    # Установка элемента таблицы по координатам (строка, столбец)
                    self.table_widget1.setItem(i, j, item)
                    # self.table_widget1.setEditTriggers(QTableWidget.NoEditTriggers)  # только для чтения
                    
            # Устанавливаем ширину столбцов таблицы по содержимому
            # Установка ширины первого столбца таблицы
            self.table_widget1.setColumnWidth(0, 80)

            # Установка ширины остальных столбцов по содержимому
            for column in range(1, self.table_widget1.columnCount()):
                self.table_widget1.resizeColumnToContents(column)
                
                # Показать таблицу
                self.table_widget1.show()
                # Устанавливаем заголовки столбцов таблицы
                self.table_widget1.setHorizontalHeaderLabels(header) 
      
        self.tableWidget = QTableWidget()

        # сортировка данных по столбцу
        def sort_column(column_idx):
            self.tableWidget.sortByColumn(column_idx, Qt.AscendingOrder)

        # Пример привязки заголовков столбцов к функции сортировки
        for col in range(self.tableWidget.columnCount()):
            item = QTableWidgetItem(self.tableWidget.horizontalHeaderItem(col).text())
            item.setTextAlignment(Qt.AlignCenter)
            self.tableWidget.setHorizontalHeaderItem(col, item)

            header = self.tableWidget.horizontalHeader()
            header.sectionClicked.connect(
                lambda state, col=col: sort_column(col)
            )

        layout = QVBoxLayout()
        layout.addWidget(self.tableWidget)
        self.setLayout(layout)
            
        # выбор фильтров
        self.chk_pchk_filter.stateChanged.connect(self.handle_pchk_filter_changed)
        self.chk_bchk_filter.stateChanged.connect(self.handle_bchk_filter_changed)
        self.chk_rp_mp_pp_filter.stateChanged.connect(self.handle_rp_mp_pp_filter_changed)

    def handle_pchk_filter_changed(self):
        self.filter_rows()
        # print('pchk вкл')

    def handle_bchk_filter_changed(self):
        self.filter_rows()
        # print('bchk вкл')

    def handle_rp_mp_pp_filter_changed(self):
        self.filter_rows()
        # print('rp_mp_pp вкл')
            
    def updateWindowTitle(self, file_base_name):
        self.file_base_name.setText(f'Файл: {file_base_name}')

    def enable_editing(self, idx):
        pass
    def copy_to_clipboard(self, i):  # копируем в буфер
        selected_value = str(self.df.iloc[i, 1])  # выбираем вторю ячейку
        clipboard = QApplication.clipboard()
        clipboard.setText(selected_value)
    def update_dates(self, start_date, end_date):
        self.label_start_date.setText(f'Начало продажи: {start_date}')
        self.label_end_date.setText(f'Окончание продажи: {end_date}')
    
        
    def update_widget_labels(self, sum_template, pchk, bchk, rp_mp_pp):
        
        self.label_sum_template.setText(f'Общее количество шаблонов: {sum_template}')
        self.label_pchk.setText(f'Количество значений ПЧК: {pchk}')
        self.label_bchk.setText(f'Количество значений БЧК: {bchk}')
        self.label_rp_mp_pp.setText(f'Количество значений РП, МП, ПП: {rp_mp_pp}')
        # Для вывода срока действия в месяцах
        # for value, count in muns_count.items():
        # print(f"Значение: {value}, Количество: {count}")
        
     # Сортировка по столбцу "Срок действия в месяцах"
    def toggle_sort_order(self, column):
        if column == 3:  # Сортировка по 4 столбцу
            if self.sort_orders[column] == Qt.AscendingOrder:
                self.table_widget1.sortItems(3, Qt.AscendingOrder)
                self.sort_orders[column] = Qt.DescendingOrder
            else:
                self.table_widget1.sortItems(3, Qt.DescendingOrder)
                self.sort_orders[column] = Qt.AscendingOrder
        

if __name__ == '__main__':  # Поправлено name на __name__
    app = QApplication(sys.argv)
    gui = MyGUI()
    gui.show()

    sys.exit(app.exec_())
