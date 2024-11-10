from PyQt5.QtGui import QIcon, QPixmap, QColor
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QTableView, QVBoxLayout, QWidget

def create_table_model(data, headers_list):
    model = QTableWidget()
    model.setColumnCount(len(data.columns))
    model.setRowCount(len(data.index))
    
    headers_list = ['Внесен', 'Название в учетной системе', 'Тип', 
                  'Срок действия в месяцах', 'Визиты',  
                  'Время посещения', 'Гостевые визиты', 
                  'Заморозка', 'Статус', 'Подарочная ФД', 
                  'Подарочная ФД', 'inBody', 
                  'Иные условия(входят в стоимость карты)', 
                  'Стоимость карты', 'Стоимость месяца/визита', 
                  'Стоимость месяца в договоре']

    if headers_list:
        model.setHorizontalHeaderLabels(headers_list)

        for row in range(len(data.index)):
            for column in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[row, column]))
                model.setItem(row, column, item)

        table_view = QTableView()
        table_view.setModel(model)  # Устанавливаем модель в QTableView

        # Создаем виджет, в котором будет отображаться таблица
        table_widget = QWidget()
        table_layout = QVBoxLayout()
        table_layout.addWidget(table_view)
        table_widget.setLayout(table_layout)

        return table_widget

 
