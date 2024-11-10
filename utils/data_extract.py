from docx import Document

class DateExtractor:
    def __init__(self, docx_file):
        self.docx_file = docx_file
        self.start_date = None
        self.end_date = None

    def extract_dates(self):
        doc = Document(self.docx_file)
       
        start_date_found = False
        end_date_found = False

        for idx, para in enumerate(doc.paragraphs):
            if start_date_found:
                start_date = para.text.strip()
                if start_date:
                    self.start_date = start_date
                    start_date_found = False
                        
            if end_date_found:
                end_date = para.text.strip()
                if end_date:
                    self.end_date = end_date
                    end_date_found = False
                    
            if "Дата введения услуги в продажу (первый день продажи):" in para.text:
                start_date_found = True
            elif "Дата окончания продажи (последний день продажи):" in para.text:
                end_date_found = True
              
# путь к файлу .docx
# docx_file_path = 'butovo.docx'

# Создаем экземпляр класса DateExtractor
# date_extractor = DateExtractor(docx_file_path)

# Извлекаем даты
# date_extractor.extract_dates()

# Получаем даты из атрибутов класса
# start_date = date_extractor.start_date
# end_date = date_extractor.end_date

# print("Дата введения услуги в продажу (первый день продажи):", start_date)
# print("Дата окончания продажи (последний день продажи):", end_date)
