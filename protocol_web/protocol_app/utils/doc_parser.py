import re
from docx import Document  # правильный импорт
import docx2python  # правильный импорт
import pandas as pd
from .address_parser import parse_moscow_address


class DocParser:
    """Парсер документов Word для протоколов измерений"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.doc = Document(file_path)
        # Исправляем импорт и использование docx2python
        from docx2python import docx2python
        self.doc_text = docx2python(file_path)
        self.all_text = self.doc_text.text
        self.all_tables = self.doc.tables

    def context_text(self, search_text, length=200):
        """Поиск контекста вокруг текста"""
        position = self.all_text.find(search_text)
        if position == -1:
            return -1
        return self.all_text[position:position + length]

    def extract_protocol_number(self):
        """Извлечение номера протокола"""
        try:
            if self.context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №') != -1:
                num = self.context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №').split('№ ')[1].split('\n')[0].split(' от ')[0]
            else:
                num = self.context_text('ПРОТОКОЛ ИССЛЕДОВАНИЙ №').split('№ ')[1].split('\n')[0].split(' от ')[0]
            return num.strip()
        except Exception as e:
            return None

    def extract_date_protocol(self):
        """Извлечение даты протокола"""
        try:
            if self.context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №') != -1:
                date = self.context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №').split('№ ')[1].split('\n')[0].split(' от ')[1]
            else:
                date = self.context_text('ПРОТОКОЛ ИССЛЕДОВАНИЙ №').split('№ ')[1].split('\n')[0].split(' от ')[1]
            return date.strip()
        except Exception:
            return None

    def check_accreditation(self):
        """Проверка наличия аккредитации"""
        return 'есть' if self.context_text('Уникальный номер записи об') != -1 else 'нет'

    def get_protocol_type(self):
        """Определение типа протокола"""
        accreditation = self.check_accreditation()
        pref_acc = '_uniq' if accreditation == 'есть' else '_dop'

        if self.context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №') != -1:
            type_prot = 'pel'
        else:
            type_prot = 'ai'

        return f'{type_prot}{pref_acc}'

    def extract_measurement_datetime(self):
        """Извлечение даты и времени измерений"""
        try:
            data_izmereniya = self.context_text('Дата и время нача').split('\n\n')[1]
            time_start = self.context_text('Дата и время нача').split('\n\n')[4]

            if self.context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №') != -1:
                time_end = self.context_text('Дата и время окончания изме').split('\n\n')[2]
            else:
                time_end = self.context_text('Дата и время окончания').split('\n\n')[2]

            return {
                'data_izmereniya': data_izmereniya.strip(),
                'time_start_izmereniya': time_start.strip(),
                'time_end_izmereniya': time_end.strip()
            }
        except Exception:
            return {
                'data_izmereniya': None,
                'time_start_izmereniya': None,
                'time_end_izmereniya': None
            }

    def extract_location(self):
        """Извлечение адреса"""
        try:
            df_structura = []
            for row in self.all_tables[1].rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text.strip())
                df_structura.append(row_data)

            df_hat = pd.DataFrame(df_structura)
            location = df_hat.loc[3, 1]
            return parse_moscow_address(location)
        except Exception:
            return {
                'okrug': '-',
                'region': '-',
                'nazvanie_territorii': '-'
            }

    def extract_measurements(self):
        """Извлечение результатов измерений"""
        measurements = {}

        try:
            # Читаем таблицу с данными
            df_structura = []
            for row in self.all_tables[3].rows:
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    row_data.append(cell_text)
                df_structura.append(row_data)

            df_data = pd.DataFrame(df_structura)

            # Определяем тип протокола
            is_pel = self.context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №') != -1

            if is_pel:
                df_data = df_data[[2, 3]]
                df_data.columns = df_data.iloc[0]
                df_data = df_data[1:]

                # Обработка средних значений
                df = df_data.copy()
                mask = df['Наименование показателя'].str.contains('Среднее значение', na=False)
                previous_values = df['Наименование показателя'].shift(1)
                df.loc[mask, 'Наименование показателя'] = df.loc[
                                                              mask, 'Наименование показателя'] + ' ' + previous_values

                df.loc[df['Наименование показателя'].str.contains('Среднее значение'), 'result_col'] = df[
                    'Наименование показателя']
                df_data = df.loc[~df['result_col'].isna()].copy()
                df_data['name_pokazatel'] = df_data['result_col'].str.replace('Среднее значение ', '')

            else:  # AI
                # Ищем строку с "Примечание:"
                note_index = None
                for idx, val in df_data[0].items():
                    if val and 'Примечание:' in str(val):
                        note_index = idx
                        break

                if note_index:
                    df_data = df_data.iloc[:note_index]

                df_data = df_data[[2, 3]]
                df_data.columns = df_data.iloc[0]
                df_data = df_data[1:]
                df_data = df_data.reset_index(drop=True)
                df_data = df_data.rename(columns={'Наименование показателя': 'name_pokazatel'})

            # Ищем столбец с результатами
            result_col = None
            for col in df_data.columns:
                if col and 'Результат измерений' in str(col):
                    result_col = col
                    break

            if result_col:
                df_result = df_data[['name_pokazatel', result_col]].copy()
                df_result['name_pokazatel'] = df_result['name_pokazatel'].str.replace(',', '')

                # Загружаем словарь соответствия
                try:
                    import os
                    from django.conf import settings

                    # Путь к файлу Columns_02_top.xlsx
                    excel_path = os.path.join(settings.BASE_DIR, 'Columns_02_top.xlsx')

                    if os.path.exists(excel_path):
                        df_kirill = pd.read_excel(excel_path, sheet_name='База')
                        df_result = df_result.merge(df_kirill, how='left',
                                                    left_on='name_pokazatel',
                                                    right_on='old_name')

                        for _, row in df_result.iterrows():
                            if pd.notna(row.get('new_pokazat_name')):
                                field_name = self._normalize_field_name(row['new_pokazat_name'])
                                measurements[field_name] = str(row[result_col]) if pd.notna(row[result_col]) else None
                except Exception as e:
                    print(f"Ошибка при загрузке словаря: {e}")

        except Exception as e:
            print(f"Ошибка при извлечении измерений: {e}")

        return measurements

    def _normalize_field_name(self, name):
        """Нормализация имени поля для модели"""
        # Заменяем русские символы и пробелы
        translit_map = {
            'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
            'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
            'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
            'ф': 'f', 'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
            'ы': 'y', 'э': 'e', 'ю': 'yu', 'я': 'ya', ' ': '_', '-': '_', ',': '',
            '(': '', ')': '', '±': '', '+': '', '/': '_', '\\': '_', '.': ''
        }

        normalized = name.lower()
        for rus, eng in translit_map.items():
            normalized = normalized.replace(rus, eng)

        # Убираем множественные подчеркивания
        normalized = re.sub(r'_+', '_', normalized)
        normalized = normalized.strip('_')

        return normalized

    def parse(self):
        """Основной метод парсинга документа"""
        result = {
            'rejd_na_istochnik': '-',
            'number_protocol': self.extract_protocol_number(),
            'date_protocol': self.extract_date_protocol(),
            'type_protocol': self.get_protocol_type(),
        }

        # Добавляем адрес
        location = self.extract_location()
        result.update(location)

        # Добавляем дату и время
        dt = self.extract_measurement_datetime()
        result.update(dt)

        # Добавляем измерения
        measurements = self.extract_measurements()
        result.update(measurements)

        return result