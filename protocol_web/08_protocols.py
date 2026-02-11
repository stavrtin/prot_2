'''
В данной версии должно быть улучшено:
 - вывод в эксельоформлен прилично (Заголовок), ширина колонок
 - добавлена функция выгрузки одного протокола в Excel

'''

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
# import numpy as np
from pathlib import Path
from docx import Document
from docx2python import docx2python
from sqlalchemy import create_engine
import logging
from datetime import datetime
import os
import time
import config
import xlsxwriter
from sqlalchemy import create_engine, text
import re

# Глобальные переменные
selected_files = []
engine = None
columns_db = []


# Настройка логирования в файл
def setup_logging():
    """Настройка логирования в файл"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_filename = f"protocol_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_path = os.path.join(log_dir, log_filename)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler()  # Также выводим в консоль
        ]
    )

    return logging.getLogger(__name__)

def log_message(log_text, message):
    """Добавляет сообщение в лог интерфейса и файл"""
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.update()
    logger.info(message)  # Также пишем в файловый лог

# Инициализируем логгер
logger = setup_logging()


def init_app():
    """Инициализация приложения с ПРИНУДИТЕЛЬНЫМ TCP подключением"""
    global engine, columns_db

    try:
        logger.info(f"Инициализация БД подключения к {config.DB_IP}:{config.DB_PORT}")
        print(f"DEBUG: Инициализация БД подключения к {config.DB_IP}:{config.DB_PORT}")

        # ЯВНОЕ создание TCP подключения через psycopg2
        import psycopg2

        def create_tcp_connection():
            """Функция создает TCP подключение, игнорируя Unix socket"""
            return psycopg2.connect(
                host=config.DB_IP,
                port=int(config.DB_PORT),
                dbname=config.DB_NAME,
                user=config.DB_LOGIN,
                password=config.DB_PASSW,
                connect_timeout=10
            )

        # Создаем engine с нашей функцией подключения
        engine = create_engine(
            'postgresql://',  # Пустой URL, т.к. все параметры в функции
            creator=create_tcp_connection,
            echo=False,
            pool_pre_ping=True,
            pool_recycle=3600
        )

        # Тестируем подключение
        print("DEBUG: Тестируем TCP подключение...")
        with engine.connect() as conn:
            result = conn.execute(text("SELECT version()"))
            version = result.fetchone()[0]
            print(f"DEBUG: ✓ PostgreSQL: {version.split(',')[0]}")

        # Загружаем колонки из БД
        print("DEBUG: Загружаем структуру таблицы protocols...")
        df_first_row = pd.read_sql_query('SELECT * FROM protocols LIMIT 1', con=engine)
        columns_db = df_first_row.columns.to_list()
        print(f"DEBUG: Загружено колонок: {len(columns_db)}")

        logger.info(f"Приложение инициализировано, {len(columns_db)} колонок")

    except Exception as e:
        error_msg = f"Ошибка инициализации БД: {e}"
        print(f"DEBUG: ❌ {error_msg}")
        logger.error(error_msg)
        import traceback
        traceback.print_exc()


def select_files_callback(file_label, process_btn, log_text):
    """Обработчик выбора нескольких файлов"""
    global selected_files

    file_paths = filedialog.askopenfilenames(
        title="Выберите файлы протоколов",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
    )

    if file_paths:
        selected_files = list(file_paths)
        file_names = [Path(file_path).name for file_path in selected_files]
        file_label.config(
            text=f"Выбрано файлов: {len(selected_files)}\n{', '.join(file_names[:3])}{'...' if len(file_names) > 3 else ''}")
        process_btn.config(state="normal")
        log_message(log_text, f"Выбрано файлов: {len(selected_files)}")
        for file_path in selected_files:
            log_message(log_text, f"  - {file_path}")
        logger.info(f"Выбрано {len(selected_files)} файлов для обработки")


def check_protocol_exists(number_protocol):
    """Проверяет, существует ли протокол с таким номером в БД"""
    try:
        query = f"SELECT COUNT(*) as count FROM protocols WHERE number_protocol = '{number_protocol}'"
        result = pd.read_sql_query(query, con=engine)
        return result['count'].iloc[0] > 0
    except Exception as e:
        logger.error(f"Ошибка при проверке протокола в БД: {e}")
        return False


def process_files_callback(log_text, progress_label):
    """Обработчик кнопки 'Обработать' для нескольких файлов"""
    global selected_files

    if not selected_files:
        messagebox.showerror("Ошибка", "Файлы не выбраны!")
        return

    try:
        log_message(log_text, f"Начало обработки {len(selected_files)} файлов...")
        progress_label.config(text="Обработка...")

        successful_count = 0
        skipped_count = 0
        error_count = 0

        for i, file_path in enumerate(selected_files, 1):
            log_message(log_text, f"\n--- Обработка файла {i}/{len(selected_files)}: {Path(file_path).name} ---")

            # Вызов основной логики обработки
            result = process_protocol(file_path, log_text)

            if result == "success":
                successful_count += 1
            elif result == "skipped":
                skipped_count += 1
            elif result == "error":
                error_count += 1

        # Итоговый отчет
        log_message(log_text, f"\n=== ОБРАБОТКА ЗАВЕРШЕНА ===")
        log_message(log_text, f"Успешно обработано: {successful_count}")
        log_message(log_text, f"Пропущено (уже в БД): {skipped_count}")
        log_message(log_text, f"С ошибками: {error_count}")

        progress_label.config(text="Готово")

        if successful_count > 0:
            messagebox.showinfo("Успех",
                                f"Обработка завершена!\nУспешно: {successful_count}\nПропущено: {skipped_count}\nОшибки: {error_count}")
        else:
            messagebox.showwarning("Предупреждение",
                                   f"Не удалось обработать ни одного файла!\nПропущено: {skipped_count}\nОшибки: {error_count}")

    except Exception as e:
        error_msg = f"Ошибка при обработке файлов: {str(e)}"
        log_message(log_text, error_msg)
        messagebox.showerror("Ошибка", error_msg)
        progress_label.config(text="Ошибка")
        logger.error(error_msg)


def search_protocol_callback(search_entry, log_text):
    """Обработчик поиска протокола по номеру (частичное совпадение)"""
    search_text = search_entry.get().strip()

    # Проверяем, не является ли текст подсказкой
    if search_text == "Введите часть номера..." or not search_text:
        messagebox.showwarning("Предупреждение", "Введите номер или часть номера протокола для поиска!")
        return

    try:
        log_message(log_text, f"Поиск протоколов по: '{search_text}'...")

        # Используем параметризованный запрос для безопасности
        from sqlalchemy import text

        # Создаем SQL-запрос с параметром - ищем ТОЛЬКО по number_protocol
        query = text("""
        SELECT 
            number_protocol,
            date_protocol,
            type_protocol
        FROM protocols 
        WHERE number_protocol ILIKE :search_pattern
        ORDER BY number_protocol
        LIMIT 100
        """)

        # Подготавливаем паттерн для поиска
        search_pattern = f"%{search_text}%"

        # Выполняем запрос с параметрами
        with engine.connect() as conn:
            result = conn.execute(query, {"search_pattern": search_pattern})
            rows = result.fetchall()

        # Преобразуем результат в DataFrame
        if rows:
            result_df = pd.DataFrame(rows, columns=['number_protocol', 'date_protocol', 'type_protocol'])
        else:
            result_df = pd.DataFrame(columns=['number_protocol', 'date_protocol', 'type_protocol'])

        if not result_df.empty:
            count = len(result_df)

            if count == 1:
                # Найден ровно один протокол
                protocol = result_df.iloc[0]
                log_message(log_text, f"✓ Найден 1 протокол:")
                log_message(log_text, f"  Номер: {protocol['number_protocol']}")
                log_message(log_text, f"  Дата: {protocol['date_protocol']}")
                log_message(log_text, f"  Тип: {protocol['type_protocol']}")

                messagebox.showinfo("Результат поиска",
                                    f"Найден протокол:\n"
                                    f"№{protocol['number_protocol']}\n"
                                    f"Дата: {protocol['date_protocol']}\n"
                                    f"Тип: {protocol['type_protocol']}")
            else:
                # Найдено несколько протоколов
                log_message(log_text, f"✓ Найдено {count} протоколов, содержащих '{search_text}':")

                # Выводим примеры найденных номеров (первые 10)
                examples_count = min(10, count)
                example_numbers = result_df['number_protocol'].head(examples_count).tolist()

                for i, num in enumerate(example_numbers, 1):
                    log_message(log_text, f"  {i}. {num}")

                if count > 10:
                    log_message(log_text, f"  ... и еще {count - 10} протоколов")

                # Получаем статистику по типам протоколов
                protocol_types = result_df['type_protocol'].value_counts()
                log_message(log_text, f"  Распределение по типам:")
                for protocol_type, type_count in protocol_types.items():
                    log_message(log_text, f"    - {protocol_type}: {type_count}")

                # Сообщение пользователю
                if count <= 10:
                    numbers_list = "\n".join([f"{i}. {num}" for i, num in enumerate(example_numbers, 1)])
                    messagebox.showinfo("Результат поиска",
                                        f"Найдено {count} протоколов:\n\n{numbers_list}")
                else:
                    first_numbers = "\n".join([f"{i}. {num}" for i, num in enumerate(example_numbers[:5], 1)])
                    messagebox.showinfo("Результат поиска",
                                        f"Найдено {count} протоколов!\n\n"
                                        f"Примеры (первые 5):\n{first_numbers}\n\n"
                                        f"... и еще {count - 5} протоколов.")
        else:
            # Не найдено ни одного протокола
            log_message(log_text, f"✗ Протоколов, содержащих '{search_text}', не найдено")

            # Пробуем предложить возможные варианты (похожие номера)
            try:
                suggestion_query = text("""
                SELECT number_protocol
                FROM protocols 
                ORDER BY number_protocol
                LIMIT 10
                """)

                with engine.connect() as conn:
                    suggestions_result = conn.execute(suggestion_query)
                    suggestions_rows = suggestions_result.fetchall()

                if suggestions_rows:
                    suggestions = [row[0] for row in suggestions_rows]
                    log_message(log_text, f"  Примеры протоколов в базе:")
                    for i, suggestion in enumerate(suggestions, 1):
                        log_message(log_text, f"    {i}. {suggestion}")
            except Exception as suggestion_error:
                logger.error(f"Ошибка при получении предложений: {suggestion_error}")
                # Пропускаем ошибку предложений

            messagebox.showinfo("Результат поиска",
                                f"Протоколов, содержащих '{search_text}', не найдено.")

    except Exception as e:
        error_msg = f"Ошибка при поиске протокола: {str(e)}"
        log_message(log_text, f"ОШИБКА: {error_msg}")
        messagebox.showerror("Ошибка", error_msg)
        logger.error(error_msg)
        import traceback
        logger.error(traceback.format_exc())

# ------------------------------------

def export_to_excel_callback(log_text):
    """Обработчик выгрузки данных в Excel"""
    try:
        log_message(log_text, "Подготовка выгрузки в Excel...")

        # Получаем все данные из БД
        query = "SELECT * FROM protocols"
        df = pd.read_sql_query(query, con=engine)

        # -------------------------- Перефигачим шапку (названия на кириллице - в файле Columns_02_top.xlsx )
        df_col = pd.read_excel('Columns_02_top.xlsx')
        kirill = df_col.old_name.to_list()
        latin = df_col.new_pokazat_name.to_list()
        # --------------------------------------- из двух списков делаю словарь ----------------
        # --------------------------------------- пробегаю по колонкам выходного датафрейма и заменяю на значения из словаря --------
        dict_trans = {}
        for i in latin:
            dict_trans[i] = kirill[latin.index(i)]

        list_combats_columns = df.columns.to_list()
        for i in list_combats_columns:
            if i in dict_trans:
                # ----------- меняю на кириллицу ----------------
                df.rename(columns={i: dict_trans[i]}, inplace=True)
            else:
                pass

        if df.empty:
            messagebox.showwarning("Предупреждение", "База данных пуста!")
            log_message(log_text, "База данных пуста, выгрузка отменена")
            return

        # Предлагаем выбрать место для сохранения
        file_path = filedialog.asksaveasfilename(
            title="Сохранить как Excel файл",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if file_path:
            # Создаем Excel writer с xlsxwriter
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet('Протоколы')

            # Создаем форматы
            # Формат для заголовков A-J (зеленый)
            header_format_a_j = workbook.add_format({
                'text_wrap': True,  # Перенос текста
                'valign': 'vcenter',  # Вертикальное выравнивание по центру
                'align': 'center',  # Горизонтальное выравнивание по центру
                'bold': True,  # Жирный шрифт
                'bg_color': '#44c14d',  # Зеленый цвет фона
                'border': 1,  # Границы
                'font_size': 9  # Размер шрифта
            })

            # Формат для заголовков K-HD (серый)
            header_format_k_hd = workbook.add_format({
                'text_wrap': True,  # Перенос текста
                'valign': 'vcenter',  # Вертикальное выравнивание по центру
                'align': 'center',  # Горизонтальное выравнивание по центру
                'bold': True,  # Жирный шрифт
                'bg_color': '#D9D9D9',  # Серый цвет фона
                'border': 1,  # Границы
                'font_size': 9  # Размер шрифта
            })

            # Формат для данных (ячейки с данными)
            data_format = workbook.add_format({
                'border': 1,  # Легкая рамка
                'font_size': 9  # Размер шрифта
            })

            # Формат для дат
            date_format = workbook.add_format({
                'border': 1,  # Легкая рамка
                'font_size': 9,  # Размер шрифта
                'num_format': 'DD.MM.YYYY'  # Формат даты
            })

            # Получаем заголовки
            headers = list(df.columns)
            num_cols = len(headers)
            num_rows = len(df)

            # Записываем заголовки
            for col_num, header in enumerate(headers):
                if col_num <= 9:  # Колонки A-J (0-9)
                    worksheet.write(0, col_num, header, header_format_a_j)
                else:  # Колонки K-HD (10+)
                    worksheet.write(0, col_num, header, header_format_k_hd)

            # Записываем данные
            for row_num in range(num_rows):
                for col_num in range(num_cols):
                    cell_value = df.iat[row_num, col_num]

                    # Проверяем тип данных для форматирования
                    if isinstance(cell_value, (pd.Timestamp, datetime)):
                        worksheet.write(row_num + 1, col_num, cell_value, date_format)
                    else:
                        worksheet.write(row_num + 1, col_num, cell_value, data_format)

            # Устанавливаем высоту строки для заголовка
            worksheet.set_row(0, 30)

            # Устанавливаем ширину колонок
            column_widths = {
                0: 15,  # A: Рейд на источник
                1: 15,  # B: Номер протокола
                2: 15,  # C: Дата протокола
                3: 30,  # D: Округ
                4: 30,  # E: Район
                5: 30,  # F: Название территории
                6: 12,  # G: Дата измерения
                7: 18,  # H: Время начала измерения
                8: 20,  # I: Время завершения измерения
                9: 15,  # J: Тип протокола
            }

            # Устанавливаем заданную ширину для A-J
            for col_num, width in column_widths.items():
                worksheet.set_column(col_num, col_num, width)

            # Автоматическая ширина для остальных колонок
            for col_num in range(num_cols):
                if col_num not in column_widths:
                    # Определяем максимальную длину в колонке
                    max_length = 0
                    # Проверяем заголовок
                    header_len = len(str(headers[col_num]))
                    max_length = max(max_length, header_len)

                    # Проверяем данные
                    for row_num in range(num_rows):
                        cell_value = df.iat[row_num, col_num]
                        if cell_value is not None:
                            cell_len = len(str(cell_value))
                            max_length = max(max_length, cell_len)

                    # Устанавливаем ширину с запасом
                    worksheet.set_column(col_num, col_num, min(max_length + 2, 50))

            # Устанавливаем фильтр на заголовок
            # Преобразуем индекс последней колонки в буквенное обозначение
            last_col_letter = xlsxwriter.utility.xl_col_to_name(num_cols - 1)
            filter_range = f'A1:{last_col_letter}{num_rows + 1}'
            worksheet.autofilter(filter_range)

            # Закрепляем заголовок
            worksheet.freeze_panes(1, 0)

            # Закрываем книгу
            workbook.close()

            log_message(log_text, f"✓ Данные успешно выгружены в файл: {file_path}")
            log_message(log_text, f"  Количество записей: {num_rows}")
            log_message(log_text, f"  Количество колонок: {num_cols}")
            log_message(log_text, f"  Диапазон фильтра: {filter_range}")

            messagebox.showinfo("Успех",
                                f"Данные успешно выгружены в Excel!\nФайл: {Path(file_path).name}\nЗаписей: {num_rows}")

    except Exception as e:
        error_msg = f"Ошибка при выгрузке в Excel: {str(e)}"
        log_message(log_text, f"ОШИБКА: {error_msg}")
        messagebox.showerror("Ошибка", error_msg)
        logger.error(error_msg)

# ------------------------------------------------- дак-------------
def export_single_protocol_callback(protocol_entry, log_text):
    """Выгрузка одного протокола в Excel с дополнительным раскрашиванием ячеек."""
    protocol_number = protocol_entry.get().strip()

    # -------------------- проверка ввода --------------------
    if protocol_number == "Введите номер протокола..." or not protocol_number:
        messagebox.showwarning("Предупреждение", "Введите номер протокола для выгрузки!")
        return

    try:
        log_message(log_text,
                    f"Поиск протокола №{protocol_number} для выгрузки в Excel...")

        # -------------------- проверка наличия протокола --------------------
        cnt_q = text("""
            SELECT COUNT(*) AS cnt
            FROM protocols
            WHERE number_protocol = :protocol_number
        """)
        with engine.connect() as conn:
            cnt_res = conn.execute(cnt_q,
                                   {"protocol_number": protocol_number}).fetchone()

        if cnt_res is None or cnt_res[0] == 0:
            log_message(log_text,
                        f"✗ Протокол №{protocol_number} не найден в базе данных")
            messagebox.showwarning("Предупреждение",
                                   f"Протокол №{protocol_number} не найден в базе данных!")
            return
        if cnt_res[0] > 1:
            log_message(log_text,
                        f"⚠ Найдено {cnt_res[0]} протоколов с номером №{protocol_number}")
            messagebox.showwarning("Предупреждение",
                                   f"Найдено {cnt_res[0]} протоколов с номером №{protocol_number}!\n"
                                   "Выгрузка отменена.")
            return

        # -------------------- чтение данных --------------------
        df = pd.read_sql_query(
            text("SELECT * FROM protocols WHERE number_protocol = :protocol_number"),
            con=engine,
            params={"protocol_number": protocol_number},
        )

        # -------------------- переименование колонок --------------------
        df_col = pd.read_excel('Columns_02_top.xlsx')
        trans = dict(zip(df_col.new_pokazat_name, df_col.old_name))
        df.rename(columns=trans, inplace=True)

        # -------------------- трансформация в «параметр‑значение» --------------------
        df_one = df.T.reset_index().rename(columns={'index': 'Параметр', 0: 'Значение'})
        df_one = df_one.loc[~df_one['Значение'].isna()]

        # -------------------- подготовка пути сохранения --------------------
        cleaned_protocol_number = clean_filename(protocol_number)
        file_path = filedialog.asksaveasfilename(
            title="Сохранить протокол как Excel файл",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"Протокол_{cleaned_protocol_number}.xlsx",
        )
        if not file_path:
            return

        # -------------------- создание книги --------------------
        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet('Протокол')

        # -------------------- форматы --------------------
        header_format = workbook.add_format({
            'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'bold': True, 'bg_color': '#44c14d', 'border': 1, 'font_size': 9
        })
        data_format = workbook.add_format({'border': 1, 'font_size': 9})
        date_format = workbook.add_format({
            'border': 1, 'font_size': 9, 'num_format': 'DD.MM.YYYY'
        })

        # **Новый** формат для заголовка строки 2 (ячейки A2 и B2)
        title_row_format = workbook.add_format({
            'bg_color': '#D5ABF5', 'border': 1, 'font_size': 9,
            'align': 'center', 'valign': 'vcenter'
        })

        # **Новый** формат для ячеек A3‑A12
        early_rows_format = workbook.add_format({
            'bg_color': '#EDEDED', 'border': 1, 'font_size': 9,
            'align': 'center', 'valign': 'vcenter'
        })

        # **Новый** формат для ячеек A13‑и далее (заполненные)
        later_rows_format = workbook.add_format({
            'bg_color': '#B8CDF5', 'border': 1, 'font_size': 9,
            'align': 'center', 'valign': 'vcenter'
        })

        # -------------------- заголовок книги --------------------
        worksheet.merge_range(
            'A1:B1',
            f'Протокол №{protocol_number}',
            workbook.add_format({
                'bold': True, 'font_size': 14,
                'align': 'center', 'valign': 'vcenter'
            })
        )

        # -------------------- заголовки колонок (строка 2) --------------------
        headers = list(df_one.columns)
        for col_num, hdr in enumerate(headers):
            worksheet.write(1, col_num, hdr, header_format)   # строка 2 (индекс 1)

        # -------------------- раскраска ячеек A2‑B2 --------------------
        # A2 и B2 – это первая и вторая колонка заголовка строки 2
        worksheet.write(1, 0, headers[0], title_row_format)
        if len(headers) > 1:
            worksheet.write(1, 1, headers[1], title_row_format)

        # -------------------- запись данных --------------------
        num_rows, num_cols = df_one.shape
        for row_idx in range(num_rows):
            for col_idx in range(num_cols):
                val = df_one.iat[row_idx, col_idx]

                # выбираем формат в зависимости от строки
                if row_idx < 10:                     # строки 3‑12 (0‑based → 2‑11)
                    cell_fmt = early_rows_format
                else:
                    cell_fmt = later_rows_format

                # если это дата – используем отдельный формат
                if isinstance(val, (pd.Timestamp, datetime)):
                    worksheet.write(row_idx + 2, col_idx, val, date_format)
                else:
                    worksheet.write(row_idx + 2, col_idx, val, cell_fmt)

        # -------------------- размеры строк/столбцов --------------------
        worksheet.set_row(0, 25)   # строка 1 – заголовок книги
        worksheet.set_row(1, 30)   # строка 2 – заголовки колонок

        for col_idx in range(num_cols):
            max_len = max(
                len(str(headers[col_idx])),
                *(len(str(df_one.iat[r, col_idx])) for r in range(num_rows) if df_one.iat[r, col_idx] is not None)
            )
            worksheet.set_column(col_idx, col_idx, min(max_len + 2, 50))

        # -------------------- фиксируем заголовки и автофильтр --------------------
        worksheet.freeze_panes(2, 0)   # фиксируем строку 2
        if num_rows:
            last_col = xlsxwriter.utility.xl_col_to_name(num_cols - 1)
            worksheet.autofilter(f'A2:{last_col}{num_rows + 1}')

        workbook.close()

        # -------------------- логирование и сообщение пользователю --------------------
        log_message(log_text,
                    f"✓ Протокол №{protocol_number} успешно выгружен в файл: {file_path}")
        log_message(log_text, f"  Количество записей: {num_rows}")

        messagebox.showinfo(
            "Успех",
            f"Протокол №{protocol_number} успешно выгружен!\n"
            f"Файл: {Path(file_path).name}\n"
            f"Количество записей: {num_rows}"
        )

    except Exception as e:
        err = f"Ошибка при выгрузке протокола: {e}"
        log_message(log_text, f"ОШИБКА: {err}")
        messagebox.showerror("Ошибка", err)
        logger.error(err)


def clean_filename(filename):
    """
    Очищает имя файла от недопустимых символов.
    Заменяет: / \ : * ? " < > | на -
    """
    # Список недопустимых символов в именах файлов Windows
    invalid_chars = r'[<>:"/\\|?*]'

    # Заменяем недопустимые символы на дефис
    cleaned = re.sub(invalid_chars, '-', filename)

    # Убираем лишние дефисы (подряд идущие)
    cleaned = re.sub(r'-+', '-', cleaned)

    # Убираем дефисы в начале и конце
    cleaned = cleaned.strip('-')

    # Если после очистки имя стало пустым, возвращаем стандартное имя
    if not cleaned:
        cleaned = 'Протокол'

    return cleaned


# Также обновим подсказку в интерфейсе
# -------------------------------копирование + ---------------------
def create_gui():
    """Создание графического интерфейса"""
    root = tk.Tk()
    root.title("Обработчик протоколов")
    root.geometry("750x750")  # Увеличиваем высоту для нового фрейма

    # Заголовок
    title_label = tk.Label(root, text="Обработка протоколов ПЭЛ/АИ",
                           font=("Arial", 16, "bold"))
    title_label.pack(pady=10)

    # Фрейм для поиска протокола
    search_frame = tk.LabelFrame(root, text="Поиск протокола", font=("Arial", 10, "bold"), padx=10, pady=10)
    search_frame.pack(pady=10, padx=20, fill=tk.X)

    # Поле ввода для поиска с поддержкой вставки
    tk.Label(search_frame, text="Номер протокола:", font=("Arial", 9)).grid(row=0, column=0, padx=5, pady=5, sticky="w")

    # Создаем поле ввода с подсказкой
    search_entry = tk.Entry(search_frame, width=40, font=("Arial", 9))
    search_entry.insert(0, "Введите часть номера...")
    search_entry.config(fg="grey")

    def on_entry_click(event):
        """Обработчик клика по полю ввода"""
        if search_entry.get() == "Введите часть номера...":
            search_entry.delete(0, tk.END)
            search_entry.config(fg="black")

    def on_focusout(event):
        """Обработчик потери фокуса"""
        if search_entry.get() == "":
            search_entry.insert(0, "Введите часть номера...")
            search_entry.config(fg="grey")

    search_entry.bind('<FocusIn>', on_entry_click)
    search_entry.bind('<FocusOut>', on_focusout)

    # Привязываем Enter для запуска поиска
    def on_enter_pressed(event):
        if search_entry.get() != "Введите часть номера...":
            search_protocol_callback(search_entry, log_text)

    search_entry.bind('<Return>', on_enter_pressed)  # Enter запускает поиск

    search_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    # Функции для работы с буфером обмена
    def paste_from_clipboard(event=None):
        """Вставить из буфера обмена в поле поиска"""
        try:
            clipboard_text = root.clipboard_get()
            if clipboard_text:
                # Очищаем поле, если там подсказка
                if search_entry.get() == "Введите часть номера...":
                    search_entry.delete(0, tk.END)
                    search_entry.config(fg="black")

                # Вставляем текст из буфера
                search_entry.insert(tk.INSERT, clipboard_text)
            return "break"
        except tk.TclError:
            return "break"

    def copy_to_clipboard(event=None):
        """Копировать выделенный текст из лога в буфер обмена"""
        try:
            if log_text.tag_ranges(tk.SEL):
                # Копируем выделенный текст
                selected_text = log_text.get(tk.SEL_FIRST, tk.SEL_LAST)
                root.clipboard_clear()
                root.clipboard_append(selected_text)
            return "break"
        except:
            return "break"

    # Создаем контекстное меню для поля поиска
    search_context_menu = tk.Menu(search_frame, tearoff=0)
    search_context_menu.add_command(label="Вставить", command=paste_from_clipboard)
    search_context_menu.add_separator()
    search_context_menu.add_command(label="Вырезать", command=lambda: search_entry.event_generate("<<Cut>>"))
    search_context_menu.add_command(label="Копировать", command=lambda: search_entry.event_generate("<<Copy>>"))

    def show_search_context_menu(event):
        search_context_menu.tk_popup(event.x_root, event.y_root)

    # Привязываем контекстное меню и горячие клавиши
    search_entry.bind("<Button-3>", show_search_context_menu)  # Правая кнопка мыши
    search_entry.bind("<Control-v>", paste_from_clipboard)  # Ctrl+V для вставки
    search_entry.bind("<Command-v>", paste_from_clipboard)  # Command+V для Mac

    # Кнопка поиска
    search_btn = tk.Button(search_frame, text="Найти протоколы",
                           font=("Arial", 9), bg="lightyellow",
                           command=lambda: search_protocol_callback(search_entry, log_text))
    search_btn.grid(row=0, column=2, padx=5, pady=5)

    # Добавляем подсказку под полем поиска
    search_hint = tk.Label(search_frame,
                           text="Можно ввести часть номера протокола. Enter - поиск",
                           font=("Arial", 8),
                           fg="gray")
    search_hint.grid(row=1, column=0, columnspan=3, padx=5, pady=(0, 5), sticky="w")

    # Настройка веса колонок для растягивания
    search_frame.columnconfigure(1, weight=1)

    # Фрейм для обработки файлов
    process_frame = tk.LabelFrame(root, text="Обработка файлов", font=("Arial", 10, "bold"), padx=10, pady=10)
    process_frame.pack(pady=10, padx=20, fill=tk.X)

    # Основной контейнер для выравнивания по центру
    process_container = tk.Frame(process_frame)
    process_container.pack(expand=True, fill=tk.X)

    # Верхний ряд: Кнопка выбора файлов
    select_frame = tk.Frame(process_container)
    select_frame.pack(pady=(0, 5))

    select_btn = tk.Button(select_frame, text="Выбрать файлы",
                           font=("Arial", 10),
                           bg="lightblue",
                           width=20,
                           height=1)
    select_btn.pack(side=tk.LEFT, padx=(0, 10))

    # Информация о выбранных файлах
    file_label = tk.Label(select_frame, text="Файлы не выбраны",
                          font=("Arial", 9),
                          wraplength=400,
                          justify="left",
                          anchor="w")
    file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # Средний ряд: Кнопка обработки (расположена ПРЯМО ПОД кнопкой "Выбрать файлы")
    process_buttons_frame = tk.Frame(process_container)
    process_buttons_frame.pack(pady=5)

    # Создаем контейнер для центрирования кнопки "Обработать все"
    center_frame = tk.Frame(process_buttons_frame)
    center_frame.pack()

    # Кнопка "Обработать все" - теперь выровнена по левому краю относительно кнопки "Выбрать файлы"
    process_btn = tk.Button(center_frame, text="Обработать все",
                            font=("Arial", 10),
                            bg="lightgreen",
                            width=20,
                            height=1,
                            state="disabled")


    # process_btn.pack()
    process_btn.pack(side=tk.BOTTOM,   # привязываем к нижнему краю контейнера
                fill=tk.X,        # растягиваем по ширине (по желанию)
                padx=5,           # небольшие отступы слева/справа, можно убрать
                pady=0)

    # Нижний ряд: Индикатор прогресса
    progress_frame = tk.Frame(process_container)
    progress_frame.pack(pady=(5, 0))

    progress_label = tk.Label(progress_frame, text="", font=("Arial", 9))
    progress_label.pack()

    # НОВЫЙ ФРЕЙМ: Выгрузить один протокол в Excel
    export_single_frame = tk.LabelFrame(root, text="Выгрузить один протокол в Эксель",
                                       font=("Arial", 10, "bold"), padx=10, pady=10)
    export_single_frame.pack(pady=10, padx=20, fill=tk.X)

    # Поле ввода для номера протокола
    tk.Label(export_single_frame, text="Номер протокола:",
            font=("Arial", 9)).grid(row=0, column=0, padx=5, pady=5, sticky="w")

    # Создаем поле ввода с подсказкой
    protocol_entry = tk.Entry(export_single_frame, width=40, font=("Arial", 9))
    protocol_entry.insert(0, "Введите номер протокола...")
    protocol_entry.config(fg="grey")

    def on_protocol_entry_click(event):
        """Обработчик клика по полю ввода номера протокола"""
        if protocol_entry.get() == "Введите номер протокола...":
            protocol_entry.delete(0, tk.END)
            protocol_entry.config(fg="black")

    def on_protocol_focusout(event):
        """Обработчик потери фокуса для поля протокола"""
        if protocol_entry.get() == "":
            protocol_entry.insert(0, "Введите номер протокола...")
            protocol_entry.config(fg="grey")

    protocol_entry.bind('<FocusIn>', on_protocol_entry_click)
    protocol_entry.bind('<FocusOut>', on_protocol_focusout)

    # Привязываем Enter для запуска выгрузки
    def on_protocol_enter_pressed(event):
        if protocol_entry.get() != "Введите номер протокола...":
            export_single_protocol_callback(protocol_entry, log_text)

    protocol_entry.bind('<Return>', on_protocol_enter_pressed)
    protocol_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    # Кнопка выгрузки одного протокола
    export_single_btn = tk.Button(export_single_frame, text="Выгрузить в Excel",
                                  font=("Arial", 9), bg="lightblue",
                                  command=lambda: export_single_protocol_callback(protocol_entry, log_text))
    export_single_btn.grid(row=0, column=2, padx=5, pady=5)

    # Подсказка для поля выгрузки одного протокола
    export_single_hint = tk.Label(export_single_frame,
                                  text="Введите точный номер протокола. Enter - выгрузка",
                                  font=("Arial", 8),
                                  fg="gray")
    export_single_hint.grid(row=1, column=0, columnspan=3, padx=5, pady=(0, 5), sticky="w")

    # Настройка веса колонок для растягивания
    export_single_frame.columnconfigure(1, weight=1)

    # Фрейм для выгрузки в Excel (всех данных)
    export_frame = tk.LabelFrame(root, text="Выгрузить все данные в Excel",
                                font=("Arial", 10, "bold"), padx=10, pady=10)
    export_frame.pack(pady=10, padx=20, fill=tk.X)

    # Описание функции выгрузки
    export_description = tk.Label(export_frame,
                                  text="Выгрузить все данные из базы данных в Excel файл",
                                  font=("Arial", 9),
                                  wraplength=600,
                                  justify="center")
    export_description.pack(pady=5)

    # Кнопка выгрузки в Excel
    export_btn = tk.Button(export_frame,
                           text="Выгрузить все в Excel",
                           font=("Arial", 10),
                           bg="lightgreen",
                           width=20,
                           height=1,
                           command=lambda: export_to_excel_callback(log_text))
    export_btn.pack(pady=10)

    # Текстовое поле для логов
    log_frame = tk.LabelFrame(root, text="Лог выполнения", font=("Arial", 10, "bold"), padx=10, pady=10)
    log_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

    log_text = tk.Text(log_frame, height=15, width=80, font=("Consolas", 9))
    log_text.pack(fill=tk.BOTH, expand=True)

    # Добавляем скроллбар для текстового поля
    scrollbar = tk.Scrollbar(log_text)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    log_text.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=log_text.yview)

    # Создаем контекстное меню для лога
    log_context_menu = tk.Menu(log_frame, tearoff=0)
    log_context_menu.add_command(label="Копировать", command=copy_to_clipboard)
    log_context_menu.add_command(label="Выделить все", command=lambda: log_text.tag_add(tk.SEL, "1.0", tk.END))
    log_context_menu.add_separator()
    log_context_menu.add_command(label="Очистить лог", command=lambda: log_text.delete(1.0, tk.END))

    def show_log_context_menu(event):
        log_context_menu.tk_popup(event.x_root, event.y_root)

    # Привязываем контекстное меню и горячие клавиши для лога
    log_text.bind("<Button-3>", show_log_context_menu)  # Правая кнопка мыши
    log_text.bind("<Control-c>", copy_to_clipboard)  # Ctrl+C для копирования
    log_text.bind("<Command-c>", copy_to_clipboard)  # Command+C для Mac
    log_text.bind("<Control-a>",
                  lambda e: (log_text.tag_add(tk.SEL, "1.0", tk.END), "break"))  # Ctrl+A для выделения всего

    # Информация о возможностях лога
    log_hint = tk.Label(log_frame,
                        text="Правый клик - меню. Ctrl+C - копировать выделенный текст",
                        font=("Arial", 8),
                        fg="gray")
    log_hint.pack(side=tk.BOTTOM, anchor="w", padx=5, pady=(0, 5))

    # Привязываем обработчики
    select_btn.config(command=lambda: select_files_callback(file_label, process_btn, log_text))
    process_btn.config(command=lambda: process_files_callback(log_text, progress_label))

    # Настраиваем фокус по умолчанию
    root.after(100, lambda: search_entry.focus_set())

    return root

# -------------------------------------------------------------------
# ОСНОВНЫЕ ФУНКЦИИ
# -------------------------------------------------------------------

def context_text(search_text, all_text):
    """Ваша существующая функция context_text"""
    position = all_text.find(search_text)
    if position == -1:
        return position
    else:
        context = all_text[position:position + 200]
        return context


def extract_protocol_number(all_text):
    """Извлекает номер протокола из текста"""
    try:
        if context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №', all_text) != -1:
            number_protocol = \
                context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №', all_text).split('№ ')[1].split('\n')[0].split(' от ')[0]
        else:
            number_protocol = \
                context_text('ПРОТОКОЛ ИССЛЕДОВАНИЙ №', all_text).split('№ ')[1].split('\n')[0].split(' от ')[0]
        return number_protocol.strip()
    except Exception as e:
        logger.error(f"Ошибка извлечения номера протокола: {e}")
        return None


def branch_pel(doc_text, doc_tables, data_var, df_columns_db, log_text=None):
    """Ваша существующая функция branch_pel"""
    if log_text:
        log_message(log_text, "Обработка протокола ПЭЛ...")

    all_text = doc_text
    all_tables = doc_tables
    data = data_var
    columns_db = df_columns_db

    # ------------------------------------------------------------------------------------------
    rejd_na_istochnik = '-'

    accreditaciya_doc = context_text('Уникальный номер записи об', all_text)
    if accreditaciya_doc != -1:
        accreditation = 'есть'
    else:
        accreditation = 'нет'

    number_protocol = context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №', all_text).split('№ ')[1].split('\n')[0].split(' от ')[0]
    date_protocol = context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №', all_text).split('№ ')[1].split('\n')[0].split(' от ')[1]

    # Дата и времена ПЭЛ
    data_izmereniya = context_text('Дата и время нача', all_text).split('\n\n')[1]
    time_start_izmereniya = context_text('Дата и время нача', all_text).split('\n\n')[4]
    time_end_izmereniya = context_text('Дата и время окончания изме', all_text).split('\n\n')[2]

    # type_protocol
    if accreditation == 'есть':
        pref_acc = '_uniq'
    else:
        pref_acc = '_dop'
    if context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №', all_text) != -1:
        type_prot = 'pel'
    else:
        type_prot = 'ai'
    type_protocol = f'{type_prot}{pref_acc}'

    # Читаем 1ю таблицу (ШАпка)
    df_structura = []
    for row in all_tables[1].rows:
        row_data = []
        for cell in row.cells:
            cell_text = cell.text.strip()
            row_data.append(cell_text)
        df_structura.append(row_data)

    df_hat = pd.DataFrame(df_structura)
    location = df_hat.loc[3, 1]

    # ------------------ Тут перебросимся в функцию парсинга АДРЕСА ------------
    okrug = parse_moscow_address(location).get('Округ')
    region = parse_moscow_address(location).get('Район')
    nazvanie_territorii = parse_moscow_address(location).get('Название территории')

    # Механизм записи переменных ШАПКИ в словарь для записи в БД
    local_vars = locals()
    for var_name in columns_db:
        if var_name in local_vars:
            data[var_name] = [local_vars[var_name]]

    # Читаем 3ю таблицу (Данные)
    df_structura = []
    for row in all_tables[3].rows:
        row_data = []
        for cell in row.cells:
            cell_text = cell.text.strip()
            row_data.append(cell_text)
        df_structura.append(row_data)

    df_data_tab3 = pd.DataFrame(df_structura)
    df_data_tab3 = df_data_tab3[[2, 3]]

    # Датафрейм с данными
    df_data_tab3.columns = df_data_tab3.iloc[0]
    df_data_tab3 = df_data_tab3[1:]

    # Находим строки с "Среднее значение" и дополняем их названием из предыдущей строки
    df = df_data_tab3.copy()
    mask = df['Наименование показателя'].str.contains('Среднее значение', na=False)
    previous_values = df['Наименование показателя'].shift(1)
    df.loc[mask, 'Наименование показателя'] = df.loc[mask, 'Наименование показателя'] + ' ' + previous_values

    df.loc[df['Наименование показателя'].str.contains('Среднее значение'), 'result_col'] = df['Наименование показателя']
    df_data_tab3_sh = df.loc[~df['result_col'].isna()].copy()

    df_data_tab3_sh['name_pokazatel'] = df_data_tab3_sh['result_col'].str.replace('Среднее значение ', '')

    # Просто ищем столбец, содержащий нужную подстроку 'Результат измерений ± \nпоказатель
    for col in df_data_tab3_sh.columns:
        if 'Результат измерений' in col:
            colonka_result_izmereniy = col
            break


    # df_data_tab3_sh = df_data_tab3_sh[['name_pokazatel', 'Результат измерений ± \nпоказатель точности, мг/м3']]
    df_data_tab3_sh = df_data_tab3_sh[['name_pokazatel', colonka_result_izmereniy]]

    # Удаляем все точки-запятые
    df_data_tab3_sh['name_pokazatel'] = df_data_tab3_sh['name_pokazatel'].str.replace(',', '')

    df_kirill = pd.read_excel('Columns_02_top.xlsx', sheet_name='База')

    # Получаем значение показателя
    df_data_tab3_sh = df_data_tab3_sh.merge(df_kirill, how='left', left_on='name_pokazatel', right_on='old_name')
    # df_data_tab3_sh = df_data_tab3_sh[['new_pokazat_name', 'Результат измерений ± \nпоказатель точности, мг/м3']]
    df_data_tab3_sh = df_data_tab3_sh[['new_pokazat_name', colonka_result_izmereniy]]

    # Грузим датафрейм с данными таблицы 3 в словарь data
    for index, row in df_data_tab3_sh.iterrows():
        pokazat_name = row['new_pokazat_name']
        # value = row['Результат измерений ± \nпоказатель точности, мг/м3']
        value = row[colonka_result_izmereniy]
        if pokazat_name in data:
            data[pokazat_name] = [value]

    if log_text:
        log_message(log_text, "Обработка ПЭЛ завершена")

    return data


def branch_ai(doc_text, doc_tables, data_var, df_columns_db, log_text=None):
    """Ваша существующая функция branch_ai"""
    if log_text:
        log_message(log_text, "Обработка протокола АИ...")

    all_text = doc_text
    all_tables = doc_tables
    data = data_var
    columns_db = df_columns_db

    # ------------------------------------------------------------------------------------------
    rejd_na_istochnik = '-'

    accreditaciya_doc = context_text('Уникальный номер записи об', all_text)
    if accreditaciya_doc != -1:
        accreditation = 'есть'
    else:
        accreditation = 'нет'

    number_protocol = context_text('ПРОТОКОЛ ИССЛЕДОВАНИЙ №', all_text).split('№ ')[1].split('\n')[0].split(' от ')[0]
    date_protocol = context_text('ПРОТОКОЛ ИССЛЕДОВАНИЙ №', all_text).split('№ ')[1].split('\n')[0].split(' от ')[1]

    # Дата и времена АИ
    data_izmereniya = context_text('Дата и время нача', all_text).split('\n\n')[1]
    time_start_izmereniya = context_text('Дата и время нача', all_text).split('\n\n')[4]
    time_end_izmereniya = context_text('Дата и время окончания', all_text).split('\n\n')[2]

    # type_protocol
    if accreditation == 'есть':
        pref_acc = '_uniq'
    else:
        pref_acc = '_dop'
    if context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №', all_text) != -1:
        type_prot = 'pel'
    else:
        type_prot = 'ai'
    type_protocol = f'{type_prot}{pref_acc}'

    # Читаем 1ю таблицу (ШАпка)
    df_structura = []
    for row in all_tables[1].rows:
        row_data = []
        for cell in row.cells:
            cell_text = cell.text.strip()
            row_data.append(cell_text)
        df_structura.append(row_data)

    df_hat = pd.DataFrame(df_structura)
    location = df_hat.loc[3, 1]
    # ------------------ Тут перебросимся в функцию парсинга АДРЕСА ------------
    # okrug = location
    # region = location
    # nazvanie_territorii = location

    okrug = parse_moscow_address(location).get('Округ')
    region = parse_moscow_address(location).get('Район')
    nazvanie_territorii = parse_moscow_address(location).get('Название территории')


    # Механизм записи переменных ШАПКИ в словарь для записи в БД
    local_vars = locals()
    for var_name in columns_db:
        if var_name in local_vars:
            data[var_name] = [local_vars[var_name]]

    # Читаем 3ю таблицу (Данные)
    df_structura = []
    for row in all_tables[3].rows:
        row_data = []
        for cell in row.cells:
            cell_text = cell.text.strip()
            row_data.append(cell_text)
        df_structura.append(row_data)

    df_data_tab3 = pd.DataFrame(df_structura)
    df_data_tab3 = df_data_tab3.iloc[:df_data_tab3[0].str.contains('Примечание:', na=False).idxmax()]
    df_data_tab3 = df_data_tab3[[2, 3]]

    df_data_tab3.columns = df_data_tab3.iloc[0]
    df_data_tab3 = df_data_tab3[1:]

    df_data_tab3_sh = df_data_tab3.reset_index()

    # Просто ищем столбец, содержащий нужную подстроку 'Результат измерений ± \nпоказатель
    for col in df_data_tab3_sh.columns:
        if 'Результат измерений' in col:
            colonka_result_izmereniy = col
            break

    df_data_tab3_ai = df_data_tab3_sh.rename(columns={'Наименование показателя': 'name_pokazatel'})
    # df_data_tab3_ai = df_data_tab3_ai[['name_pokazatel', 'Результат измерений ±\n показатель точности, мг/м3']]
    df_data_tab3_ai = df_data_tab3_ai[['name_pokazatel', colonka_result_izmereniy]]

    # Удаляем все точки-запятые
    df_data_tab3_ai['name_pokazatel'] = df_data_tab3_ai['name_pokazatel'].str.replace(',', '')

    df_kirill = pd.read_excel('Columns_02_top.xlsx', sheet_name='База')

    df_data_tab3_ai_clean = df_data_tab3_ai.merge(df_kirill, how='left', left_on='name_pokazatel', right_on='old_name')
    df_data_tab3_ai_clean = df_data_tab3_ai_clean[
        # ['new_pokazat_name', 'Результат измерений ±\n показатель точности, мг/м3']]
        ['new_pokazat_name', colonka_result_izmereniy]]

    # Грузим датафрейм с данными таблицы 3 в словарь data
    for index, row in df_data_tab3_ai_clean.iterrows():
        pokazat_name = row['new_pokazat_name']
        # value = row['Результат измерений ±\n показатель точности, мг/м3']
        value = row[colonka_result_izmereniy]
        if pokazat_name in data:
            data[pokazat_name] = [value]

    if log_text:
        log_message(log_text, "Обработка АИ завершена")

    return data


def parse_moscow_address(address_str):
    """
    Разбирает московский адрес на компоненты:
    - Округ (ВАО, ЗАО, САО, и т.д.)
    - Район (если указан)
    - Название территории (улица/переулок/шоссе + номер дома)
    """
    result = {
        'Округ': '-',
        'Район': '-',
        'Название территории': '-'
    }

    if not address_str or not isinstance(address_str, str):
        return result

    # Чистим строку от лишних пробелов
    address_str = ' '.join(address_str.split())

    # Убираем "г. Москва" в начале, если есть
    address_str = address_str.replace('г. Москва,', '').replace('г. Москва', '').strip()
    if address_str.startswith(','):
        address_str = address_str[1:].strip()

    # Разбиваем на части по запятым
    parts = [part.strip() for part in address_str.split(',')]

    # Список всех возможных аббревиатур округов Москвы
    okrug_abbr = {
        'ВАО', 'ЗАО', 'САО', 'СЗАО', 'СВАО', 'ЮАО', 'ЮВАО', 'ЮЗАО',
        'ЗелАО', 'ТиНАО', 'НАО', 'ЦАО'
    }

    # Словари для идентификации частей адреса
    street_types = {'ул.', 'ул', 'улица', 'улица.', 'просп.', 'проспект', 'пр.', 'пр-кт',
                    'ш.', 'шоссе', 'пер.', 'переулок', 'б-р', 'бульвар', 'бул.',
                    'наб.', 'набережная', 'аллея', 'ал.', 'пл.', 'площадь',
                    'проезд',
                    'мкр.', 'микрорайон', 'пос.', 'поселок', 'тер.', 'территория',
                    'кв-л', 'квартал', 'с-р', 'северный', 'дор.', 'дорога'}

    house_prefixes = {'д.', 'дом', 'д', 'к.', 'корпус', 'корп.', 'стр.', 'строение',
                      'соор.', 'сооружение', 'вл.', 'владение', 'лит.', 'литера'}

    # Шаг 1: Найти округ
    okrug_found = None
    okrug_index = -1

    for i, part in enumerate(parts):
        # Проверяем аббревиатуры округов
        if part in okrug_abbr:
            okrug_found = part
            okrug_index = i
            result['Округ'] = part
            break

        # Проверяем полные названия округов
        if ('Восточный' in part or 'ВАО' in part) and 'административный' in part:
            result['Округ'] = 'ВАО'
            okrug_index = i
            break
        elif ('Западный' in part or 'ЗАО' in part) and 'административный' in part and 'Зеленоградский' not in part:
            result['Округ'] = 'ЗАО'
            okrug_index = i
            break
        elif ('Северный' in part or 'САО' in part) and 'административный' in part:
            result['Округ'] = 'САО'
            okrug_index = i
            break
        elif 'ТиНАО' in part or 'Троицкий' in part or 'Новомосковский' in part:
            result['Округ'] = 'ТиНАО'
            okrug_index = i
            break
        elif 'Зеленоградский' in part or 'ЗелАО' in part:
            result['Округ'] = 'ЗелАО'
            okrug_index = i
            break

    # Шаг 2: Если нашли округ, ищем район
    if okrug_index != -1 and okrug_index + 1 < len(parts):
        # Проверяем следующие части после округа
        for i in range(okrug_index + 1, min(okrug_index + 3, len(parts))):
            part = parts[i]

            # Пропускаем очевидные улицы/дома
            if any(part.startswith(prefix) for prefix in street_types.union(house_prefixes)):
                continue

            # Проверяем, является ли эта часть районом
            # Район обычно:
            # 1) Содержит слово "район"
            # 2) ИЛИ это многословное название без типовых префиксов
            # 3) ИЛИ это название территории

            # Если содержит "район" или "поселение"
            if 'район' in part.lower() or 'поселение' in part.lower():
                # Извлекаем название района (убираем слово "район")
                district_name = part
                for word in ['район', 'поселение', 'муниципальный']:
                    district_name = district_name.replace(word, '').replace('  ', ' ').strip()
                result['Район'] = district_name
                break

            # Если это не улица и не дом, возможно это район
            elif not any(word in part.lower() for word in street_types.union({'д.', 'дом', 'к.', 'стр.'})):
                # Проверяем по характерным окончаниям или структуре
                # Районы часто заканчиваются на -ский, -ово, -ино и т.д.
                if part.endswith(('ский', 'вое', 'вое', 'во', 'ино', 'ово', 'ево')):
                    result['Район'] = part
                    break
                # Или если это составное название
                elif ' ' in part and not any(x in part.lower() for x in ['ул.', 'пер.', 'ш.', 'просп.']):
                    result['Район'] = part
                    break

    # Шаг 3: Находим "Название территории" (улица + номер дома)
    # Ищем начало адресной части (улица/переулок/шоссе)
    territory_parts = []

    # Если нашли округ, начинаем поиск после него (и после района, если есть)
    start_idx = okrug_index + 1 if okrug_index != -1 else 0
    if result['Район'] != '-':
        # Находим индекс района в частях
        for i in range(start_idx, len(parts)):
            if result['Район'] in parts[i]:
                start_idx = i + 1
                break

    # Собираем все части, начиная с первой улицы/переулка/шоссе
    for i in range(start_idx, len(parts)):
        part = parts[i]

        # Добавляем часть, если:
        # 1) Это улица/переулок/шоссе ИЛИ
        # 2) Это уже начало адресной части ИЛИ
        # 3) Это номер дома/корпуса

        if any(prefix in part.lower() for prefix in street_types):
            territory_parts.append(part)
        elif part and (territory_parts or any(word in part.lower() for word in house_prefixes)):
            territory_parts.append(part)
        elif i == start_idx and not any(word in part.lower() for word in ['район', 'поселение']):
            # Если с этого начинается адрес (нет явного указания типа улицы)
            territory_parts.append(part)

    # Формируем итоговую строку
    if territory_parts:
        result['Название территории'] = ', '.join(territory_parts)

    return result


def process_protocol(file_path, log_text=None):
    """Основная логика обработки протокола"""
    global engine, columns_db

    try:
        if log_text:
            log_message(log_text, "Загрузка документа...")

        tab_doc = Document(file_path)
        text_doc = docx2python(file_path)

        all_text = text_doc.text
        all_tables = tab_doc.tables

        if log_text:
            log_message(log_text, "Документ загружен, извлечение данных...")

        # Извлекаем номер протокола для проверки
        number_protocol = extract_protocol_number(all_text)
        if not number_protocol:
            error_msg = f"Не удалось извлечь номер протокола из файла {Path(file_path).name}"
            if log_text:
                log_message(log_text, f"ОШИБКА: {error_msg}")
            logger.error(error_msg)
            return "error"

        # Проверяем, существует ли протокол в БД
        if check_protocol_exists(number_protocol):
            warning_msg = f"Протокол №{number_protocol} уже существует в БД. Файл пропущен."
            if log_text:
                log_message(log_text, f"ПРЕДУПРЕЖДЕНИЕ: {warning_msg}")
            logger.warning(warning_msg)
            return "skipped"

        if log_text:
            log_message(log_text, f"Номер протокола: {number_protocol} (новый)")

        # Инициализация данных
        data = {col: [None] for col in columns_db}

        # Определение типа протокола и обработка
        if context_text('ПРОТОКОЛ ИЗМЕРЕНИЙ №', all_text) != -1:
            data = branch_pel(all_text, all_tables, data, columns_db, log_text)
        else:
            data = branch_ai(all_text, all_tables, data, columns_db, log_text)

        if log_text:
            log_message(log_text, "Сохранение в БД...")

        # Сохранение в БД
        # Сохранение с защитой от сбоев
        pd.DataFrame(data).to_sql(
                'protocols',
            con=engine,
            schema='public',  # схема в PostgreSQL
            if_exists='append',
            index=False
        )

        success_msg = f"Протокол №{number_protocol} успешно сохранен в БД"
        if log_text:
            log_message(log_text, success_msg)
        logger.info(success_msg)

        return "success"

    except Exception as e:
        error_msg = f"Ошибка при обработке файла {Path(file_path).name}: {str(e)}"
        if log_text:
            log_message(log_text, f"ОШИБКА: {error_msg}")
        logger.error(error_msg)
        return "error"


def main():
    """Главная функция"""
    # Инициализация приложения
    init_app()

    # Создание GUI
    root = create_gui()

    # Запуск главного цикла
    root.mainloop()


if __name__ == "__main__":
    main()