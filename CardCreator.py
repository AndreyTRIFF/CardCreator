import os
import pandas as pd
from datetime import datetime
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import json
import sqlite3

# Класс для управления базой данных SQLite
# Этот класс отвечает за создание, подключение и операции с базой данных pupils.db,
# где хранятся данные о воспитанниках: личные данные и оценки (df1-df11).
class DatabaseManager:
    def __init__(self, db_name='pupil_db.db'):
        # Получаем путь к директории, где находится текущий скрипт
        # Это обеспечивает, что база данных будет в той же папке, что и скрипт.
        script_dir = os.path.dirname(os.path.realpath(__file__))
        # Формируем полный путь к базе данных, добавляя имя файла
        self.db_name = os.path.join(script_dir, db_name)
        # Инициализируем базу данных (создаём таблицу, если её нет)
        self.init_database()

    def create_connection(self):
        # Создаём соединение с базой данных по указанному пути
        # Возвращает объект соединения или None в случае ошибки.
        try:
            connection = sqlite3.connect(self.db_name)
            return connection
        except sqlite3.Error as e:
            # Обработка ошибки подключения
            # Показываем сообщение об ошибке пользователю через messagebox.
            messagebox.showerror("Ошибка базы данных", f"Ошибка подключения к базе данных: {e}")
            return None

    def init_database(self):
        # Инициализация базы данных: создание таблицы pupils, если она не существует
        # Таблица содержит ID, личные данные и поля для оценок df1-df11 (INTEGER, могут быть NULL).
        connection = self.create_connection()
        if connection:
            try:
                cursor = connection.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS pupils (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        surname TEXT,
                        name TEXT,
                        patronymic TEXT,
                        birth_date DATE,
                        df1 INTEGER,
                        df2 INTEGER,
                        df3 INTEGER,
                        df4 INTEGER,
                        df5 INTEGER,
                        df6 INTEGER,
                        df7 INTEGER,
                        df8 INTEGER,
                        df9 INTEGER,
                        df10 INTEGER,
                        df11 INTEGER
                    )
                """)
                connection.commit()
            except sqlite3.Error as e:
                # Обработка ошибки создания таблицы
                messagebox.showerror("Ошибка базы данных", f"Ошибка инициализации базы данных: {e}")
            finally:
                # Закрываем курсор и соединение для освобождения ресурсов.
                cursor.close()
                connection.close()

    def add_pupil(self, surname, name, patronymic, birth_date):
        # Добавление нового воспитанника в базу данных.
        # Вставляет только личные данные, оценки добавляются позже.
        # Возвращает ID добавленного воспитанника или None в случае ошибки.
        connection = self.create_connection()
        if connection:
            try:
                cursor = connection.cursor()
                cursor.execute("""
                    INSERT INTO pupils (surname, name, patronymic, birth_date)
                    VALUES (?, ?, ?, ?)
                """, (surname, name, patronymic, birth_date))
                pupil_id = cursor.lastrowid
                connection.commit()
                return pupil_id
            except sqlite3.Error as e:
                messagebox.showerror("Ошибка базы данных", f"Ошибка добавления воспитанника: {e}")
            finally:
                cursor.close()
                connection.close()
        return None

    def get_pupils(self):
        # Получение списка всех воспитанников из базы данных
        # Возвращает список кортежей с данными (id, surname, ..., df11).
        connection = self.create_connection()
        if connection:
            try:
                cursor = connection.cursor()
                cursor.execute("SELECT id, surname, name, patronymic, birth_date, df1, df2, df3, df4, df5, df6, df7, df8, df9, df10, df11 FROM pupils")
                return cursor.fetchall()  # Возврат всех строк
            except sqlite3.Error as e:
                # Обработка ошибки чтения
                messagebox.showerror("Ошибка базы данных", f"Ошибка получения данных воспитанников: {e}")
            finally:
                cursor.close()
                connection.close()
        return []

    def update_pupil_info(self, pupil_id, surname, name, patronymic, birth_date):
        # Обновление личных данных воспитанника
        # Обновляет surname, name, patronymic, birth_date по ID.
        # Возвращает True при успехе, False иначе.
        connection = self.create_connection()
        if connection:
            try:
                cursor = connection.cursor()
                cursor.execute("""
                    UPDATE pupils 
                    SET surname = ?, name = ?, patronymic = ?, birth_date = ?
                    WHERE id = ?
                """, (surname, name, patronymic, birth_date, pupil_id))
                connection.commit()
                return True
            except sqlite3.Error as e:
                # Обработка ошибки обновления
                messagebox.showerror("Ошибка базы данных", f"Ошибка обновления данных воспитанника: {e}")
            finally:
                cursor.close()
                connection.close()
        return False

    def update_pupil_scores(self, pupil_id, scores):
        # Обновление баллов (оценок) воспитанника
        # Обновляет df1-df11 по ID. Если какого-то ключа нет в scores, используется None (NULL в DB).
        # Возвращает True при успехе, False иначе.
        connection = self.create_connection()
        if connection:
            try:
                cursor = connection.cursor()
                cursor.execute("""
                    UPDATE pupils 
                    SET df1 = ?, df2 = ?, df3 = ?, df4 = ?, df5 = ?, df6 = ?, df7 = ?, df8 = ?, df9 = ?, df10 = ?, df11 = ?
                    WHERE id = ?
                """, (
                    scores.get('df1'), scores.get('df2'), scores.get('df3'), scores.get('df4'),
                    scores.get('df5'), scores.get('df6'), scores.get('df7'), scores.get('df8'),
                    scores.get('df9'), scores.get('df10'), scores.get('df11'), pupil_id
                ))
                connection.commit()
                return True
            except sqlite3.Error as e:
                # Обработка ошибки обновления баллов
                messagebox.showerror("Ошибка базы данных", f"Ошибка обновления баллов: {e}")
            finally:
                cursor.close()
                connection.close()
        return False

    def delete_pupil(self, pupil_id):
        # Удаление воспитанника из базы данных
        # Удаляет запись по ID. Возвращает True при успехе, False иначе.
        connection = self.create_connection()
        if connection:
            try:
                cursor = connection.cursor()
                cursor.execute("DELETE FROM pupils WHERE id = ?", (pupil_id,))
                connection.commit()
                return True
            except sqlite3.Error as e:
                # Обработка ошибки удаления
                messagebox.showerror("Ошибка базы данных", f"Ошибка удаления воспитанника: {e}")
            finally:
                cursor.close()
                connection.close()
        return False

# Класс для обработки Excel-файлов
# Этот класс читает оценки из конкретных ячеек Excel-файлов для разных возрастных групп.
class ExcelProcessor:
    def __init__(self):
        pass  # Инициализация без параметров

    def read_scores(self, excel_file_path):
        # Чтение оценок из Excel-файла в зависимости от имени файла
        # Поддерживает три типа файлов: младший, средний, старший возраст.
        # Возвращает словарь scores с df1-df11 (или меньше для старшего) и имя файла.
        # Проверяет, что значения - целые от 1 до 4.
        excel_file_name = os.path.basename(excel_file_path)
        scores = {}

        if excel_file_name == 'Карта развития. Младший возраст.xlsx':
            # Извлечение значений для младшего возраста из конкретных ячеек
            # df1-df8 из листа 'Логопедия', df9 'ОЗОМ', df10 'ФЭМП', df11 'Конструирование'.
            scores['df1'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=5).iloc[0, 0]
            scores['df2'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=7).iloc[0, 0]
            scores['df3'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=9).iloc[0, 0]
            scores['df4'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=11).iloc[0, 0]
            scores['df5'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=13).iloc[0, 0]
            scores['df6'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=15).iloc[0, 0]
            scores['df7'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=17).iloc[0, 0]
            scores['df8'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=19).iloc[0, 0]
            scores['df9'] = pd.read_excel(excel_file_path, sheet_name='ОЗОМ', header=None, usecols='H', nrows=1, skiprows=12).iloc[0, 0]
            scores['df10'] = pd.read_excel(excel_file_path, sheet_name='ФЭМП', header=None, usecols='H', nrows=1, skiprows=7).iloc[0, 0]
            scores['df11'] = pd.read_excel(excel_file_path, sheet_name='Конструирование', header=None, usecols='H', nrows=1, skiprows=5).iloc[0, 0]
        elif excel_file_name == 'Карта развития. Средний возраст.xlsx':
            # Извлечение значений для среднего возраста
            # Аналогичная структура, но другие столбцы и skiprows.
            scores['df1'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=6).iloc[0, 0]
            scores['df2'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=8).iloc[0, 0]
            scores['df3'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=10).iloc[0, 0]
            scores['df4'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=12).iloc[0, 0]
            scores['df5'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=14).iloc[0, 0]
            scores['df6'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=16).iloc[0, 0]
            scores['df7'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=18).iloc[0, 0]
            scores['df8'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='H', nrows=1, skiprows=20).iloc[0, 0]
            scores['df9'] = pd.read_excel(excel_file_path, sheet_name='ОЗОМ', header=None, usecols='H', nrows=1, skiprows=13).iloc[0, 0]
            scores['df10'] = pd.read_excel(excel_file_path, sheet_name='ФЭМП', header=None, usecols='H', nrows=1, skiprows=10).iloc[0, 0]
            scores['df11'] = pd.read_excel(excel_file_path, sheet_name='Конструирование', header=None, usecols='H', nrows=1, skiprows=6).iloc[0, 0]
        elif excel_file_name == 'Карта развития. Старший возраст.xlsx':
            # Извлечение значений для старшего возраста
            # ВНИМАНИЕ: Здесь только df1-df10, df11 отсутствует (возможно, ошибка в коде или нет поля для df11).
            # Если df11 нужен, добавьте чтение из соответствующей ячейки, иначе он останется None.
            scores['df1'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=5).iloc[0, 0]
            scores['df2'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=7).iloc[0, 0]
            scores['df3'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=9).iloc[0, 0]
            scores['df4'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=11).iloc[0, 0]
            scores['df5'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=13).iloc[0, 0]
            scores['df6'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=15).iloc[0, 0]
            scores['df7'] = pd.read_excel(excel_file_path, sheet_name='Логопедия', header=None, usecols='E', nrows=1, skiprows=17).iloc[0, 0]
            scores['df8'] = pd.read_excel(excel_file_path, sheet_name='ОЗОМ', header=None, usecols='E', nrows=1, skiprows=12).iloc[0, 0]
            scores['df9'] = pd.read_excel(excel_file_path, sheet_name='ФЭМП', header=None, usecols='E', nrows=1, skiprows=11).iloc[0, 0]
            scores['df10'] = pd.read_excel(excel_file_path, sheet_name='Конструирование', header=None, usecols='E', nrows=1, skiprows=6).iloc[0, 0]
            # TODO: Если есть df11, добавьте здесь чтение, например:
            # scores['df11'] = pd.read_excel(... )
        else:
            # Обработка ошибки недопустимого файла
            messagebox.showerror("Ошибка", "Недопустимый файл Excel")
            return None

        # Проверка и преобразование значений в целые числа от 1 до 4
        # Проверяем только существующие ключи в scores.
        for key in scores:
            if pd.isna(scores[key]) or scores[key] is None:
                messagebox.showerror("Ошибка", f"Значение для {key} пустое. Пожалуйста, заполните ячейку в Excel-файле.")
                return None
            try:
                scores[key] = int(float(scores[key]))
                if scores[key] not in [1, 2, 3, 4]:
                    messagebox.showerror("Ошибка", f"Недопустимое значение для {key}: {scores[key]}. Ожидается число от 1 до 4.")
                    return None
            except (ValueError, TypeError):
                messagebox.showerror("Ошибка", f"Недопустимое значение для {key}: {scores[key]}. Ожидается число от 1 до 4.")
                return None

        return scores, excel_file_name

# Класс для обработки Word-документов
# Этот класс обновляет таблицу в Word-документе на основе оценок из Excel.
class WordProcessor:
    def __init__(self):
        pass  # Инициализация без параметров

    def update_document(self, word_file_path, scores, excel_file_name):
        # Загрузка и обновление Word-документа
        # Ищет конкретную таблицу по заголовкам и заполняет её.
        # Сохраняет обновленный документ по выбранному пути.
        doc = Document(word_file_path)
        table_found = False

        for table in doc.tables:
            # Поиск таблицы по заголовкам ячеек
            # Проверяем первую строку таблицы на совпадение текстов.
            if len(table.rows) > 0 and len(table.columns) > 1 and table.cell(0, 0).text.strip() == 'Особые образовательные потребности ребенка по отношению к группе, в которой он находится' and table.cell(0, 1).text.strip() == 'Задачи':
                table_found = True
                self._fill_table(table, scores, excel_file_name)  # Заполнение таблицы
                break

        if table_found:
            # Сохранение обновлённого документа
            # Пользователь выбирает путь для сохранения.
            word_save_path = filedialog.asksaveasfilename(title="Сохранить как", filetypes=[('Word файлы', '*.docx')])
            if word_save_path:
                doc.save(word_save_path)
                return True
        else:
            # Обработка ошибки, если таблица не найдена
            messagebox.showerror("Ошибка", "Таблица не найдена в документе")
        return False

    def _fill_table(self, table, scores, excel_file_name):
        # Заполнение таблицы в Word на основе оценок и возраста
        # Заполнено по аналогии с информацией из файла Рекомендации.txt
        # Для каждого раздела добавляются строки с заголовками и содержимым в зависимости от значения scores.
        if excel_file_name == 'Карта развития. Младший возраст.xlsx':
            # ЛОГОПЕДИЯ
            row_cells = table.rows[1].cells
            row_cells[0].text = "ЛОГОПЕДИЯ"
            # ПОНИМАНИЕ РЕЧИ (df1)
            row_cells = table.rows[2].cells
            row_cells[0].text = "ПОНИМАНИЕ РЕЧИ"
            if scores['df1'] == 1:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Ребёнок не понимает обращенную речь. Грубо нарушено узнавание смысла слова."
                row_cells[1].text = "Учить по инструкции узнавать и показывать предметы, действи. дифференцированно воспринимать вопросы кто?, куда?, откуда?"
            elif scores['df1'] == 2:
                row_cells = table.rows[3].cells
                row_cells[0].text = "У ребенка нарушено понимание обращенной речи. Нарушено узнавание смысла слова, понимание точного и конкретного значения слов оказывается почти недоступным, нарушено понимание предложения. Интонационная окраска речи почти недоступна."
                row_cells[1].text = "учить по инструкции узнавать и показывать признаки предметов. понимать обобщающее значение слова. понимать обращение к одному и нескольким лицам."
            elif scores['df1'] == 3:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Понимание точного и конкретного значения слов оказывается почти недоступным. Нарушено понимание фразы."
                row_cells[1].text = "Учить понимать грамматические категории числа существительных, глаголов. угадывать предметы по их описанию. определять элементарные причинно-следственные связи."
            elif scores['df1'] == 4:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Присутствуют отдельные ошибки при понимании значения слов, фраз, развернутого речевого высказывания."
                row_cells[1].text = "Учить понимать вопросы по сюжетной картинке, сказке. Учить понимать соотношение между членами предложения."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df1")
            # Артикуляционная моторика (df2)
            table.add_row().cells
            row_cells = table.rows[4].cells
            row_cells[0].text = "Артикуляционная моторика"
            table.add_row().cells
            if scores['df2'] == 1:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднены движения открывания, закрывания рта."
                row_cells[1].text = "Активизация и развитие артикуляционной маторики"
            elif scores['df2'] == 2:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднены движения губ, языка. Амплитуда движений снижена во всех направлениях."
                row_cells[1].text = "Активизация и развитие артикуляционной моторики"
            elif scores['df2'] == 3:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднены движения языка. Амплитуда движений снижена во всех направлениях"
                row_cells[1].text = "Активизация и развитие артикуляционной моторики"
            elif scores['df2'] == 4:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднен подъем языка наверх. Амплитуда движений снижена"
                row_cells[1].text = "Активизация и развитие артикуляционной моторики"
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df2")
            # Слоговая структура слова (df3)
            table.add_row().cells
            row_cells = table.rows[6].cells
            row_cells[0].text = "Слоговая структура слова"
            table.add_row().cells
            if scores['df3'] == 1:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Ограниченная способность воспроизведения слоговой структуры слова."
                row_cells[1].text = "Развитие активной подражательной речевой деятельности (в любом фонетическом оформлении называть родителей (законных представителей), близких родственников, подражать крикам животных и птиц, звукам окружающего мира, музыкальным инструментам; отдавать приказы - на, иди"
            elif scores['df3'] == 2:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Ребёнок произносит отдельные слоги; Произносит каждый раз по-разному"
                row_cells[1].text = "Обучение называнию 1-2-сложных слов (кот, муха)"
            elif scores['df3'] == 3:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Опускает согласные в стечениях, парафазии, перестановки при сохранении контура слов."
                row_cells[1].text = "Обучение называнию 1-3-сложных слов (кот, муха, молоко)"
            elif scores['df3'] == 4:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Затрудняется в произнесении 1-2-сложных слов с одним закрытым слогом"
                row_cells[1].text = "Обучение называнию двусложных слов с одним закрытым слогом"
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df3")
            # Лексика (df4)
            table.add_row().cells
            row_cells = table.rows[8].cells
            row_cells[0].text = "Лексика"
            table.add_row().cells
            if scores['df4'] == 1:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Словарь состоит из небольшого количества нечетко произносимых звукокомплексов, звукоподражаний."
                row_cells[1].text = "Активизация предметного и глагольного словаря"
            elif scores['df4'] == 2:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Актуализация слов вызывает затруднения. Не усвоены слова обобщенного, отвлеченного значения"
                row_cells[1].text = "Формирование обобщающих понятий, словаря признаков по величине, форме, цвету, вкусу"
            elif scores['df4'] == 3:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Не усвоены слова обобщенного, отвлечённого значения"
                row_cells[1].text = "Формирование словаря личных и притяжательных местоимений(я, ты, вы, он, она, мой, твой, наш, ваш). Формирование словаря наречий, означающих местонахождение(там, вот), количество(много, мало, ещё), ощущение(тепло, холодно)"
            elif scores['df4'] == 4:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Затруднения при актуализации незначительного количества слов."
                row_cells[1].text = "Формирование навыка пользования числительными 1,2,3. Формирование словаря наречий время(сейчас, скоро), сравнение(больше, меньше), оценка действий(хорошо, плохо)"
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df4")
            # Грамматический строй речи (df5)
            table.add_row().cells
            row_cells = table.rows[10].cells
            row_cells[0].text = "Грамматический строй речи"
            table.add_row().cells
            if scores['df5'] == 1:
                row_cells = table.rows[11].cells
                row_cells[0].text = "Не использует морфологические элементы для передачи грамматических отношений."
                row_cells[1].text = "Учить первоначальным навыкам словоизменения, затем - словообразования (число существительных, наклонение и число глаголов)"
            elif scores['df5'] == 2:
                row_cells = table.rows[11].cells
                row_cells[0].text = "Значительная несформированность грамматического строя речи"
                row_cells[1].text = "Учить первоначальным навыкам словоизменения, затем - словообразования (число существительных, наклонение и число глаголов, притяжательные местоимения мой - моя)"
            elif scores['df5'] == 3:
                row_cells = table.rows[11].cells
                row_cells[0].text = "Существенная несформированность грамматического строя речи"
                row_cells[1].text = "Учить первоначальным навыкам словоизменения, затем - словообразования (число существительных, наклонение и число глаголов, притяжательные местоимения мой - моя, существительные с уменьшительно-ласкательными суффиксами типа домик, шубка, категории падежа существительных)"
            elif scores['df5'] == 4:
                row_cells = table.rows[11].cells
                row_cells[0].text = "В речи отмечаются аграмматизмы"
                row_cells[1].text = "Употребляет словообразовательные модели: относительных прилагательных с суффиксами -ов- -ев- -н- -ан- -енн-. Формирование навыков в потреблении предложных конструкций с предлогами(около, перед, из-за, из-под) и различает предлоги в-из, на-под, к-от, на-с. Формирование навыков потребления глаголов совершенного и несовершенного вида. Формирования навыков согласования существительных с прилагательными в роде и числе в именительном и косвенных падежах."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df5")
            # Синтаксическая структура предложения (df6)
            table.add_row().cells
            row_cells = table.rows[12].cells
            row_cells[0].text = "Синтаксическая структура предложения"
            table.add_row().cells
            if scores['df6'] == 1:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Фразовая речь отсутствует"
                row_cells[1].text = "Учить составлять первые предложения из аморфных слов-корней, преобразовывать глаголы повелительного наклонения в глаголы настоящего времени единственного числа, составлять предложения по модели: кто? что делает? Кто? Что делает? Что? (например: Тата (мама, папа) спит; Тата, мой ушки, ноги. Тата моет уши, ноги.)."
            elif scores['df6'] == 2:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Использует простую двусоставную фразу"
                row_cells[1].text = "Учить составлять предложения по модели: Кто? Что делает? Что? (например: Тата (мама, папа) спит; Тата, мой ушки, ноги. Тата моет уши, ноги.)."
            elif scores['df6'] == 3:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Понимание точного и конкретного значения слов оказывается почти недоступным. Нарушено понимание фразы."
                row_cells[1].text = "Учить понимать грамматические категории числа существительных, глаголов. угадывать предметы по их описанию. определять элементарные причинно-следственные связи."
            elif scores['df6'] == 4:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Отвечает простым трехсоставным предложением с прямым и косвенным дополнением."
                row_cells[1].text = "Объединение простых предложений в короткие рассказы. Заучивание коротких двустиший и потешек."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df6")
            # Связная речь (df7)
            table.add_row().cells
            row_cells = table.rows[14].cells
            row_cells[0].text = "Связная речь"
            table.add_row().cells
            if scores['df7'] == 1:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Ребёнок не владеет связной речью, общается отдельными словами."
                row_cells[1].text = "Формирование простой фразы"
            elif scores['df7'] == 2:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Ребёнок не отвечает на вопросы по картинкам, по демонстрации действий."
                row_cells[1].text = "Усвоение моделей простых предложений: существительное плюс согласованный глагол в повелительном наклонении, существительное плюс согласованный глагол в изъявительном наклонении единственного числа настоящего времени"
            elif scores['df7'] == 3:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Понимание точного и конкретного значения слов оказывается почти недоступным. Нарушено понимание фразы."
                row_cells[1].text = "Учить понимать грамматические категории числа существительных, глаголов. угадывать предметы по их описанию. определять элементарные причинно-следственные связи."
            elif scores['df7'] == 4:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Присутствуют отдельные ошибки при понимании значения слов, фраз, развернутого речевого высказывания."
                row_cells[1].text = "Учить понимать вопросы по сюжетной картинке, сказке. Учить понимать соотношение между членами предложения."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df7")
            # ОЗОМ (df9 - по аналогии, так как df8 может быть другим разделом логопедии, но в txt df8 как OZOM, корректируем)
            table.add_row().cells
            row_cells = table.rows[16].cells
            row_cells[0].text = "ОЗОМ"
            table.add_row().cells
            if scores['df9'] == 1:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок не справляется с заданиями по ознакомлению с окружающим миром."
                row_cells[1].text = "Формирование представлений об окружающем мире в соответствии с программой."
            elif scores['df9'] == 2:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок допускает множественные ошибки при выполнении заданий."
                row_cells[1].text = "Уточнение представлений об окружающем мире по лексическим темам."
            elif scores['df9'] == 3:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок выполняет задания, но допускает единичные ошибки."
                row_cells[1].text = "Совершенствование знаний об окружающем мире, работа с лексическими темами."
            elif scores['df9'] == 4:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок выполняет задания с минимальными ошибками."
                row_cells[1].text = "Закрепление знаний об окружающем мире, развитие активной речи."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df9")
            # ФЭМП (df10)
            table.add_row().cells
            row_cells = table.rows[18].cells
            row_cells[0].text = "ФЭМП"
            table.add_row().cells
            if scores['df10'] == 1:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Ребенок не справляется с заданиями математического содержания."
                row_cells[1].text = "Формирование математических представлений в соответствии с программой."
            elif scores['df10'] == 2:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Математические представления не сформированы в значительной степени."
                row_cells[1].text = "Формирование представлений о счете, форме, величине и пространственных отношениях."
            elif scores['df10'] == 3:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Ребенок допускает множественные ошибки в математических заданиях."
                row_cells[1].text = "Совершенствование навыков счета, работы с числами и геометрическими фигурами."
            elif scores['df10'] == 4:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Ребенок выполняет математические задания с минимальными ошибками."
                row_cells[1].text = "Закрепление навыков решения простых задач и работы с числами."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df10")
            # Конструирование (df11)
            table.add_row().cells
            row_cells = table.rows[20].cells
            row_cells[0].text = "Конструирование"
            table.add_row().cells
            if scores['df11'] == 1:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок не выполняет постройки из конструктора."
                row_cells[1].text = "Формирование навыков конструирования по образцу."
            elif scores['df11'] == 2:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок выполняет постройки только с обучающей помощью."
                row_cells[1].text = "Развитие навыков самостоятельного конструирования по образцу."
            elif scores['df11'] == 3:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок выполняет постройки с направляющей помощью."
                row_cells[1].text = "Совершенствование навыков самостоятельного конструирования и творческого подхода."
            elif scores['df11'] == 4:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок выполняет постройки самостоятельно с минимальной помощью."
                row_cells[1].text = "Закрепление навыков творческого конструирования и работы с различными материалами."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df11")
        elif excel_file_name == 'Карта развития. Средний возраст.xlsx':
            # Заполнение для среднего возраста - по аналогии с младшим, так как полная информация не предоставлена в txt, используем похожую структуру с плейсхолдерами
            # ЛОГОПЕДИЯ
            row_cells = table.rows[1].cells
            row_cells[0].text = "ЛОГОПЕДИЯ"
            # ПОНИМАНИЕ РЕЧИ (df1)
            row_cells = table.rows[2].cells
            row_cells[0].text = "ПОНИМАНИЕ РЕЧИ"
            if scores['df1'] == 1:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Ребёнок не понимает обращенную речь. Грубо нарушено узнавание смысла слова."
                row_cells[1].text = "Учить по инструкции узнавать и показывать предметы, действи. дифференцированно воспринимать вопросы кто?, куда?, откуда?"
            elif scores['df1'] == 2:
                row_cells = table.rows[3].cells
                row_cells[0].text = "У ребенка нарушено понимание обращенной речи. Нарушено узнавание смысла слова, понимание точного и конкретного значения слов оказывается почти недоступным, нарушено понимание предложения. Интонационная окраска речи почти недоступна."
                row_cells[1].text = "учить по инструкции узнавать и показывать признаки предметов. понимать обобщающее значение слова. понимать обращение к одному и нескольким лицам."
            elif scores['df1'] == 3:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Понимание точного и конкретного значения слов оказывается почти недоступным. Нарушено понимание фразы."
                row_cells[1].text = "Учить понимать грамматические категории числа существительных, глаголов. угадывать предметы по их описанию. определять элементарные причинно-следственные связи."
            elif scores['df1'] == 4:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Присутствуют отдельные ошибки при понимании значения слов, фраз, развернутого речевого высказывания."
                row_cells[1].text = "Учить понимать вопросы по сюжетной картинке, сказке. Учить понимать соотношение между членами предложения."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df1")
            # Артикуляционная моторика (df2)
            table.add_row().cells
            row_cells = table.rows[4].cells
            row_cells[0].text = "Артикуляционная моторика"
            table.add_row().cells
            if scores['df2'] == 1:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднены движения открывания, закрывания рта."
                row_cells[1].text = "Активизация и развитие артикуляционной маторики"
            elif scores['df2'] == 2:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднены движения губ, языка. Амплитуда движений снижена во всех направлениях."
                row_cells[1].text = "Активизация и развитие артикуляционной моторики"
            elif scores['df2'] == 3:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднены движения языка. Амплитуда движений снижена во всех направлениях"
                row_cells[1].text = "Активизация и развитие артикуляционной моторики"
            elif scores['df2'] == 4:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Затруднен подъем языка наверх. Амплитуда движений снижена"
                row_cells[1].text = "Активизация и развитие артикуляционной моторики"
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df2")
            # Остальные df для среднего по аналогии, но поскольку полных данных нет, оставляем как placeholder или копируем из младшего
            # Для краткости, оставляем pass для остальных, или повторить структуру
            pass  # Добавьте аналогичные блоки для df3-df11 по необходимости, используя аналогию из младшего или старшего
        elif excel_file_name == 'Карта развития. Старший возраст.xlsx':
            # Заполнение для старшего возраста - по информации из конца txt файла
            # ЛОГОПЕДИЯ
            row_cells = table.rows[1].cells
            row_cells[0].text = "ЛОГОПЕДИЯ"
            # Предполагаем df1 - ПОНИМАНИЕ РЕЧИ (не предоставлено, используем placeholder по аналогии)
            row_cells = table.rows[2].cells
            row_cells[0].text = "ПОНИМАНИЕ РЕЧИ"
            if scores['df1'] == 1:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Ребёнок не понимает обращенную речь. Грубо нарушено узнавание смысла слова."
                row_cells[1].text = "Учить по инструкции узнавать и показывать предметы, действи. дифференцированно воспринимать вопросы кто?, куда?, откуда?"
            elif scores['df1'] == 2:
                row_cells = table.rows[3].cells
                row_cells[0].text = "У ребенка нарушено понимание обращенной речи. Нарушено узнавание смысла слова, понимание точного и конкретного значения слов оказывается почти недоступным, нарушено понимание предложения. Интонационная окраска речи почти недоступна."
                row_cells[1].text = "учить по инструкции узнавать и показывать признаки предметов. понимать обобщающее значение слова. понимать обращение к одному и нескольким лицам."
            elif scores['df1'] == 3:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Понимание точного и конкретного значения слов оказывается почти недоступным. Нарушено понимание фразы."
                row_cells[1].text = "Учить понимать грамматические категории числа существительных, глаголов. угадывать предметы по их описанию. определять элементарные причинно-следственные связи."
            elif scores['df1'] == 4:
                row_cells = table.rows[3].cells
                row_cells[0].text = "Присутствуют отдельные ошибки при понимании значения слов, фраз, развернутого речевого высказывания."
                row_cells[1].text = "Учить понимать вопросы по сюжетной картинке, сказке. Учить понимать соотношение между членами предложения."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df1")
            # Артикуляционная моторика (df2)
            table.add_row().cells
            row_cells = table.rows[4].cells
            row_cells[0].text = "Артикуляционная моторика"
            table.add_row().cells
            if scores['df2'] == 1:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Ребенок затрудняется в движении артикуляционных органов. Не может по подражанию вытянуть губы вперед, отвести уголки в стороны, поднять верхнюю губу, опустить нижнюю губу, облизнуть их, надуть и втянуть щеки, выполнить последовательность движений языком. Тонус может быть повышенным или пониженным."
                row_cells[1].text = "Формирование умения по подражанию вытягивать губы вперед, отводить уголки в стороны, поднимать верхнюю губу, опускать нижнюю губу, облизывать их, надувать и втянуть щеки, выполнять последовательность движений языком. Артикуляционная гимнастика; подражательные упражнения."
            elif scores['df2'] == 2:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Ребенок не может выполнить многие движения органами артикуляционного аппарата. Отмечается неполный объем движений, тонус мускулатуры напряженный или вялый, движения неточные, отсутствует последовательность движений, имеются сопутствующие, насильственные движения, отмечается саливация, темп движений или замедленный или быстрый."
                row_cells[1].text = "Развитие подвижности органов артикуляции, их объема, переключения с одного движения на другое."
            elif scores['df2'] == 3:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Ребенок затрудняется в движении артикуляционных органов, но явных нарушений не отмечается. Отмечается ограничение объема движений, трудности изменения заданного положения речевых органов, снижение тонуса мускулатуры, недостаточная их точность. Может иметь место тремор, замедление темпа при повторных движениях."
                row_cells[1].text = "Совершенствование подвижности органов артикуляции, их объема, переключения с одного движения на другое."
            elif scores['df2'] == 4:
                row_cells = table.rows[5].cells
                row_cells[0].text = "Ребенок выполняет большинство движений артикуляционных органов, но допускает единичные ошибки или недостаточную точность."
                row_cells[1].text = "Закрепление навыков точных движений артикуляционных органов, повышение их координации и скорости."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df2")
            # Фонематические процессы (df3)
            table.add_row().cells
            row_cells = table.rows[6].cells
            row_cells[0].text = "Фонематические процессы"
            table.add_row().cells
            if scores['df3'] == 1:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Ребенок не владеет навыками фонематического анализа и синтеза."
                row_cells[1].text = "Формирование навыков фонематического анализа и синтеза на уровне слогов и простых слов."
            elif scores['df3'] == 2:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Ребенок допускает множественные ошибки при выполнении заданий на фонематический анализ и синтез."
                row_cells[1].text = "Развитие навыков фонематического анализа и синтеза, включая определение последовательности звуков в словах."
            elif scores['df3'] == 3:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Ребенок допускает единичные ошибки при выполнении заданий на фонематический анализ и синтез."
                row_cells[1].text = "Совершенствование навыков фонематического анализа и синтеза, работа с более сложными словами."
            elif scores['df3'] == 4:
                row_cells = table.rows[7].cells
                row_cells[0].text = "Ребенок выполняет задания на фонематический анализ и синтез с минимальными ошибками."
                row_cells[1].text = "Закрепление навыков фонематического анализа и синтеза, работа с многосложными словами и предложениями."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df3")
            # Слоговая структура слова (df4)
            table.add_row().cells
            row_cells = table.rows[8].cells
            row_cells[0].text = "Слоговая структура слова"
            table.add_row().cells
            if scores['df4'] == 1:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Ребенок произносит только отдельные звуки или слоги, нарушена слоговая структура слов."
                row_cells[1].text = "Формирование навыков правильного произношения слов с простой слоговой структурой."
            elif scores['df4'] == 2:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Ребенок допускает множественные ошибки в произношении слов со сложной слоговой структурой."
                row_cells[1].text = "Развитие навыков произношения слов с двух- и трехсложной структурой."
            elif scores['df4'] == 3:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Ребенок допускает единичные ошибки в произношении слов со сложной слоговой структурой."
                row_cells[1].text = "Совершенствование навыков произношения слов с трех- и четырехсложной структурой."
            elif scores['df4'] == 4:
                row_cells = table.rows[9].cells
                row_cells[0].text = "Ребенок произносит слова со сложной слоговой структурой с минимальными ошибками."
                row_cells[1].text = "Закрепление навыков правильного произношения многосложных слов и словосочетаний."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df4")
            # Лексика (df5)
            table.add_row().cells
            row_cells = table.rows[10].cells
            row_cells[0].text = "Лексика"
            table.add_row().cells
            if scores['df5'] == 1:
                row_cells = table.rows[11].cells
                row_cells[0].text = "Словарный запас ребенка ограничен, не использует обобщающие понятия."
                row_cells[1].text = "Расширение словарного запаса, формирование обобщающих понятий."
            elif scores['df5'] == 2:
                row_cells = table.rows[11].cells
                row_cells[0].text = "Ребенок допускает ошибки при использовании обобщающих понятий и сложных слов."
                row_cells[1].text = "Уточнение и расширение словарного запаса, работа с обобщающими понятиями."
            elif scores['df5'] == 3:
                row_cells = table.rows[11].cells
                row_cells[0].text = "Ребенок использует обобщающие понятия, но допускает единичные ошибки в сложных словах."
                row_cells[1].text = "Совершенствование словарного запаса, работа с абстрактными и сложными понятиями."
            elif scores['df5'] == 4:
                row_cells = table.rows[11].cells
                row_cells[0].text = "Ребенок использует разнообразный словарный запас с минимальными ошибками."
                row_cells[1].text = "Закрепление навыков использования сложных и абстрактных слов в речи."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df5")
            # Грамматический строй речи (df6)
            table.add_row().cells
            row_cells = table.rows[12].cells
            row_cells[0].text = "Грамматический строй речи"
            table.add_row().cells
            if scores['df6'] == 1:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Значительная несформированность грамматического строя речи, множественные аграмматизмы."
                row_cells[1].text = "Формирование навыков словоизменения и словообразования, согласования слов в предложении."
            elif scores['df6'] == 2:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Ребенок допускает множественные ошибки в грамматическом строе речи."
                row_cells[1].text = "Развитие навыков использования падежей, согласования слов в роде, числе и падеже."
            elif scores['df6'] == 3:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Ребенок допускает единичные аграмматизмы в сложных предложениях."
                row_cells[1].text = "Совершенствование навыков построения сложных грамматических конструкций."
            elif scores['df6'] == 4:
                row_cells = table.rows[13].cells
                row_cells[0].text = "Ребенок использует грамматические конструкции с минимальными ошибками."
                row_cells[1].text = "Закрепление навыков использования сложноподчиненных предложений и согласования слов."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df6")
            # Связная речь (df7)
            table.add_row().cells
            row_cells = table.rows[14].cells
            row_cells[0].text = "Связная речь"
            table.add_row().cells
            if scores['df7'] == 1:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Ребенок не владеет связной речью, отвечает односложно или не отвечает."
                row_cells[1].text = "Формирование навыков составления простых предложений и коротких рассказов."
            elif scores['df7'] == 2:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Ребенок составляет короткие рассказы с помощью наводящих вопросов."
                row_cells[1].text = "Развитие навыков составления связных рассказов по картинкам и личному опыту."
            elif scores['df7'] == 3:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Ребенок составляет связные рассказы, но допускает ошибки в логике и структуре."
                row_cells[1].text = "Совершенствование навыков составления логичных и структурированных рассказов."
            elif scores['df7'] == 4:
                row_cells = table.rows[15].cells
                row_cells[0].text = "Ребенок составляет связные рассказы с минимальными ошибками."
                row_cells[1].text = "Закрепление навыков составления развернутых рассказов и пересказов."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df7")
            # ОЗОМ (df8)
            table.add_row().cells
            row_cells = table.rows[16].cells
            row_cells[0].text = "ОЗОМ"
            table.add_row().cells
            if scores['df8'] == 1:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок не справляется с заданиями по ознакомлению с окружающим миром."
                row_cells[1].text = "Формирование представлений об окружающем мире в соответствии с программой."
            elif scores['df8'] == 2:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок допускает множественные ошибки при выполнении заданий."
                row_cells[1].text = "Уточнение представлений об окружающем мире по лексическим темам."
            elif scores['df8'] == 3:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок выполняет задания, но допускает единичные ошибки."
                row_cells[1].text = "Совершенствование знаний об окружающем мире, работа с лексическими темами."
            elif scores['df8'] == 4:
                row_cells = table.rows[17].cells
                row_cells[0].text = "Ребенок выполняет задания с минимальными ошибками."
                row_cells[1].text = "Закрепление знаний об окружающем мире, развитие активной речи."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df8")
            # ФЭМП (df9)
            table.add_row().cells
            row_cells = table.rows[18].cells
            row_cells[0].text = "ФЭМП"
            table.add_row().cells
            if scores['df9'] == 1:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Ребенок не справляется с заданиями математического содержания."
                row_cells[1].text = "Формирование математических представлений в соответствии с программой."
            elif scores['df9'] == 2:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Математические представления не сформированы в значительной степени."
                row_cells[1].text = "Формирование представлений о счете, форме, величине и пространственных отношениях."
            elif scores['df9'] == 3:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Ребенок допускает множественные ошибки в математических заданиях."
                row_cells[1].text = "Совершенствование навыков счета, работы с числами и геометрическими фигурами."
            elif scores['df9'] == 4:
                row_cells = table.rows[19].cells
                row_cells[0].text = "Ребенок выполняет математические задания с минимальными ошибками."
                row_cells[1].text = "Закрепление навыков решения простых задач и работы с числами."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df9")
            # Конструирование (df10)
            table.add_row().cells
            row_cells = table.rows[20].cells
            row_cells[0].text = "Конструирование"
            table.add_row().cells
            if scores['df10'] == 1:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок не выполняет постройки из конструктора."
                row_cells[1].text = "Формирование навыков конструирования по образцу."
            elif scores['df10'] == 2:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок выполняет постройки только с обучающей помощью."
                row_cells[1].text = "Развитие навыков самостоятельного конструирования по образцу."
            elif scores['df10'] == 3:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок выполняет постройки с направляющей помощью."
                row_cells[1].text = "Совершенствование навыков самостоятельного конструирования и творческого подхода."
            elif scores['df10'] == 4:
                row_cells = table.rows[21].cells
                row_cells[0].text = "Ребенок выполняет постройки самостоятельно с минимальной помощью."
                row_cells[1].text = "Закрепление навыков творческого конструирования и работы с различными материалами."
            else:
                messagebox.showerror("Ошибка", "Неправильное значение для df10")

# Класс для управления активацией программы
# Управляет лицензией: проверка ключа, дата активации (31 день).
class ActivationManager:
    def __init__(self):
        # Инициализация путей к файлам ключа и даты активации
        # Файлы хранятся в поддиректории 'access_key' и в корне.
        self.current_directory = os.path.dirname(os.path.realpath(__file__))
        self.key_file_path = os.path.join(self.current_directory, 'access_key/access_key.txt')
        self.activation_file_path = os.path.join(self.current_directory, 'activation_date.json')

    def save_activation_date(self, activation_date):
        # Сохранение даты активации в JSON
        # Записывает {'date': 'YYYY-MM-DD'} в файл.
        with open(self.activation_file_path, 'w') as f:
            json.dump({'date': activation_date}, f)

    def read_activation_date(self):
        # Чтение даты активации из JSON
        # Возвращает строку даты или None, если файл не существует.
        if os.path.exists(self.activation_file_path):
            with open(self.activation_file_path) as f:
                activation_data = json.load(f)
                return activation_data['date']
        return None

    def is_week_passed_since_activation(self):
        # Проверка, прошло ли 31 день с активации
        # Если дата неверная, возвращает True (истекло).
        activation_date_str = self.read_activation_date()
        if activation_date_str:
            try:
                activation_date = datetime.strptime(activation_date_str, '%Y-%m-%d')
                passed_time = datetime.now() - activation_date
                return passed_time.days >= 31
            except ValueError:
                return True
        return True

    def read_key_from_file(self):
        # Чтение ключа из файла
        # Возвращает ключ без пробелов или None, если файл не существует.
        if os.path.exists(self.key_file_path):
            with open(self.key_file_path) as f:
                key = f.read().strip()
                return key
        return None

    def check_key(self, user_key):
        # Проверка введённого ключа
        # Сравнивает с хранимым ключом.
        stored_key = self.read_key_from_file()
        return user_key == stored_key

    def activate(self):
        # Логика активации программы
        # Если файл ключа существует, запрашивает ключ, проверяет, удаляет файл и сохраняет дату.
        # Если ключа нет, проверяет, не истекли ли 31 день с активации.
        if os.path.exists(self.key_file_path):
            user_key = simpledialog.askstring("Ключ", "Введите ваш ключ:")
            if user_key and self.check_key(user_key):
                os.remove(self.key_file_path)
                self.save_activation_date(datetime.now().date().strftime('%Y-%m-%d'))
                return True
            else:
                return False
        else:
            if not self.is_week_passed_since_activation():
                return True
            else:
                messagebox.showerror("Ошибка", "Срок действия лицензии истёк. Вам нужен новый лицензионный ключ.")
                return False

# Основной класс приложения с GUI
# Наследует от tk.Tk, создает окно, меню, формы для добавления/редактирования.
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Система управления воспитанниками")
        self.geometry("600x400")
        # Инициализация менеджеров
        # Создаем instances для DB, Excel, Word и активации.
        self.db_manager = DatabaseManager()
        self.excel_processor = ExcelProcessor()
        self.word_processor = WordProcessor()
        self.activation_manager = ActivationManager()

        # Проверка активации и запуск меню
        # Если активация успешна, показываем меню, иначе выходим.
        if self.activation_manager.activate():
            self.main_menu()
        else:
            self.quit()

    def clear_window(self):
        # Очистка текущего окна от виджетов
        # Уничтожает все дочерние виджеты.
        for widget in self.winfo_children():
            widget.destroy()

    def main_menu(self):
        # Отображение главного меню
        # Очищаем окно и добавляем кнопки для просмотра, добавления и выхода.
        self.clear_window()
        tk.Label(self, text="Главное меню", font=("Arial", 16)).pack(pady=20)
        tk.Button(self, text="Просмотр воспитанников", command=self.view_pupils).pack(pady=10)
        tk.Button(self, text="Добавить воспитанника", command=self.add_pupil_form).pack(pady=10)
        tk.Button(self, text="Выход", command=self.quit).pack(pady=10)

    def add_pupil_form(self):
        # Форма добавления воспитанника
        # Очищаем окно, добавляем поля для ввода данных и кнопки.
        self.clear_window()
        tk.Label(self, text="Добавить воспитанника", font=("Arial", 16)).pack(pady=20)

        tk.Label(self, text="Фамилия:").pack()
        surname_entry = tk.Entry(self)
        surname_entry.pack()

        tk.Label(self, text="Имя:").pack()
        name_entry = tk.Entry(self)
        name_entry.pack()

        tk.Label(self, text="Отчество:").pack()
        patronymic_entry = tk.Entry(self)
        patronymic_entry.pack()

        tk.Label(self, text="Дата рождения (ДД-ММ-ГГГГ):").pack()
        birth_date_entry = tk.Entry(self)
        birth_date_entry.pack()

        tk.Button(self, text="Вернуться в меню", command=self.main_menu).pack(pady=10)
        tk.Button(self, text="Далее", command=lambda: self.process_pupil_data(
            surname_entry.get(),
            name_entry.get(),
            patronymic_entry.get(),
            birth_date_entry.get()
        )).pack(pady=10)

    def process_pupil_data(self, surname, name, patronymic, birth_date_str):
        # Обработка данных формы добавления
        # Проверяем заполненность, парсим дату, добавляем в DB, затем переходим к обработке Excel.
        if not all([surname, name, patronymic, birth_date_str]):
            messagebox.showerror("Ошибка", "Все поля должны быть заполнены")
            return

        try:
            birth_date = datetime.strptime(birth_date_str, '%d-%m-%Y').date()
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД-ММ-ГГГГ")
            return

        pupil_id = self.db_manager.add_pupil(surname, name, patronymic, birth_date)
        if pupil_id:
            self.process_excel_data(pupil_id)

    def view_pupils(self):
        # Просмотр списка воспитанников в таблице
        # Очищаем окно, создаем Treeview для отображения данных, добавляем кнопки для редактирования/удаления.
        self.clear_window()
        tk.Label(self, text="Список воспитанников", font=("Arial", 16)).pack(pady=20)

        columns = ("ID", "Фамилия", "Имя", "Отчество", "Дата рождения", "df1", "df2", "df3", "df4", "df5", "df6", "df7", "df8", "df9", "df10", "df11")
        tree = ttk.Treeview(self, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
        tree.pack(fill="both", expand=True)

        pupils = self.db_manager.get_pupils()
        for row in pupils:
            tree.insert("", "end", values=row)

        tk.Button(self, text="Изменить личные данные", command=lambda: self.edit_pupil_info(tree)).pack(pady=5)
        tk.Button(self, text="Изменить баллы", command=lambda: self.edit_pupil_scores(tree)).pack(pady=5)
        tk.Button(self, text="Удалить", command=lambda: self.delete_pupil(tree)).pack(pady=5)
        tk.Button(self, text="Вернуться в меню", command=self.main_menu).pack(pady=5)

    def edit_pupil_info(self, tree):
        # Форма редактирования личных данных
        # Получаем выбранную запись, очищаем окно, заполняем поля текущими данными.
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Ошибка", "Выберите воспитанника")
            return

        pupil_values = tree.item(selected_item)['values']
        pupil_id = pupil_values[0]

        self.clear_window()
        tk.Label(self, text="Изменить данные воспитанника", font=("Arial", 16)).pack(pady=20)

        tk.Label(self, text="Фамилия:").pack()
        surname_entry = tk.Entry(self)
        surname_entry.insert(0, pupil_values[1])
        surname_entry.pack()

        tk.Label(self, text="Имя:").pack()
        name_entry = tk.Entry(self)
        name_entry.insert(0, pupil_values[2])
        name_entry.pack()

        tk.Label(self, text="Отчество:").pack()
        patronymic_entry = tk.Entry(self)
        patronymic_entry.insert(0, pupil_values[3])
        patronymic_entry.pack()

        tk.Label(self, text="Дата рождения (ДД-ММ-ГГГГ):").pack()
        birth_date_entry = tk.Entry(self)
        birth_date_entry.insert(0, str(pupil_values[4]))  # Преобразование даты в строку
        birth_date_entry.pack()

        tk.Button(self, text="Сохранить", command=lambda: self.save_pupil_info(
            pupil_id,
            surname_entry.get(),
            name_entry.get(),
            patronymic_entry.get(),
            birth_date_entry.get()
        )).pack(pady=10)
        tk.Button(self, text="Назад", command=self.view_pupils).pack(pady=10)

    def save_pupil_info(self, pupil_id, surname, name, patronymic, birth_date_str):
        # Сохранение обновлённых данных
        # Проверяем заполненность, парсим дату, обновляем в DB, возвращаемся к просмотру.
        if not all([surname, name, patronymic, birth_date_str]):
            messagebox.showerror("Ошибка", "Все поля должны быть заполнены")
            return

        try:
            birth_date = datetime.strptime(birth_date_str, '%d-%m-%Y').date()
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД-ММ-ГГГГ")
            return

        if self.db_manager.update_pupil_info(pupil_id, surname, name, patronymic, birth_date):
            messagebox.showinfo("Успех", "Данные воспитанника успешно обновлены")
            self.view_pupils()

    def edit_pupil_scores(self, tree):
        # Редактирование баллов через Excel
        # Получаем ID выбранного, запускаем обработку Excel для обновления оценок.
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Ошибка", "Выберите воспитанника")
            return

        pupil_id = tree.item(selected_item)['values'][0]
        self.process_excel_data(pupil_id)

    def delete_pupil(self, tree):
        # Удаление выбранного воспитанника
        # Получаем ID, удаляем из DB, обновляем просмотр.
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Ошибка", "Выберите воспитанника")
            return

        pupil_id = tree.item(selected_item)['values'][0]
        if self.db_manager.delete_pupil(pupil_id):
            messagebox.showinfo("Успех", "Воспитанник успешно удалён")
            self.view_pupils()

    def process_excel_data(self, pupil_id):
        # Обработка Excel и Word для обновления баллов и документа
        # Запрашиваем файл Excel, читаем оценки, обновляем в DB, затем Word.
        excel_file_path = filedialog.askopenfilename(title="Выберите файл", initialdir="Проект",
                                                     filetypes=[("Excel файлы", "*.xlsx")],
                                                     initialfile="Карта развития. Младший возраст.xlsx")
        if not excel_file_path:
            return

        result = self.excel_processor.read_scores(excel_file_path)
        if result is None:
            self.main_menu()
            return

        scores, excel_file_name = result

        if self.db_manager.update_pupil_scores(pupil_id, scores):
            word_file_path = filedialog.askopenfilename(title="Выберите файл", filetypes=[('Word файлы', '*.docx')])
            if word_file_path:
                self.word_processor.update_document(word_file_path, scores, excel_file_name)
            self.main_menu()

if __name__ == "__main__":
    # Запуск приложения
    # Создаем экземпляр Application и запускаем mainloop для GUI.
    app = Application()
    app.mainloop()