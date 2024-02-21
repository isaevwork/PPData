import os
import warnings
import pandas as pd
from openpyxl.styles import Alignment
from pptx.dml.color import RGBColor
from openpyxl import load_workbook
from pptx import Presentation
from datetime import datetime
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

# Путь к папке work
user_profile = os.getenv('USERPROFILE')
work_folder = os.path.join(user_profile, 'Downloads', 'work')

# Список для хранения путей к файлам Excel
excel_files = []

# Глобальная переменная для хранения имени главной папки
folder_name = None

# Глобальная переменная для хранения Excel файла
wb = None

# Расширения для изображений
extensions = ['.jpg', '.jpeg', '.png']

# Путь к файлу Excel с таблицей
# Рекурсивно обходим все подпапки внутри папки work
for root, dirs, files in os.walk(work_folder):
    # Просматриваем все файлы в текущей подпапке
    for file in files:
        # Если файл имеет расширение .xlsx, добавляем его путь в список
        if file.endswith(".xlsx"):
            excel_files.append(os.path.join(root, file))

# Если найден хотя бы один файл Excel, загружаем первый из них
if excel_files:
    excel_file_path = excel_files[0]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(filename=excel_file_path, data_only=True)
        print(f"Excel файл {os.path.basename(excel_file_path)} найден в папке {os.path.dirname(excel_file_path)}.")

        # Извлекаем имя папки из пути
        folder_name = os.path.basename(os.path.dirname(excel_file_path))
else:
    print("Файл Excel не найден в папке work.")

# Загружаем презентацию
prs = Presentation(os.path.join(os.getenv('USERPROFILE'), 'Downloads', 'parser', 'PPData', 'FDTemp.pptx'))

# Слайд № 1
# Задаем имя пациента, врача и дату
left = Inches(2.9)  # Расстояние от правого края слайда
top = Inches(7.75)  # Расстояние от верхнего края слайда
width = Inches(4)
height = Inches(0.5)  # Высота, чтобы занять всю высоту слайда
name_textbox = prs.slides[0].shapes.add_textbox(left, top, width, height)
tf = name_textbox.text_frame
tf.word_wrap = True
p = tf.add_paragraph()
p.text = f"{folder_name}"
p.font.size = Pt(14)
p.font.bold = False
# -------------------------------------------------
date_left = Inches(5)
date_top = Inches(8.9)  # Расстояние от верхнего края слайда
date_width = Inches(4)
name_textbox = prs.slides[0].shapes.add_textbox(date_left, date_top, date_width, height)
tf_date = name_textbox.text_frame
tf_date.word_wrap = True
p_date = tf_date.add_paragraph()
p_date.text = f"{datetime.today().strftime('%d.%m.%Y')}"
# -------------------------------------------------------
print("Первый слайд готов")
# Слайд № 2
# Задаем имя пациента вдоль правой границы
left = Inches(6.9)  # Расстояние от правого края слайда
top = Inches(7.9)  # Расстояние от верхнего края слайда
width = Inches(3)
height = Inches(2)  # Высота, чтобы занять всю высоту слайда
name_textbox = prs.slides[1].shapes.add_textbox(left, top, width, height)
name_textbox.rotation = 270
tf = name_textbox.text_frame
tf.word_wrap = True
p = tf.add_paragraph()
p.text = f"{folder_name}"
p.font.bold = False
p.font.size = Pt(14)
# -------------------------------------------------------
print("Второй слайд готов")
# Слайд № 3
# Вставляем изображения и ФИО пациента
left = Inches(6.9)  # Расстояние от правого края слайда
top = Inches(7.9)  # Расстояние от верхнего края слайда
width = Inches(3)
height = Inches(2)  # Высота, чтобы занять всю высоту слайда
name_textbox = prs.slides[2].shapes.add_textbox(left, top, width, height)
name_textbox.rotation = 270
tf = name_textbox.text_frame
tf.word_wrap = True
p = tf.add_paragraph()
p.text = f"{folder_name}"
p.font.bold = False
p.font.size = Pt(14)

# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["2.1", "2.2", "2.3", "2.4"]]
image_names = ["2.1", "2.2", "2.3", "2.4", ]

# Путь к моим изображениям
image_folder = os.path.join(work_folder, folder_name)

# Задаем путь к изображению
image_paths = [os.path.join(image_folder, name) for name in image_names]

# Размеры и положения областей для изображений
image_positions = [
    (Inches(0.4), Inches(1.5), Inches(2.6), Inches(3.6)),
    (Inches(5.4), Inches(1.5), Inches(2.6), Inches(3.6)),
    (Inches(0.7), Inches(7.8), Inches(3), Inches(3.6)),
    (Inches(4.5), Inches(7.8), Inches(3), Inches(3.6))
]

# Устанавливаем размеры и позицию изображения на слайде
# left = Inches(1)     # Левая граница изображения
# top = Inches(1)      # Верхняя граница изображения
# width = Inches(5)    # Ширина изображения
# height = Inches(3)   # Высота изображения

# Вставка изображений на слайд
for image_name, position in zip(image_names, image_positions):
    found_image = False
    for extension in extensions:
        image_path = os.path.join(image_folder, image_name + extension)
        if os.path.exists(image_path):
            left, top, width, height = position
            prs.slides[2].shapes.add_picture(image_path, left, top, width, height)
            found_image = True
            break  # Нашли изображение, прекращаем поиск расширения
    if not found_image:
        print(f"Изображение {image_name} не найдено.")
# -------------------------------------------------------
print("Третий слайд готов")

# Слайд № 4
# Вставляем таблицу из Excel файла и ФИО пациента
left = Inches(6.9)
top = Inches(7.9)  # Расстояние от верхнего края слайда
width = Inches(3)
height = Inches(2)
name_textbox = prs.slides[3].shapes.add_textbox(left, top, width, height)
name_textbox.rotation = 270
tf = name_textbox.text_frame
tf.word_wrap = True
p = tf.add_paragraph()
p.text = f"{folder_name}"
p.font.bold = False
p.font.size = Pt(14)

ws = wb["Лист2"]

# Создаем пустой DataFrame
slideThree_data = []

# Проходимся по строкам и столбцам в Excel и добавляем их в DataFrame
for row in ws.iter_rows(values_only=True):
    slideThree_data.append(row)

# Создаем DataFrame из данных
df = pd.DataFrame(slideThree_data).iloc[0:9, 14:18]
sub_up_df = pd.DataFrame(slideThree_data).iloc[1:2, 1:13]
sub_lower_df = pd.DataFrame(slideThree_data).iloc[2:3, 1:13]

value = sub_up_df.iloc[0, 2]
print(value)
# Получаем количество строк и столбцов в таблице
num_rows, num_cols = df.shape
sub_up_num_rows, sub_up_num_cols = sub_up_df.shape
sub_lower_num_rows, sub_lower_num_cols = sub_lower_df.shape

# Определяем размеры и позицию таблицы на слайде
left = Inches(0.55)
top = Inches(6.2)
width = Inches(7.1)
height = Inches(3.5)

sub_up_left = Inches(1)
sub_up_top = Inches(2.5)
sub_up_width = Inches(5.9)
sub_up_height = Inches(0.4)

sub_lower_left = Inches(1.05)
sub_lower_top = Inches(3.67)
sub_lower_width = Inches(6.2)
sub_lower_height = Inches(0.4)

# Добавляем таблицу на слайд
table = prs.slides[3].shapes.add_table(9, 3, left, top, width, height).table
sub_up_table = prs.slides[3].shapes.add_table(1, 12, sub_up_left, sub_up_top, sub_up_width, sub_up_height).table
sub_lower_table = prs.slides[3].shapes.add_table(1, 12, sub_lower_left, sub_lower_top, sub_lower_width,
                                                 sub_lower_height).table


# Функция для форматирования значения
def format_value(value):
    if isinstance(value, (float, int)):
        if (r == 7 and c == 1) or (r == 7 and c == 2) or (r == 8 and c == 1):
            # Форматирование процентного значения
            return "{:.2f}%".format(value * 100)
        else:
            # Округление числовых значений до сотых
            return str(round(value, 2))
    else:
        return str(value) if value is not None else ""


def fill_table_from_df(data_frame, taret_table):
    # Определение количества строк и столбцов в DataFrame
    temp_num_rows, temp_num_cols = data_frame.shape
    # Проход по каждой строке DataFrame
    for rows in range(temp_num_rows):
        # Проход по каждому столбцу DataFrame
        for col in range(temp_num_cols):
            # Получение значения из DataFrame
            temp_value = data_frame.iloc[rows, col]
            # Получение ячейки таблицы PowerPoint
            temp_cell = taret_table.cell(rows, col)
            # Преобразование значения в строку и запись в ячейку таблицы
            temp_cell.text = str(int(temp_value)) if isinstance(temp_value, float) else str(temp_value)
            temp_cell.text_frame.paragraphs[0].font.size = Pt(18)


# Заполнение таблицы данными из DataFrame и центрирование текста
for r in range(num_rows):
    for c in range(num_cols):
        value = df.iloc[r, c]
        cell = table.cell(r, c)
        cell.text = format_value(value)
        if c == 1 or c == 2:
            cell.text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

# Применение функции для заполнения верхней и нижней таблиц из DataFrame
fill_table_from_df(sub_up_df, sub_up_table)
fill_table_from_df(sub_lower_df, sub_lower_table)

sub_up_table.cell(0, 0).text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
sub_up_table.columns[0].width = Inches(0.6)
sub_up_table.columns[1].width = Inches(0.6)
sub_up_table.columns[5].width = Inches(0.6)
sub_up_table.columns[6].width = Inches(0.5)
sub_up_table.columns[10].width = Inches(0.6)

sub_lower_table.cell(0, 0).text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
sub_lower_table.columns[0].width = Inches(0.9)
sub_lower_table.columns[1].width = Inches(0.5)
sub_lower_table.columns[3].width = Inches(0.4)
sub_lower_table.columns[4].width = Inches(0.4)
sub_lower_table.columns[5].width = Inches(0.4)
sub_lower_table.columns[6].width = Inches(0.4)
sub_lower_table.columns[7].width = Inches(0.4)
sub_lower_table.columns[10].width = Inches(0.6)

# Устанавливаем текст "Индекс Тона" в ячейку
table.cell(4, 0).text = "Индекс Тона"
# Объединяем ячейки 4 и 5 в первом столбце
table.cell(4, 0).merge(table.cell(5, 0))

# Устанавливаем прозрачный цвет заливки для каждой ячейки таблицы
for row in sub_up_table.rows:
    for cell in row.cells:
        cell.fill.solid()
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

for row in sub_lower_table.rows:
    for cell in row.cells:
        cell.fill.solid()
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

if folder_name:
    prs.save(os.path.join(work_folder, f"{folder_name}.pptx"))
