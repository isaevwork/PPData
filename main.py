import os
import warnings
import pandas as pd
from pptx.dml.color import RGBColor
from openpyxl import load_workbook
from pptx import Presentation
from datetime import datetime
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.dml import MSO_THEME_COLOR_INDEX

# Путь к папке work
user_profile = os.getenv('USERPROFILE')
work_folder = os.path.join(user_profile, 'Downloads', 'work')

# Список для хранения путей к файлам Excel
excel_files = []

# Глобальная переменная для хранения имени главной папки
folder_name = None

# Глобальная переменная для хранения Excel файла
wb = None

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
image_folder = os.path.join(work_folder, folder_name)

# -----------------------------------------------------------------------------------------
# Определение параметров текстового блока
left = Inches(6.9)
top = Inches(7.9)
width = Inches(3)
height = Inches(2)

# Список индексов слайдов, для которых нужно создать текстовые блоки
slide_indexes = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]

# Цикл для создания текстовых блоков на каждом слайде
for index in slide_indexes:
    name_textbox = prs.slides[index].shapes.add_textbox(left, top, width, height)
    name_textbox.rotation = 270
    tf = name_textbox.text_frame
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = f"{folder_name}"
    p.font.bold = False
    p.font.size = Pt(15)


# Функция для вставки изображений на слайд

def insert_images(names, positions, idx):
    """
      Функция для добавления изображений на конкретный слайд презентации.
      Args:
          names (list): Список имен изображений.
          positions (list): Словарь с позициями изображений для каждого слайда.
          idx (int): Индекс слайда, на который добавляются изображения.
      """
    extensions = [".jpg", ".png", ".jpeg", ".gif"]  # Расширения изображений для проверки
    slide = prs.slides[idx]  # Получаем слайд по индексу

    for name, position in zip(names, positions):
        found_image = False
        for extension in extensions:
            image_path = os.path.join(image_folder, name + extension)
            if os.path.exists(image_path):
                img_left, img_top, img_width, img_height = position
                slide.shapes.add_picture(image_path, img_left, img_top, img_width, img_height)
                found_image = True
                break  # Нашли изображение, прекращаем поиск расширения
        if not found_image:
            print(f"Изображение {name} не найдено на слайде {idx}.")


# -----------------------------------------------------------------------------------------

# Задаем имя пациента, врача и дату
left = Inches(2.9)  # Расстояние от правого края слайда
top = Inches(7.75)  # Расстояние от верхнего края слайда
width = Inches(4)  # Ширина, чтобы занять всю ширину слайда
height = Inches(0.5)  # Высота, чтобы занять всю высоту слайда
name_textbox = prs.slides[0].shapes.add_textbox(left, top, width, height)
tf = name_textbox.text_frame
tf.word_wrap = True
p = tf.add_paragraph()
p.text = f"{folder_name}"
p.font.size = Pt(14)
p.font.bold = False

date_left = Inches(5)
date_top = Inches(8.9)
date_width = Inches(4)
name_textbox = prs.slides[0].shapes.add_textbox(date_left, date_top, date_width, height)
tf_date = name_textbox.text_frame
tf_date.word_wrap = True
p_date = tf_date.add_paragraph()
p_date.text = f"{datetime.today().strftime('%d.%m.%Y')}"
# -------------------------------------------------------

# Слайд № 1
print("Слайд №1 сформирован")
# -------------------------------------------------------

# Слайд № 2
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["2.1", "2.2", "2.3", "2.4"]]
images_name_2 = ["2q", "2w", "2e", "2r"]
images_position_2 = [
    (Inches(0.4), Inches(1.5), Inches(2.6), Inches(3.6)),
    (Inches(5.4), Inches(1.5), Inches(2.6), Inches(3.6)),
    (Inches(0.7), Inches(7.8), Inches(3), Inches(3.6)),
    (Inches(4.5), Inches(7.8), Inches(3), Inches(3.6))
]
slide_index_2 = 2
insert_images(images_name_2, images_position_2, slide_index_2)
print("Слайд №2 сформирован")
# -------------------------------------------------------

# Слайд № 3
ws = wb["Лист2"]

# Создаем пустой DataFrame
slideThree_data = []
slideThree_MT = []

for row in ws.iter_rows(min_row=2, max_row=9, min_col=16, max_col=17, values_only=True):
    slideThree_MT.append(list(row))

# Размер и положение данных на слайде
current_left = Inches(2.8)  # Левая граница
current_top = Inches(5.53)  # Верхняя граница
cell_width = Inches(2.5)  # Ширина ячейки
cell_height = Inches(0.27)  # Высота ячейки
font_size = Pt(12)  # Размер шрифта


def transform_data(data):
    transformed_data = []
    for idx, sublist in enumerate(data):
        transformed_sublist = []
        for item in sublist:
            if isinstance(item, (int, float)):
                # Проверяем, является ли текущий подсписок последним или предпоследним в массиве данных
                if idx == len(data) - 2 or idx == len(data) - 1:
                    transformed_sublist.append('{:.1f}%'.format(item * 100).replace('.', ','))
                else:
                    transformed_sublist.append(round(item, 2))
            else:
                transformed_sublist.append(item)
        transformed_data.append(transformed_sublist)
    return transformed_data


transformed_dataframe = transform_data(slideThree_MT)

# Размещение данных на слайде
for i, row_data in enumerate(transformed_dataframe):
    for j, value in enumerate(row_data):
        # Рассчитываем координаты для текущей ячейки
        cell_left = current_left + j * cell_width
        cell_top = current_top + i * cell_height

        # Добавление текстового блока на слайд с текущими координатами
        text_frame = prs.slides[3].shapes.add_textbox(cell_left, cell_top, cell_width, cell_height).text_frame
        p = text_frame.add_paragraph()
        p.text = str(value)
        p.font.size = font_size
        p.font.name = "Montserrat Medium"
        p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

# Проходимся по строкам и столбцам в Excel и добавляем их в DataFrame
for row in ws.iter_rows(values_only=True):
    slideThree_data.append(row)

# Создаем DataFrame из данных
df = pd.DataFrame(slideThree_data).iloc[0:9, 14:18]
sub_up_df = pd.DataFrame(slideThree_data).iloc[1:2, 1:13]
sub_lower_df = pd.DataFrame(slideThree_data).iloc[2:3, 1:13]

# Получаем количество строк и столбцов в таблице
num_rows, num_cols = df.shape
sub_up_num_rows, sub_up_num_cols = sub_up_df.shape
sub_lower_num_rows, sub_lower_num_cols = sub_lower_df.shape

# Определяем размеры и позицию таблицы на слайде
sub_up_left = Inches(0.9)
sub_up_top = Inches(2.3)
sub_up_width = Inches(6)
sub_up_height = Inches(0.3)

sub_lower_left = Inches(1.05)
sub_lower_top = Inches(3.8)
sub_lower_width = Inches(6.1)
sub_lower_height = Inches(0.3)

# Добавляем таблицу на слайд
sub_up_table = prs.slides[3].shapes.add_table(1, 12, sub_up_left, sub_up_top, sub_up_width, sub_up_height).table
sub_lower_table = prs.slides[3].shapes.add_table(1, 12, sub_lower_left, sub_lower_top, sub_lower_width,
                                                 sub_lower_height).table


def fill_table_from_df(data_frame, target_table):
    # Определение количества строк и столбцов в DataFrame
    temp_num_rows, temp_num_cols = data_frame.shape
    # Проход по каждой строке DataFrame
    for rows in range(temp_num_rows):
        # Проход по каждому столбцу DataFrame
        for col in range(temp_num_cols):
            # Получение значения из DataFrame
            temp_value = data_frame.iloc[rows, col]
            # Получение ячейки таблицы PowerPoint
            temp_cell = target_table.cell(rows, col)
            # Преобразование значения в строку и запись в ячейку таблицы
            temp_cell.text = str(int(temp_value)) if isinstance(temp_value, float) else str(temp_value)
            temp_cell.text_frame.paragraphs[0].font.name = "Montserrat Medium"


# Применение функции для заполнения верхней и нижней таблиц из DataFrame
fill_table_from_df(sub_up_df, sub_up_table)
fill_table_from_df(sub_lower_df, sub_lower_table)

sub_up_table.cell(0, 0).text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
sub_up_table.columns[0].width = Inches(0.75)
sub_up_table.columns[1].width = Inches(0.55)
sub_up_table.columns[5].width = Inches(0.6)
sub_up_table.columns[6].width = Inches(0.5)
sub_up_table.columns[10].width = Inches(0.6)

sub_lower_table.cell(0, 0).text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
sub_lower_table.columns[0].width = Inches(0.92)
sub_lower_table.columns[1].width = Inches(0.5)
sub_lower_table.columns[3].width = Inches(0.4)
sub_lower_table.columns[4].width = Inches(0.4)
sub_lower_table.columns[5].width = Inches(0.4)
sub_lower_table.columns[6].width = Inches(0.4)
sub_lower_table.columns[7].width = Inches(0.4)
sub_lower_table.columns[10].width = Inches(0.6)


# Устанавливаем прозрачный цвет заливки для каждой ячейки таблицы
def set_transparent_fill(t_table):
    for t_row in t_table.rows:
        for t_cell in t_row.cells:
            t_cell.fill.solid()
            t_cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
            t_cell.fill.background()


# Применяем функцию к верхней и нижней таблицам
set_transparent_fill(sub_up_table)
set_transparent_fill(sub_lower_table)

print("Слайд №3 сформирован")
# -------------------------------------------------------

# Слайд № 4
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["2.1", "2.2", "2.3", "2.4"]]
images_name_4 = ["4q", "4w", "4e", "4r", "4t", "4y"]
images_position_4 = [
    (Inches(0.9), Inches(1.2), Inches(2.9), Inches(2.75)),
    (Inches(4.4), Inches(1.2), Inches(2.9), Inches(2.75)),
    (Inches(0.9), Inches(4), Inches(2.9), Inches(2.75)),
    (Inches(4.4), Inches(4), Inches(2.9), Inches(2.75)),
    (Inches(0.8), Inches(9.1), Inches(2.4), Inches(2.2)),
    (Inches(4.3), Inches(9.3), Inches(3.2), Inches(2))
]
slide_index_4 = 4
insert_images(images_name_4, images_position_4, slide_index_4)
print("Слайд №4 сформирован")
# -------------------------------------------------------

# Слайд № 5
# Размеры и положения областей для изображений
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["5q", "5w", "5e", "5r", "5t", "5y"]]
images_name_5 = ["5q", "5w", "5e", "5r", "5t", "5y",]
images_position_5 = [
    (Inches(0.6), Inches(1.2), Inches(3.3), Inches(3)),
    (Inches(4.4), Inches(1.2), Inches(3.3), Inches(3)),

    (Inches(0.6), Inches(5), Inches(3.1), Inches(3)),
    (Inches(4.5), Inches(5), Inches(3.1), Inches(3)),

    (Inches(0.6), Inches(9.2), Inches(3.4), Inches(1.4)),
    (Inches(4.2), Inches(9.2), Inches(3.4), Inches(1.4))
]
slide_index_5 = 5
insert_images(images_name_5, images_position_5, slide_index_5)
print("Слайд №5 сформирован")
# -------------------------------------------------------


# Слайд № 6
# Размеры и положения областей для изображений
# left = Inches(1)     # Левая граница изображения
# top = Inches(1)      # Верхняя граница изображения
# width = Inches(5)    # Ширина изображения
# height = Inches(3)   # Высота изображения

# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["6q", "6w", "6e", "6r", "6t", "6y", "6u", "6i", "6o"]]
images_name_6 = ["6q", "6w", "6e", "6r", "6t", "6y", "6u", "6i", "6o"]
images_position_6 = [
    (Inches(0.8), Inches(1.1), Inches(3.2), Inches(1.7)),
    (Inches(4.2), Inches(1.1), Inches(3), Inches(1.6)),

    (Inches(0.8), Inches(2.95), Inches(3.2), Inches(1.7)),
    (Inches(4.2), Inches(3), Inches(3), Inches(1.6)),

    (Inches(0.8), Inches(4.8), Inches(3.2), Inches(1.7)),
    (Inches(4.2), Inches(4.9), Inches(3), Inches(1.6)),

    (Inches(0.5), Inches(9.2), Inches(2.4), Inches(2)),
    (Inches(2.9), Inches(9.3), Inches(2.4), Inches(1.8)),
    (Inches(5.3), Inches(9.2), Inches(2.4), Inches(2))
]
slide_index_6 = 6
insert_images(images_name_6, images_position_6, slide_index_6)
print("Слайд №6 сформирован")
# -------------------------------------------------------

# Слайд № 7
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["9", "6"]]
images_name_7 = ["9", "6"]
images_position_7 = [
    (Inches(0.6), Inches(1.5), Inches(7), Inches(4)),
    (Inches(0.6), Inches(6.3), Inches(7), Inches(4.6)),
]
slide_index_7 = 7
insert_images(images_name_7, images_position_7, slide_index_7)
print("Слайд №7 сформирован")
# -------------------------------------------------------

# Слайд № 8
# TODO image_names = [f"{folder_name}_{image}" for image in ["11", "вч", "нч"]]
images_name_8 = ["11", "вч", "нч"]
images_position_8 = [
    (Inches(0.5), Inches(1.1), Inches(7.4), Inches(3.4)),
    (Inches(0.5), Inches(5.1), Inches(7.2), Inches(2.9)),
    (Inches(0.5), Inches(8.5), Inches(7.2), Inches(2.9)),
]
slide_index_8 = 8
insert_images(images_name_8, images_position_8, slide_index_8)
print("Слайд №8 сформирован")
# -------------------------------------------------------

# Слайд № 9
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["22", "1", "2"]]
images_name_9 = ["22", "1", "2"]
images_position_9 = [
    (Inches(0.4), Inches(1.8), Inches(7.6), Inches(3.2)),
    (Inches(1.5), Inches(5.8), Inches(5.9), Inches(2.6)),
    (Inches(1.5), Inches(8.7), Inches(5.9), Inches(2.6)),
]
slide_index_9 = 9
insert_images(images_name_9, images_position_9, slide_index_9)
print("Слайд №9 сформирован")
# -------------------------------------------------------

# Слайд № 10
# Размеры и положения областей для изображений
# TODO image_names = [f"{folder_name}_{image}" for image in ["444", "33", "44"]]
images_name_10 = ["444", "33", "44"]
images_position_10 = [
    (Inches(0.6), Inches(1.5), Inches(7), Inches(4.6)),

    (Inches(0.5), Inches(7.5), Inches(3.5), Inches(3.5)),
    (Inches(4.1), Inches(7.5), Inches(3.5), Inches(3.5)),
]
slide_index_10 = 10
insert_images(images_name_10, images_position_10, slide_index_10)
print("Слайд №10 сформирован")
# -------------------------------------------------------

# Слайд № 11
# TODO image_names = [f"{folder_name}_{image}" for image in ["222", "333"]]
images_name_11 = ["222", "333"]
images_position_11 = [
    (Inches(0.5), Inches(1.3), Inches(7.1), Inches(4.7)),
    (Inches(0.5), Inches(6.6), Inches(7.1), Inches(4.7)),
]
slide_index_11 = 11
insert_images(images_name_11, images_position_11, slide_index_11)
print("Слайд №11 сформирован")
# -------------------------------------------------------

# Слайд № 12
# TODO image_names = [f"{folder_name}_{image}" for image in ["0"]]
images_name_12 = ["0"]
images_position_12 = [
    (Inches(0.8), Inches(1.3), Inches(6.8), Inches(6.8)),
]
slide_index_12 = 12
insert_images(images_name_12, images_position_12, slide_index_12)
print("Слайд №12 сформирован")
# -------------------------------------------------------

# Слайд № 13
# Размеры и положения областей для изображений
# left = Inches(1)     # Левая граница изображения
# top = Inches(1)      # Верхняя граница изображения
# width = Inches(5)    # Ширина изображения
# height = Inches(3)   # Высота изображения
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["77", "88", "3", "000"]]
images_name_13 = ["77", "88", "3", "000"]
images_position_13 = [
    (Inches(1.4), Inches(3.2), Inches(2.3), Inches(2.1)),
    (Inches(4.8), Inches(3.2), Inches(2.3), Inches(2.1)),
    (Inches(1), Inches(8.6), Inches(3), Inches(2.5)),
    (Inches(4.5), Inches(8.6), Inches(3), Inches(2.5)),
]
slide_index_13 = 13
insert_images(images_name_13, images_position_13, slide_index_13)

mt = wb["Лист1"]

# Пустой DataFrame
slideFour_data = []

for row in mt.iter_rows(values_only=True):
    slideFour_data.append(row)

up_dff = pd.DataFrame(slideFour_data).iloc[8:14, 2:6]
# lower_dff = pd.DataFrame(slideFour_data).iloc[14:24, 0:4]

up_num_rows, up_num_cols = up_dff.shape
# lower_num_rows, lower_num_cols = dff_lower.shape

# Определяем размеры и позицию таблицы на слайде
up_left = Inches(3.7)
up_top = Inches(0.8)
up_width = Inches(3)
up_height = Inches(1)

# lower_left = Inches(0.55)
# lower_top = Inches(6.2)
# lower_width = Inches(7.1)
# lower_height = Inches(3.5)


# Добавляем таблицу на слайд
# up_table = prs.slides[13].shapes.add_table(6, 4, up_left, up_top, up_width, up_height).table
up_table = prs.slides[0].shapes.add_table(6, 4, up_left, up_top, up_width, up_height).table
# lower_table = prs.slides[13].shapes.add_table(9, 3, lower_left, lower_top, lower_width, lower_height).table

# # Применение функции для заполнения верхней и нижней таблиц из DataFrame
# fill_table_from_df(up_dff, up_table)
# Стилизация верхней таблицы
for r in range(up_num_rows):
    for c in range(up_num_cols):
        value = up_dff.iloc[r, c]
        cell = up_table.cell(r, c)
        cell.text = str(value)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        cell.fill.fore_color.theme_color = MSO_THEME_COLOR_INDEX.LIGHT_1
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        cell.text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
        # cell.fill.background()
        # if c == 1 or c == 2:

print("Слайд №13 сформирован")
# -------------------------------------------------------

if folder_name:
    prs.save(os.path.join(work_folder, f"{folder_name}.pptx"))
