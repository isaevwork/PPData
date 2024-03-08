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
        # Составляем полный путь к текущему файлу
        file_path = os.path.join(root, file)
        # Если файл имеет расширение .xlsx и находится не в корневой папке work, добавляем его путь в список
        if file.endswith(".xlsx") and root != work_folder:
            excel_files.append(file_path)

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
    print("Файл Excel не найден в подпапках папки work.")

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


def get_text_color(last_value):
    """
    Возвращает цвет текста в зависимости от значения last_value.
    """
    if isinstance(last_value, (int, float)):
        last_value = float(last_value)

        if last_value is not None:
            if last_value > 3 or last_value < -3:
                return RGBColor(255, 0, 0)  # Красный цвет
            elif -2 <= last_value <= 2:
                if -0.9 <= last_value <= 0.9:
                    return RGBColor(0, 0, 0)  # Черный цвет
                elif last_value < 0:
                    return RGBColor(0, 0, 255)  # Синий цвет
                else:
                    return RGBColor(0, 255, 0)  # Зеленый цвет
    return RGBColor(0, 0, 0)  # Черный цвет по умолчанию


def add_text_to_slide(presentation, slide_index, slide_data, current_left, current_top, cell_width, cell_height,
                      font_s):
    for i, row_data in enumerate(slide_data):
        for j, value in enumerate(row_data):
            # Рассчитываем координаты для текущей ячейки
            cell_left = current_left + j * cell_width
            cell_top = current_top + i * cell_height

            # Получаем значение last_value из последней ячейки текущей строки
            last_value = row_data[-1]

            # Получаем цвет текста на основе значения last_value
            color = get_text_color(last_value)

            # Преобразуем значение в строку, если оно не None
            if value is not None:
                if isinstance(value, (int, float)):
                    text_value = str(round(value, 2))
                else:
                    text_value = str(value)
            else:
                text_value = ""

            if j == 1:
                cell_left += Inches(0.8)
            if j == 2:
                cell_left += Inches(0.8)
            if j == 3:
                cell_left += Inches(0.8)
            if j == 4:
                cell_left += Inches(0.8)

            if i == 6:
                cell_top += Inches(0.04)
            if i == 7:
                cell_top += Inches(0.04)

            # Добавление текстового блока на слайд с текущими координатами и цветом
            table_frame = presentation.slides[slide_index].shapes.add_textbox(cell_left, cell_top, cell_width,
                                                                              cell_height).text_frame
            q = table_frame.add_paragraph()
            q.text = text_value
            q.font.size = font_s
            q.font.name = "Montserrat Medium"
            q.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
            q.font.color.rgb = color  # Устанавливаем цвет текста


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
    (Inches(4.6), Inches(7.7), Inches(3), Inches(3.7))
]
slide_index_2 = 2
insert_images(images_name_2, images_position_2, slide_index_2)
print("Слайд №2 сформирован")
# -------------------------------------------------------

# Слайд № 3
ws2 = wb["Лист2"]

# Создаем пустой DataFrame
slideThree_data = []
slideThree_MT = []

for row in ws2.iter_rows(min_row=2, max_row=9, min_col=16, max_col=17, values_only=True):
    slideThree_MT.append(list(row))

# Размер и положение данных на слайде
c_left = Inches(2.8)  # Левая граница
c_top = Inches(5.53)  # Верхняя граница
c_width = Inches(2.5)  # Ширина ячейки
c_height = Inches(0.27)  # Высота ячейки
f_size = Pt(12)  # Размер шрифта


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
        cell_left = c_left + j * c_width
        cell_top = c_top + i * c_height

        # Добавление текстового блока на слайд с текущими координатами
        text_frame = prs.slides[3].shapes.add_textbox(cell_left, cell_top, c_width, c_height).text_frame
        p = text_frame.add_paragraph()
        p.text = str(value)
        p.font.size = f_size
        p.font.name = "Montserrat Medium"
        p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

# Проходимся по строкам и столбцам в Excel и добавляем их в DataFrame
for row in ws2.iter_rows(values_only=True):
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
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["5q", "5w", "5e", "5r", "5t", "5y"]]
images_name_5 = ["5q", "5w", "5e", "5r", "5t", "5y", ]
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
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["6q", "6w", "6e", "6r", "6t", "6y", "6u", "6i", "6o"]]
images_name_6 = ["6q", "6w", "6e", "6r", "6t", "6y"]
images_position_6 = [
    (Inches(0.8), Inches(1.3), Inches(3.5), Inches(1.8)),

    (Inches(0.8), Inches(3.15), Inches(3.5), Inches(1.8)),

    (Inches(0.8), Inches(5), Inches(3.5), Inches(1.8)),

    (Inches(0.8), Inches(7.1), Inches(2.6), Inches(2.2)),
    (Inches(2.9), Inches(9.4), Inches(2.6), Inches(2.2)),
    (Inches(5), Inches(7.1), Inches(2.6), Inches(2.2))
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
    (Inches(1.2), Inches(1.4), Inches(6), Inches(5.5)),

    (Inches(0.5), Inches(7.5), Inches(3.5), Inches(3.5)),
    (Inches(4.1), Inches(7.5), Inches(3.5), Inches(3.5)),
]
slide_index_10 = 10
insert_images(images_name_10, images_position_10, slide_index_10)
print("Слайд №10 сформирован")
# -------------------------------------------------------
# Слайд № 11
# Размеры и положения областей для изображений
# TODO image_names = [f"{folder_name}_{image}" for image in ["4", "5", "55"]]
images_name_11 = ["4", "5", "55"]
images_position_11 = [
    (Inches(0.6), Inches(1.4), Inches(3.5), Inches(3.8)),
    (Inches(4.2), Inches(1.4), Inches(3.5), Inches(3.8)),

    (Inches(1.4), Inches(5.8), Inches(5.4), Inches(5.6)),
]
slide_index_11 = 11
insert_images(images_name_11, images_position_11, slide_index_11)
print("Слайд №11 сформирован")
# -------------------------------------------------------

# Слайд № 12
# TODO image_names = [f"{folder_name}_{image}" for image in ["12q"]]
images_name_12 = ["12q"]
images_position_12 = [
    (Inches(0.65), Inches(1.6), Inches(7.1), Inches(8.5)),
]
slide_index_12 = 12
insert_images(images_name_12, images_position_12, slide_index_12)
print("Слайд №12 сформирован")
# -------------------------------------------------------

# Слайд № 13
# TODO image_names = [f"{folder_name}_{image}" for image in ["13q"]]
images_name_13 = ["13q"]
images_position_13 = [
    (Inches(0.6), Inches(1.8), Inches(6.8), Inches(6.6)),
]
slide_index_13 = 13
insert_images(images_name_13, images_position_13, slide_index_13)

ws1 = wb["Лист1"]
params13_data = []

# Заполняем DataFrame данными из листа Excel
for row in ws1.iter_rows(min_row=29, max_row=36, min_col=2, max_col=6, values_only=True):
    params13_data.append(list(row))

# Размеры и положение данных на слайде

params13_left = Inches(1.3)  # Левая граница
params13_top = Inches(9.08)  # Верхняя граница
params13_width = Inches(1.12)  # Ширина ячейки
params13_height = Inches(0.27)  # Высота ячейки
font_size = Pt(11)  # Размер шрифта

add_text_to_slide(prs, 13, params13_data, params13_left, params13_top, params13_width, params13_height, font_size)

print("Слайд №13 сформирован")
# -------------------------------------------------------


# Слайд № 14
# TODO image_names = [f"{folder_name}_{image}" for image in ["77", "88"]]
images_name_14 = ["77", "88"]
images_position_14 = [
    (Inches(0.6), Inches(8.1), Inches(3.4), Inches(3.3)),
    (Inches(4.2), Inches(8.1), Inches(3.4), Inches(3.3)),
]
slide_index_14 = 14
insert_images(images_name_14, images_position_14, slide_index_14)

up_params14_data = []
lower_params14_data = []

# Заполняем DataFrame данными из листа Excel
for row in ws1.iter_rows(min_row=9, max_row=14, min_col=2, max_col=6, values_only=True):
    up_params14_data.append(list(row))
for row in ws1.iter_rows(min_row=17, max_row=26, min_col=2, max_col=6, values_only=True):
    lower_params14_data.append(list(row))

# Размеры и положение данных на слайде
up_params14_left = Inches(1.2)  # Левая граница
up_params14_top = Inches(1.96)  # Верхняя граница
up_params14_width = Inches(1.12)  # Ширина ячейки
up_params14_height = Inches(0.27)  # Высота ячейки

lower_params14_left = Inches(1.2)  # Левая граница
lower_params14_top = Inches(4.3)  # Верхняя граница
lower_params14_width = Inches(1.12)  # Ширина ячейки
lower_params14_height = Inches(0.27)  # Высота ячейки

font_size = Pt(9)

add_text_to_slide(prs, 14, up_params14_data, up_params14_left, up_params14_top, up_params14_width, up_params14_height,
                  font_size)
add_text_to_slide(prs, 14, lower_params14_data, lower_params14_left, lower_params14_top, lower_params14_width,
                  lower_params14_height, font_size)
print("Слайд №14 сформирован")
# -------------------------------------------------------

# Слайд №15
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["333", "temp15"]]
images_name_15 = ["333", "temp15"]
images_position_15 = [
    (Inches(0.6), Inches(1.5), Inches(7), Inches(4.7)),
    (Inches(0.6), Inches(7), Inches(7), Inches(4.4)),
]
slide_index_15 = 15
insert_images(images_name_15, images_position_15, slide_index_15)
print("Слайд №15 сформирован")
# -------------------------------------------------------

# Слайд №16
# Массив имен изображений с префиксом папки
# TODO image_names = [f"{folder_name}_{image}" for image in ["222", "0"]]
images_name_16 = ["222", "0"]
images_position_16 = [
    (Inches(0.8), Inches(1.2), Inches(6.5), Inches(4)),
    (Inches(2), Inches(7.65), Inches(4.3), Inches(3.9)),
]
slide_index_16 = 16
insert_images(images_name_16, images_position_16, slide_index_16)
print("Слайд №16 сформирован")
# -------------------------------------------------------


# Слайд №17
# Определяем тенденцию к классу
# anb_value = ws1['L42'].value
# beta_angle = ws1['L44'].value
# wits_appraisal = ws1['L46'].value
# sassouni = ws1['L125'].value
# apdi_value = ws1['L43'].value
# pnsa_value = ws1['C9'].value
# jj_value = ws1['C10'].value
# sna_value = ws1['C13'].value
# ppsn_value = ws1['C14'].value
#
#
# # Определяем класс в зависимости от значения ANB
# anb_trend_class = ""
# if anb_value > 4:
#     anb_skeletal_class = "II"
# elif anb_value < 0:
#     anb_skeletal_class = "III"
# else:
#     anb_skeletal_class = "I"
#     # Если класс "I", проверяем дополнительные условия
#     if 0 <= anb_value <= 0.4:
#         anb_trend_class = "с тенденцией к III классу"
#     elif 3.6 <= anb_value <= 4:
#         anb_trend_class = "с тенденцией к II классу"
#
# # Определяем класс в зависимости от значения BETA ANGLE
# beta_trend_class = ""
# if beta_angle > 35:
#     beta_skeletal_class = "III"
# elif beta_angle < 27:
#     beta_skeletal_class = "II"
# else:
#     beta_skeletal_class = "I"
#     # Если класс "I", проверяем дополнительные условия
#     if 34.6 <= beta_angle <= 35:
#         beta_trend_class = "с тенденцией к III классу"
#     elif 27 <= beta_angle <= 27.4:
#         beta_trend_class = "с тенденцией к II классу"
#
# # Определяем класс в зависимости от значения Wits Appraisal
# has_value = ""
# wits_trend_class = ""
# if wits_appraisal > 2.1:
#     wits_skeletal_class = "II"
#     has_value = "наличие"
# elif wits_appraisal < -2.9:
#     wits_skeletal_class = "III"
#     has_value = "наличие"
# else:
#     wits_skeletal_class = "I"
#     has_value = "отсутствие"
#     # Если класс "I", проверяем дополнительные условия
#     if -2.5 <= wits_appraisal <= -2.9:
#         wits_trend_class = "с тенденцией к III классу"
#     elif 1.7 <= wits_appraisal <= 2.1:
#         wits_trend_class = "с тенденцией к II классу"
#
#
# # Определяем класс в зависимости от значения Wits Appraisal
# sassouni_text = ""
# sassouni_trend_class = ""
# has_direction = ""
# if sassouni > 3:
#     sassouni_skeletal_class = "III"
#     has_direction = "кзади"
# elif sassouni < 0:
#     sassouni_skeletal_class = "II"
#     has_direction = "кпереди"
# else:
#     sassouni_skeletal_class = "I"
#     # Если класс "I", проверяем дополнительные условия
#     if 2.6 <= sassouni <= 3:
#         sassouni_trend_class = "с тенденцией к III классу"
#     elif 0.1 <= sassouni <= 0.4:
#         sassouni_trend_class = "с тенденцией к II классу"
#
# sassouni_text_not_null = f"""Соотношение челюстей по методике Sassouni говорит за {sassouni_skeletal_class} скелетный класс {sassouni_trend_class} — базальная дуга проходит на {sassouni} мм {has_direction} от точки В (N = 0,0 мм ± 3,0 мм)."""
# sassouni_text_null = f"""Соотношение челюстей по методике Sassouni говорит за I скелетный класс — базальная дуга проходит через точку B (N = 0,0 мм ± 3,0 мм)."""
#
# if sassouni == 0:
#     sassouni_text = sassouni_text_null
# else:
#     sassouni_text = sassouni_text_not_null
#
# # Определяем класс в зависимости от значения APDI
# apdi_trend_class = ""
# if apdi_value > 86.4:
#     apdi_skeletal_class = "III"
# elif apdi_value < 76.4:
#     apdi_skeletal_class = "II"
# else:
#     apdi_skeletal_class = "I"
#     # Если класс "I", проверяем дополнительные условия
#     if 86 <= apdi_value <= 86.4:
#         apdi_trend_class = "с тенденцией к III классу"
#     elif 76.4 <= apdi_value <= 76.8:
#         apdi_trend_class = "с тенденцией к II классу"
#
# # Определяем класс в зависимости от значения PNSA
# pnsa_upper_limit = ws1['D9'].value + 3.5
# pnsa_lower_limit = ws1['D9'].value - 3.5
# pnsa_trend_class = ""
# pnsa_status = ""
# if pnsa_value > pnsa_upper_limit:
#     pnsa_status = "увеличению"
# elif pnsa_value < pnsa_lower_limit:
#     pnsa_status = "уменьшению"
# else:
#     pnsa_status = "норме"
#     # Если класс "I", проверяем дополнительные условия
#     if 86 <= pnsa_value <= 86.4:
#         pnsa_trend_class = "с тенденцией к III классу"
#     elif 76.4 <= pnsa_value <= 76.8:
#         pnsa_trend_class = "с тенденцией к II классу"
#
# # Определяем класс в зависимости от значения Ширины скелетного базиса (J-J)
# jj_upper_limit = ws1['D10'].value + 3
# jj_lower_limit = ws1['D10'].value - 3
# jj_status = ""
# if jj_value > jj_upper_limit:
#     jj_status = "расширению"
# elif jj_value < jj_lower_limit:
#     jj_status = "сужению"
# else:
#     jj_status = "норме"
#
# # Определяем класс в зависимости от значения SNA
# if sna_value > 85:
#     sna_status = "прогнатии"
# elif sna_value < 79:
#     sna_status = "ретрогнатии"
# else:
#     sna_status = "нормогнатии"
#
# # Определяем класс в зависимости от значения PP/SN
# if ppsn_value > 12:
#     ppsn_status = "ретроинклинации"
# elif ppsn_value < 5:
#     ppsn_status = "антеинклинации"
# else:
#     ppsn_status = "нормоинклинации"
#
# # Формируем текст, вставляя значения переменных
# resume_text1 = f"""
# Соотношение челюстей в сагиттальной плоскости.
# Межапикальный угол (<ANB) – {anb_value}˚, что соответствует соотношению челюстей по {anb_skeletal_class} скелетному классу {anb_trend_class} (N = 2,0˚ ± 2,0˚).
# Угол Бета (< Beta Angle) – {beta_angle}˚, что cоответствует соотношению челюстей по {beta_skeletal_class} скелетному классу {beta_trend_class} (N = 31,0˚ ± 4,0˚).
# Параметр Wits (Wits Appraisal.) – {wits_appraisal} мм что указывает на {has_value} в расположении апикальных базисов верхней и нижней челюстей в сагиттальной плоскости и говорит за {wits_skeletal_class} скелетный класс {wits_trend_class} (N = -0,4 мм ± 2,5 мм).
# {sassouni_text}
# Параметр APDI, указывающий на дисплазию развития челюстей в сагиттальной плоскости, равен {apdi_value}˚ и говорит за {apdi_skeletal_class} скелетный класс {apdi_trend_class} (N = 81,4˚ ± 5,0˚).
# Размер и положение верхней челюсти.
# Длина основания верхней челюсти (PNS-A) – {pnsa_value} мм, что соответствует {pnsa_status} в пределах нормы (N = {ws1['D9'].value} мм ± 3,5 мм).
# Ширина основания верхней (J-J) челюсти –  {jj_value} мм, что соответствует {jj_status} в пределах нормы (N = {ws1['D10'].value} мм ± 3,0 мм):  справа – {ws1['C11'].value} мм, слева – {ws1['C12'].value} мм (N = {ws1['D10'].value / 2} мм ± 1,5 мм).
# Положение верхней челюсти по сагиттали  (<SNA) – {sna_value}˚, что соответствует {sna_status} (N = 82,0˚ ±  3,0˚).
# Положение верхней челюсти по вертикали  (<SN-Palatal Plane) – {ppsn_value}˚, что соответствует {ppsn_status} (N = 8,0˚ ± 3,0˚).
# Roll ротация отсутствует\  вправо (по часовой стрелке) \ влево (против часовой стрелки).
# Yaw ротация отсутствует \ вправо  (по часовой стрелке) \ влево (против часовой стрелки).
#
# """
#
#
#
# # Добавляем текст на слайд
# text_left = Inches(0.6)
# text_top = Inches(1)
# text_width = Inches(7)
# text_height = Inches(5)
# name_textbox = prs.slides[0].shapes.add_textbox(text_left, text_top, text_width, text_height)
# text_frame = name_textbox.text_frame
# text_frame.word_wrap = True
# paragraph = text_frame.add_paragraph()
# paragraph.font.size = Pt(10.5)
# paragraph.font.bold = False
# paragraph.font.name = "Montserrat"
# paragraph.text = cephalometry_text
#
# print("Слайд №17 сформирован")
# -------------------------------------------------------

# -------------------------------------------------------
# Слайд №16
# Переменные с тенденциями к классу

go_me_r_value = ws1['C17'].value
go_me_l_value = ws1['C18'].value

go_go_r_value = ws1['C19'].value
go_go_l_value = ws1['C20'].value

ar_go_r_value = ws1['C21'].value
ar_go_l_value = ws1['C22'].value

md_md_value = ws1['c23'].value

snb_value = ws1['c24'].value

mp_sn_value = ws1['C25'].value

chin_displacement = ws1['C26'].value

ans_quotient = ws1['N9'].value

assessment_growth_type = ws1['L98'].value



def compare_value(value1, value2, name):
    if value1 is None or value2 is None:
        print(f"Ошибка: Недостаточно данных для сравнения {name}")
        return "None"
    return 'меньше' if value1 < value2 else 'больше'


# Определяем класс в зависимости от значения Go-Me
go_me_status = compare_value(go_me_r_value, go_me_l_value, "Go-Me")

# Определяем класс в зависимости от значения Go-Go
go_go_status = compare_value(go_go_r_value, go_go_l_value, "Go-Go")

# Определяем класс в зависимости от значения Ar-Go
ar_go_status = compare_value(ar_go_r_value, ar_go_l_value, "Ar-Go")

# Определяем класс в зависимости от значения Md-Md
md_upper_limit = ws1['D23'].value + 3
md_lower_limit = ws1['D23'].value - 3
md_status = ""
if md_md_value > md_upper_limit:
    md_status = "расширению"
elif md_md_value < md_lower_limit:
    md_status = "сужению"
else:
    md_status = "норме"

# Определяем класс в зависимости от значения <SNB
if snb_value > 83:
    snb_status = "прогнатии"
elif snb_value < 77:
    snb_status = "ретрогнатии"
else:
    snb_status = "нормогнатии"

# Определяем класс в зависимости от значения <MP\SN
if mp_sn_value > 36:
    mp_sn_status = "ретроинклинации"
elif mp_sn_value < 28:
    mp_sn_status = "антеинклинации"
else:
    mp_sn_status = "нормоинклинации"

# Определение смещения подбородка
if chin_displacement > 0:
    chin_displacement_status = f"влево на {round(chin_displacement, 2)} мм."
elif chin_displacement < 0:
    chin_displacement_status = f"вправо на {round(chin_displacement, 2)} мм."
else:
    chin_displacement_status = "не выявлено."

# Определяем класс в зависимости от значения (N-ANS) / (ANS-Gn)
if ans_quotient > 0.89:
    ans_quotient_status = "негармоничное"
elif ans_quotient < 0.71:
    ans_quotient_status = "негармоничное"
else:
    ans_quotient_status = "гармоничное."

# Формируем текст, вставляя значения переменных
resume_text2 = f"""
Размер и положение нижней челюсти.
Длина тела нижней челюсти (Go-Me): справа – {go_me_r_value} мм, слева – {go_me_l_value}  мм (N = {ws1['M59'].value} мм ± 5,0 мм).
Длина тела нижней челюсти справа {go_me_status}, чем слева на {round(abs(go_me_r_value - go_me_l_value), 2)} мм.
Длина ветви нижней челюсти (Co-Go) : справа – {go_go_r_value}  мм,  слева – {go_go_l_value}  мм (N = {ws1['D19'].value} мм ± 4,0 мм).
Длина ветви нижней челюсти справа {go_go_status}, чем слева на {round(abs(go_go_r_value - go_go_l_value), 2)} мм.
Гониальный угол (<Ar-Go-Me): справа –  {ar_go_r_value}˚,  слева – {ar_go_l_value}˚ (N = {ws1['D21'].value}˚ ± 5,0˚).
Гониальный угол справа {ar_go_status}, чем слева на {round(abs(ar_go_r_value - ar_go_l_value), 2)}˚.
Ширина базиса нижней челюсти (Md-Md) – {md_md_value} мм, что соответствует {md_status} (N = {ws1['D23'].value} мм ± 3,0 мм).
Положение нижней челюсти по сагиттали  (<SNB) – {snb_value}˚, что соответствует {snb_status} (N = 80,0˚ ± 3,0˚).
Положение нижней челюсти по вертикали (<MP-SN) – {mp_sn_value}˚, что соответствует {mp_sn_status} (N = 32,0˚ ± 4,0˚).
Смещение подбородка {chin_displacement_status}
Roll ротация отсутствует \  вправо (по часовой стрелке) \ влево (против часовой стрелки).
Yaw ротация отсутствует \ вправо  (по часовой стрелке) \ влево (против часовой стрелки).
Вертикальные параметры.
Вертикальное лицевое соотношение (N-ANS/ANS-Gn) {ans_quotient_status} – {round(ans_quotient, 2)} (N = 0,8 ± 0,09).
Отношение задней высоты лица к передней (S-Go/N-Gn) – {assessment_growth_type}% (N = 63,0% ± 2,0%).

"""

# Добавляем текст на слайд
text_left = Inches(0.6)
text_top = Inches(1)
text_width = Inches(7)
text_height = Inches(5)
name_textbox = prs.slides[0].shapes.add_textbox(text_left, text_top, text_width, text_height)
text_frame = name_textbox.text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(10.5)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = resume_text2

print("Слайд №18 сформирован")

if folder_name:
    save_folder = os.path.join(work_folder, folder_name)
    prs.save(os.path.join(save_folder, f"{folder_name}.pptx"))



# Высота нижней трети лица по Ricketts (<ANS-Xi-Pm) – 42,9˚, что соответствует норме \ увеличению \ уменьшению (N = IVP 44,6˚ ± 5,5˚).
# Параметр ODI – 53,9˚, что соответствует норме  \ тенденции к  \ вертикальной резцовой дизокклюзии \ глубокой резцовой окклюзии (N = 74,5˚ ±  5,0˚).
#
# Положение резцов.
# Межрезцовый угол : справа – 134,0˚, слева – 129,0˚(N = 130,0˚ ± 5,0˚).
# Наклон зуба 1.1 – 102,1˚,  зуба 2.1 – 104,9˚ (N = 110,0˚± 5,0˚).
# Наклон зуба 3.1 – 93,4˚, зуба 4.1 – 95,8˚ (N = 95,0˚ ± 5,0˚).
#
# Penn анализ.
# Ширина базиса нижней челюсти – 55,0 мм. Фактическая ширина базиса верхней челюсти – 58,0 мм.
# Требуемая ширина базиса верхней челюсти = 60,0 мм.
# Дефицит ширины скелетного базиса верхней челюсти  отсутствует \ составляет 2,0 мм.
#
# Воздухоносные пути.
# Сужение \и уменьшение объема нижних воздухоносных путей. Сужения \и уменьшения объема нижних воздухоносных путей не выявлено. Значительное сужение на  уровне середины\ верхней  \ нижней трети C2 \ между С2-С3 = 182 мм², (N = 182 мм²).
