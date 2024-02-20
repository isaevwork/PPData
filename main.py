import os
import warnings
from openpyxl import load_workbook
from pptx import Presentation
from datetime import datetime
import pandas as pd
from pptx.util import Inches, Pt
from openpyxl.utils.dataframe import dataframe_to_rows

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
        wb = load_workbook(filename=excel_file_path)
        print(f"Excel файл {os.path.basename(excel_file_path)} найден в папке {os.path.dirname(excel_file_path)}.")

        # Извлекаем имя папки из пути
        folder_name = os.path.basename(os.path.dirname(excel_file_path))
else:
    print("Файл Excel не найден в папке work.")



# Загружаем презентацию
prs = Presentation(os.path.join(os.getenv('USERPROFILE'), 'Downloads', 'parser', 'PPData', 'FDTemp.pptx'))



# slide_layout = root.slide_layouts[1]
# slide = root.slides.add_slide(slide_layout)
#
# for row in ws.iter_rows(min_row=1):
#     title = row[0].value
    # content = row[1].value

    # title_placeholder = slide.shapes.title
    # content_placeholder = slide.placeholders[1]

    # title_placeholder.text = str(title)
    # content_placeholder.text = str(content)

# title_placeholder = slide.shapes.title
# content_placeholder = slide.placeholders[1]

# title_placeholder.text = str("A" + str(len(title_placeholder.text)))
# content_placeholder.text = str(content)

# Слайд № 1
# Задаем имя пациента, врача и дату
left = Inches(2.9)  # Расстояние от правого края слайда
top = Inches(7.75)   # Расстояние от верхнего края слайда
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
date_top = Inches(8.9)   # Расстояние от верхнего края слайда
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
top = Inches(7.9)   # Расстояние от верхнего края слайда
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
top = Inches(7.9)   # Расстояние от верхнего края слайда
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
image_names = ["2.1", "2.2", "2.3", "2.4",]

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
top = Inches(7.9)   # Расстояние от верхнего края слайда
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

#
# # Создаем пустой DataFrame
# fourSlide_data = []
# ws = wb["Лист2"]
#
# # Проходимся по строкам и столбцам в Excel и добавляем их в DataFrame
# for row in ws.iter_rows(values_only=True):
#     fourSlide_data.append(row)
#
#
# # Создаем DataFrame из данных
# df = pd.DataFrame(fourSlide_data).iloc[1:9, 14:18]
# print(df.shape)
#
# # Получаем количество строк и столбцов в таблице
# num_rows, num_cols = df.shape
#
# # Определяем размеры и позицию таблицы на слайде
# left = Inches(0.55)
# top = Inches(6.2)
# width = Inches(7.1)
# height = Inches(3.5)
#
# # Добавляем таблицу на слайд
# table = prs.slides[3].shapes.add_table(num_rows, num_cols, left, top, width, height).table
#
# # Заполнение таблицы данными из DataFrame
# # Заполнение таблицы данными из Excel
# for r in range(num_rows):
#     for c in range(num_cols):
#         table.cell(r, c).text = str(df.iloc[r, c])
#





if folder_name:
    prs.save(os.path.join(work_folder, f"{folder_name}.pptx"))

