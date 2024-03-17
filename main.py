import os
import warnings
from PIL import Image
from pptx.dml.color import RGBColor
from openpyxl import load_workbook
from pptx import Presentation
from datetime import datetime
from pptx.util import Inches, Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT


def draw_table(text):
    # Определение длины самой длинной строки в тексте
    max_length = max(len(line) for line in text.split('\n'))

    # Верхняя граница таблицы
    print("-" * (max_length + 4))

    # Выводим текст с отступами слева и справа
    for line in text.split('\n'):
        print(f"| {line.center(max_length)} |")

    # Нижняя граница таблицы
    print("-" * (max_length + 4))


# Текст для помещения внутрь таблицы
text = """Предупреждение: использование этого скрипта полностью лежит на ответственности конечного пользователя.
Автор скрипта не несет никакой ответственности!
Мы просим всех пользователей внимательно изучить результаты этого скрипта перед его использованием."""


def main():
    work_folder = os.path.join(os.environ['USERPROFILE'], 'Downloads', 'WORK')
    os.chdir(work_folder)

    # Вызываем функцию и передаем текст внутрь таблицы
    draw_table(text)

    extensions = [".jpg", ".png", ".jpeg", ".gif"]  # Расширения изображений для переименования

    for folder_name in os.listdir():
        if os.path.isdir(folder_name):
            os.chdir(folder_name)
            renamed_files = set()  # Сет для отслеживания уже переименованных файлов
            for filename in os.listdir():
                if any(filename.lower().endswith(ext) for ext in extensions):
                    if folder_name not in filename and filename not in renamed_files:
                        new_filename = f"{folder_name}_{filename}"
                        os.rename(filename, new_filename)
                        renamed_files.add(new_filename)  # Добавляем новое имя файла в сет
            os.chdir('..')


main()

# Путь к папке work
user_profile = os.getenv('USERPROFILE')
work_folder = os.path.join(user_profile, 'Downloads', 'work')

# Список для хранения путей к файлам Excel
excel_files = []

# Глобальная переменная для хранения имени главной папки
folder_name = None

# Глобальная переменная для хранения Excel файла
wb = None

passed = "\033[92m✓\033[0m"
notpassed = "\033[91m✗\033[0m"
well = "\033[93m◆\033[0m"

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
        print(f"{well} Excel файл {os.path.basename(excel_file_path)} найден в папке {os.path.dirname(excel_file_path)}.")

        # Извлекаем имя папки из пути
        folder_name = os.path.basename(os.path.dirname(excel_file_path))
else:
    print(f"{notpassed} Файл Excel не найден в подпапках папки work.")

# Загружаем презентацию
prs = Presentation(os.path.join(os.getenv('USERPROFILE'), 'Downloads', 'work', 'FDTemp.pptx'))
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
    p.font.size = Pt(12)
    p.font.name = "Montserrat"


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
        for extension in extensions:
            image_path = os.path.join(image_folder, name + extension)
            if os.path.exists(image_path):
                img_left, img_top, img_width, img_height = position
                slide.shapes.add_picture(image_path, img_left, img_top, img_width, img_height)
                break
        else:
            print(f" {notpassed} Изображение {name} не найдено на слайде {idx}.")


def format_with_comma(number):
    """
    Форматирует число с одним десятичным знаком и заменяет точку на запятую.
    """
    formatted_number = "{:,.1f}".format(number)
    return formatted_number.replace(".", ",")


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
                    return RGBColor(6, 102, 6)  # Зеленый цвет
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


def crop_image(img_path, out_path, new_width, new_height):
    """
    Обрезает и изменяет размеры изображения и сохраняет его.
    Args:
        img_path (str): Путь к исходному изображению.
        out_path (str): Путь для сохранения обрезанного изображения.
        new_width (int): Новая ширина изображения.
        new_height (int): Новая высота изображения.
    """
    image = Image.open(img_path)
    width_i, height_i = image.size

    # Определяем координаты области обрезки относительно центра изображения
    left_i = (width_i - new_width) // 2
    top_i = (height_i - new_height) // 2
    right_i = (width_i + new_width) // 2
    bottom_i = (height_i + new_height) // 2

    cropped_image = image.crop((left_i, top_i, right_i, bottom_i))
    cropped_image.save(out_path)


def apply_crop_to_images(images_list, new_w, new_h, suffix=""):
    for img_name in images_list:
        # Формируем полный путь к файлу изображения
        for extension in ['.jpg', '.png']:  # Проверяем два наиболее распространенных формата
            img_path = os.path.join(image_folder, f"{img_name}{extension}")
            if os.path.exists(img_path):  # Проверяем существует ли файл
                break
        else:
            continue  # Пропускаем это изображение, если оно не найдено

        # Формируем путь для сохранения обрезанного изображения
        out_path = os.path.join(image_folder, f"{img_name}{suffix}{extension}")

        # Обрезаем изображение и сохраняем его
        crop_image(img_path, out_path, new_w, new_h)


def rename_image(old_name, new_name):
    temp_folder = os.path.join(image_folder, "temp")

    # Проверяем существование папки temp, если ее нет, создаем
    if not os.path.exists(temp_folder):
        try:
            os.makedirs(temp_folder)
        except Exception as e:
            print(f"{notpassed}Ошибка при создании папки temp: {str(e)}")
            return

    for extension in ['.jpg', '.png']:
        img_path = os.path.join(image_folder, f"{old_name}{extension}")
        if os.path.exists(img_path):
            break
    else:
        print(f"{notpassed} Изображение {old_name} не найдено.")
        return

    try:
        # Открываем изображение
        image = Image.open(img_path)
        # Формируем новый путь для сохранения в папку temp
        new_img_path = os.path.join(temp_folder, f"{new_name}{extension}")
        # Создаем копию изображения с новым именем
        image_copy = image.copy()
        # Сохраняем копию в новом пути
        image_copy.save(new_img_path)
        print(f"""{passed} Изображение "{old_name}" успешно скопировано с новым именем "{new_name}" и сохранено в папку temp.""")
    except Exception as e:
        print(f"{notpassed} Ошибка при копировании и сохранении изображения: {str(e)}")


print("<------------------------------------------------------------------------------------------->")

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

# Слайд № 1
print(f"{passed} Слайд №1 сформирован 🎉")

print("<------------------------------------------------------------------------------------------->")
# Слайд № 2
# Массив имен изображений с префиксом папки
images_name_2 = [f"{folder_name}_{image}" for image in ["2q", "2w", "2e", "2r"]]

images_position_2 = [
    (Inches(0.4), Inches(1.5), Inches(2.6), Inches(3.6)),
    (Inches(5.4), Inches(1.5), Inches(2.6), Inches(3.6)),
    (Inches(0.7), Inches(7.8), Inches(3), Inches(3.6)),
    (Inches(4.6), Inches(7.7), Inches(3), Inches(3.7))
]

# Применяем функцию к каждому изображению
apply_crop_to_images([f"{folder_name}_{image}" for image in ["2q", "2w"]], 1700, 2200)
apply_crop_to_images([f"{folder_name}_{image}" for image in ["2e", "2r"]], 2300, 2600)

insert_images(images_name_2, images_position_2, 2)
print(f"{passed} Слайд №2 сформирован 🎉")

# Слайд № 3
ws2 = wb["Лист2"]

# Создаем пустой DataFrame
slideThree_MT = list(ws2.iter_rows(min_row=2, max_row=9, min_col=16, max_col=17, values_only=True))

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


def fill_table(present, slide_index, slide_data, cl, ct, cw, ch, fs, column_offsets):
    for i, row_datas in enumerate(slide_data):
        for j, ce_value in enumerate(row_datas):
            # Рассчитываем координаты для текущей ячейки
            ce_left = cl + j * cw

            # Преобразуем значение в строку, если оно не None
            text_value = str(round(ce_value, 2)) if isinstance(ce_value, (int, float)) else str(ce_value)

            # Получаем информацию о смещении для текущего столбца
            offset_info = column_offsets.get(j)
            if offset_info:
                add_offset, offset_value = offset_info
                if add_offset:
                    ce_left += offset_value
                else:
                    ce_left = offset_value

            # Добавление текстового блока на слайд с текущими координатами и цветом
            table_frame = present.slides[slide_index].shapes.add_textbox(ce_left, ct, cw, ch).text_frame
            p = table_frame.add_paragraph()
            p.text = text_value
            p.font.size = fs
            p.font.name = "Montserrat Medium"
            p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER


up_data3 = list(ws2.iter_rows(min_row=2, max_row=2, min_col=2, max_col=13, values_only=True))
lower_data3 = list(ws2.iter_rows(min_row=3, max_row=3, min_col=2, max_col=13, values_only=True))

column_offsets_up = {
    1: (True, Inches(0.1)),
    2: (True, Inches(0.05)),
    7: (False, Inches(4.65)),
    8: (False, Inches(5.15)),
    9: (False, Inches(5.65)),
    10: (False, Inches(6.2)),
    11: (False, Inches(6.76))
}
column_offsets_lower = {
    1: (True, Inches(0.15)),
    2: (True, Inches(0.2)),
    3: (True, Inches(0.2)),
    4: (False, Inches(3.3)),
    5: (False, Inches(3.74)),
    6: (False, Inches(4.1)),
    10: (False, Inches(5.9)),
    11: (False, Inches(6.5))
}

fill_table(prs, 3, up_data3, Inches(0.9), Inches(2.45), Inches(0.55), Inches(0.27), Pt(14), column_offsets_up)
fill_table(prs, 3, lower_data3, Inches(1.3), Inches(3.65), Inches(0.45), Inches(0.27), Pt(14), column_offsets_lower)

print("<------------------------------------------------------------------------------------------->")
print(f"{passed} Слайд №3 сформирован 🎉")

print("<------------------------------------------------------------------------------------------->")
# Слайд № 4
# Массив имен изображений с префиксом папки
images_name_4 = [f"{folder_name}_{image}" for image in ["4q", "4w", "4e", "4r", "00", "4t"]]
images_position_4 = [
    (Inches(0.9), Inches(1.2), Inches(2.9), Inches(2.75)),
    (Inches(4.4), Inches(1.2), Inches(2.9), Inches(2.75)),
    (Inches(0.9), Inches(4), Inches(2.9), Inches(2.75)),
    (Inches(4.4), Inches(4), Inches(2.9), Inches(2.75)),
    (Inches(0.8), Inches(9.1), Inches(2.4), Inches(2.2)),
    (Inches(4.3), Inches(9.3), Inches(3.2), Inches(2))
]

insert_images(images_name_4, images_position_4, 4)
rename_image(images_name_4[4], "размеры апикальных базисов вч и нч")
print(f"{passed} Слайд №4 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 5
# Массив имен изображений с префиксом папки
images_name_5 = [f"{folder_name}_{image}" for image in ["5q", "5w", "5e", "5r", "5t", "5y"]]
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
print(f"{passed} Слайд №5 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")


# Слайд № 6
# Массив имен изображений с префиксом папки
images_name_6 = [f"{folder_name}_{image}" for image in ["6q", "6w", "6e", "6r", "6t", "6y"]]
images_position_6 = [
    (Inches(0.8), Inches(1.3), Inches(3.5), Inches(1.8)),

    (Inches(0.8), Inches(3.15), Inches(3.5), Inches(1.8)),

    (Inches(0.8), Inches(5), Inches(3.5), Inches(1.8)),

    (Inches(0.8), Inches(7.1), Inches(3), Inches(2.2)),
    (Inches(2.9), Inches(9.4), Inches(3.4), Inches(2.2)),
    (Inches(5), Inches(7.1), Inches(3), Inches(2.2))
]
slide_index_6 = 6
insert_images(images_name_6, images_position_6, slide_index_6)
print(f"{passed} Слайд №6 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 7
# Массив имен изображений с префиксом папки
images_name_7 = [f"{folder_name}_{image}" for image in ["9", "6"]]
images_position_7 = [
    (Inches(0.6), Inches(1.5), Inches(7), Inches(4)),
    (Inches(0.6), Inches(6.3), Inches(7), Inches(4.6)),
]

insert_images(images_name_7, images_position_7, 7)
rename_image(images_name_7[0], "мягкие ткани")
rename_image(images_name_7[1], "костная ткань")
print(f"{passed} Слайд №7 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 8
images_name_8 = [f"{folder_name}_{image}" for image in ["11", "вч", "нч"]]
images_position_8 = [
    (Inches(0.5), Inches(1.1), Inches(7.4), Inches(3.4)),
    (Inches(0.5), Inches(5.1), Inches(7.2), Inches(2.9)),
    (Inches(0.5), Inches(8.5), Inches(7.2), Inches(2.9)),
]

insert_images(images_name_8, images_position_8, 8)
rename_image(images_name_8[0], "ОПТГ")
print(f"{passed} Слайд №8 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 9
# Массив имен изображений с префиксом папки
images_name_9 = [f"{folder_name}_{image}" for image in ["22", "1", "2"]]
images_position_9 = [
    (Inches(0.4), Inches(1.8), Inches(7.6), Inches(3.2)),
    (Inches(1.5), Inches(5.8), Inches(5.9), Inches(2.6)),
    (Inches(1.5), Inches(8.7), Inches(5.9), Inches(2.6)),
]
insert_images(images_name_9, images_position_9, 9)
rename_image(images_name_9[0], "ВНЧС")
rename_image(images_name_9[1], "ВНЧС прав")
rename_image(images_name_9[2], "ВНЧС лев")
print(f"{passed} Слайд №9 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 10
# Размеры и положения областей для изображений
images_name_10_444 = [f"{folder_name}_{image}" for image in ["444"]]
images_name_10 = [f"{folder_name}_{image}" for image in ["33", "44"]]
img_name10_1 = os.path.join(images_name_10_444[0] + ".jpg")

# Пример использования
crop_image(os.path.join(image_folder, images_name_10_444[0] + ".jpg"),
           os.path.join(image_folder, images_name_10_444[0] + ".jpg"), 1200, 1068)

images_position_10 = [
    (Inches(0.5), Inches(7.5), Inches(3.5), Inches(3.5)),
    (Inches(4.1), Inches(7.5), Inches(3.5), Inches(3.5)),
]

prs.slides[10].shapes.add_picture(os.path.join(image_folder, images_name_10_444[0] + ".jpg"), Inches(1.2), Inches(1.4),
                                  Inches(6), Inches(5.5))
insert_images(images_name_10, images_position_10, 10)
rename_image(images_name_10[0], "ТРГ фронт")
rename_image(images_name_10[1], "SMV")
rename_image(images_name_10_444[0], "симметрия")
print(f"{passed} Слайд №10 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 11
# Размеры и положения областей для изображений
images_name_11 = [f"{folder_name}_{image}" for image in ["4", "5", "55"]]
images_position_11 = [
    (Inches(0.4), Inches(1.4), Inches(3.6), Inches(3.6)),
    (Inches(4.2), Inches(1.4), Inches(3.6), Inches(3.6)),

    (Inches(1), Inches(5.8), Inches(6.2), Inches(5.6)),
]

insert_images(images_name_11, images_position_11, 11)
rename_image(images_name_11[0], "ТРГ прав")
rename_image(images_name_11[1], "ТРГ лев")
rename_image(images_name_11[2], "трассированная трг")
print(f"{passed} Слайд №11 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 12
images_name_12 = [f"{folder_name}_{image}" for image in ["12q"]]
images_position_12 = [
    (Inches(0.65), Inches(1.6), Inches(7.1), Inches(8.5)),
]
slide_index_12 = 12
insert_images(images_name_12, images_position_12, slide_index_12)
print(f"{passed} Слайд №12 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд № 13
images_name_13 = [f"{folder_name}_{image}" for image in ["13q"]]
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

print(f"{passed} Слайд №13 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")


# Слайд № 14
images_name_14 = [f"{folder_name}_{image}" for image in ["3_crop", "000"]]
images_position_14 = [
    (Inches(0.6), Inches(8.1), Inches(3.4), Inches(3.3)),
    (Inches(4.2), Inches(8.1), Inches(3.4), Inches(3.3)),
]

apply_crop_to_images([f"{folder_name}_{image}" for image in ["3"]], 880, 900, "_crop")

insert_images(images_name_14, images_position_14, 14)
rename_image([f"{folder_name}_{image}" for image in ["3"]][0], "фронтальный расчет")
rename_image(images_name_14[1], "Yaw ротация нижней челюсти")
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
print(f"{passed} Слайд №14 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд №15
# Массив имен изображений с префиксом папки
images_name_15 = [f"{folder_name}_{image}" for image in ["333", "temp15"]]
images_position_15 = [
    (Inches(0.6), Inches(1.5), Inches(7), Inches(4.7)),
    (Inches(0.6), Inches(7), Inches(7), Inches(4.4)),
]

insert_images(images_name_15, images_position_15, 15)
rename_image(images_name_15[0], "аксиальные срезы")

print(f"{passed} Слайд №15 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд №16
# Массив имен изображений с префиксом папки
images_name_16 = [f"{folder_name}_{image}" for image in ["222", "0"]]
images_position_16 = [
    (Inches(0.6), Inches(1.2), Inches(7), Inches(4)),
    (Inches(2), Inches(7.65), Inches(4.3), Inches(3.9)),
]

insert_images(images_name_16, images_position_16, 16)
rename_image(images_name_16[0], "корональные срезы")
rename_image(images_name_16[1], "воздухоносные пути")

print(f"{passed} Слайд №16 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")


# Слайд №17
# Определяем тенденцию к классу
anb_value = ws1['L42'].value
beta_angle = ws1['L44'].value
wits_appraisal = ws1['L46'].value
sassouni = ws1['L125'].value
apdi_value = ws1['L43'].value
pnsa_value = ws1['C9'].value
jj_value = ws1['C10'].value
sna_value = ws1['C13'].value
ppsn_value = ws1['C14'].value

# Определяем класс в зависимости от значения ANB
anb_trend_class = ""
if anb_value > 4:
    anb_skeletal_class = "II"
elif anb_value < 0:
    anb_skeletal_class = "III"
else:
    anb_skeletal_class = "I"
    # Если класс "I", проверяем дополнительные условия
    if 0 <= anb_value <= 0.4:
        anb_trend_class = "с тенденцией к III классу"
    elif 3.6 <= anb_value <= 4:
        anb_trend_class = "с тенденцией к II классу"

# Определяем класс в зависимости от значения BETA ANGLE
beta_trend_class = ""
if beta_angle > 35:
    beta_skeletal_class = "III"
elif beta_angle < 27:
    beta_skeletal_class = "II"
else:
    beta_skeletal_class = "I"
    # Если класс "I", проверяем дополнительные условия
    if 34.6 <= beta_angle <= 35:
        beta_trend_class = "с тенденцией к III классу"
    elif 27 <= beta_angle <= 27.4:
        beta_trend_class = "с тенденцией к II классу"

# Определяем класс в зависимости от значения Wits Appraisal
has_value = ""
wits_trend_class = ""
if wits_appraisal > 2.1:
    wits_skeletal_class = "II"
    has_value = "наличие"
elif wits_appraisal < -2.9:
    wits_skeletal_class = "III"
    has_value = "наличие"
else:
    wits_skeletal_class = "I"
    has_value = "отсутствие"
    # Если класс "I", проверяем дополнительные условия
    if -2.5 <= wits_appraisal <= -2.9:
        wits_trend_class = "с тенденцией к III классу"
    elif 1.7 <= wits_appraisal <= 2.1:
        wits_trend_class = "с тенденцией к II классу"

# Определяем класс в зависимости от значения Wits Appraisal
sassouni_text = ""
sassouni_trend_class = ""
has_direction = ""
if sassouni > 3:
    sassouni_skeletal_class = "III"
    has_direction = "кзади"
elif sassouni < 0:
    sassouni_skeletal_class = "II"
    has_direction = "кпереди"
else:
    sassouni_skeletal_class = "I"
    # Если класс "I", проверяем дополнительные условия
    has_direction = "кпереди"
    if 2.6 <= sassouni <= 3:
        sassouni_trend_class = "с тенденцией к III классу"
    elif 0.1 <= sassouni <= 0.4:
        sassouni_trend_class = "с тенденцией к II классу"

sassouni_text_not_null = f"""Соотношение челюстей по методике Sassouni говорит за {sassouni_skeletal_class} скелетный класс {sassouni_trend_class} — базальная дуга проходит на {format_with_comma(sassouni)} мм {has_direction} от точки В (N = 0,0 мм ± 3,0 мм)."""
sassouni_text_null = f"""Соотношение челюстей по методике Sassouni говорит за I скелетный класс — базальная дуга проходит через точку B (N = 0,0 мм ± 3,0 мм)."""

if sassouni == 0:
    sassouni_text = sassouni_text_null
else:
    sassouni_text = sassouni_text_not_null

# Определяем класс в зависимости от значения APDI
apdi_trend_class = ""
if apdi_value > 86.4:
    apdi_skeletal_class = "III"
elif apdi_value < 76.4:
    apdi_skeletal_class = "II"
else:
    apdi_skeletal_class = "I"
    # Если класс "I", проверяем дополнительные условия
    if 86 <= apdi_value <= 86.4:
        apdi_trend_class = "с тенденцией к III классу"
    elif 76.4 <= apdi_value <= 76.8:
        apdi_trend_class = "с тенденцией к II классу"

# Определяем класс в зависимости от значения PNSA
pnsa_upper_limit = ws1['D9'].value + 3.5
pnsa_lower_limit = ws1['D9'].value - 3.5
pnsa_trend_class = ""
pnsa_status = ""
if pnsa_value > pnsa_upper_limit:
    pnsa_status = "увеличению"
elif pnsa_value < pnsa_lower_limit:
    pnsa_status = "уменьшению"
else:
    pnsa_status = "норме"
    # Если класс "I", проверяем дополнительные условия
    if 86 <= pnsa_value <= 86.4:
        pnsa_trend_class = "с тенденцией к III классу"
    elif 76.4 <= pnsa_value <= 76.8:
        pnsa_trend_class = "с тенденцией к II классу"

# Определяем класс в зависимости от значения Ширины скелетного базиса (J-J)
jj_upper_limit = ws1['D10'].value + 3
jj_lower_limit = ws1['D10'].value - 3
jj_status = ""
if jj_value > jj_upper_limit:
    jj_status = "расширению"
elif jj_value < jj_lower_limit:
    jj_status = "сужению"
else:
    jj_status = "норме"

# Определяем класс в зависимости от значения SNA
if sna_value > 85:
    sna_status = "прогнатии"
elif sna_value < 79:
    sna_status = "ретрогнатии"
else:
    sna_status = "нормогнатии"

sna_status_uppercase = sna_status.capitalize()

# Определяем класс в зависимости от значения PP/SN
if ppsn_value > 12:
    ppsn_status = "ретроинклинации"
elif ppsn_value < 5:
    ppsn_status = "антеинклинации"
else:
    ppsn_status = "нормоинклинации"

ppsn_status_uppercase = ppsn_status.capitalize()

# Формируем текст, вставляя значения переменных
resume_text1 = f"""
Межапикальный угол (<ANB) – {format_with_comma(anb_value)}˚, что соответствует соотношению челюстей по {anb_skeletal_class} скелетному классу {anb_trend_class} (N = 2,0˚ ± 2,0˚).
Угол Бета (< Beta Angle) – {format_with_comma(beta_angle)}˚, что cоответствует соотношению челюстей по {beta_skeletal_class} скелетному классу {beta_trend_class} (N = 31,0˚ ± 4,0˚).
Параметр Wits (Wits Appraisal.) –  {format_with_comma(wits_appraisal)} мм что указывает на {has_value} диспропорции в расположении апикальных базисов верхней и нижней челюстей в сагиттальной плоскости и говорит за {wits_skeletal_class} скелетный класс {wits_trend_class} (N = -1,0 мм ± 2,0 мм).
{sassouni_text}
Параметр APDI, указывающий на дисплазию развития челюстей в сагиттальной плоскости, равен {format_with_comma(apdi_value)}˚ и говорит за {apdi_skeletal_class} скелетный класс {apdi_trend_class} (N = 81,4˚ ± 5,0˚).
"""
resume_text1_1 = f"""
Длина основания верхней челюсти (PNS-A) – {format_with_comma(pnsa_value)} мм, что соответствует {pnsa_status} (N = {format_with_comma(round(ws1['D9'].value, 1))} мм ± 3,5 мм).
Ширина основания верхней (J-J) челюсти –  {format_with_comma(jj_value)} мм, что соответствует {jj_status} (N = {format_with_comma(round(ws1['D10'].value, 1))} мм ± 3,0 мм):  справа – {format_with_comma(ws1['C11'].value)} мм, слева – {format_with_comma(ws1['C12'].value)} мм (N = {format_with_comma(ws1['D10'].value / 2)} мм ± 1,5 мм).
Положение верхней челюсти по сагиттали  (<SNA) – {format_with_comma(sna_value)}˚, что соответствует {sna_status} (N = 82,0˚ ±  3,0˚).
Положение верхней челюсти по вертикали  (<SN-Palatal Plane) – {format_with_comma(ppsn_value)}˚, что соответствует {ppsn_status} (N = 8,0˚ ± 3,0˚).
Roll ротация отсутствует\  вправо (по часовой стрелке) \ влево (против часовой стрелки).
Yaw ротация отсутствует \ вправо  (по часовой стрелке) \ влево (против часовой стрелки).
"""

# Добавляем текст на слайд
text_left_17 = Inches(0.4)
text_top_17 = Inches(6.7)
text_width_17 = Inches(7.2)
text_height_17 = Inches(5)
name_textbox_17 = prs.slides[17].shapes.add_textbox(text_left_17, text_top_17, text_width_17, text_height_17)
text_frame = name_textbox_17.text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(10.5)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = resume_text1

# Добавляем текст на слайд
name_textbox_17 = prs.slides[17].shapes.add_textbox(Inches(0.4), Inches(8.85), Inches(7.21), Inches(5.1))
text_frame = name_textbox_17.text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(10.5)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = resume_text1_1

print(f"{passed} Слайд №17 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

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

ans_xi_pm = ws1['L113'].value

odi_value = ws1['L120'].value

u1_l1_r = ws1['C31'].value
u1_l1_l = ws1['C32'].value


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

mp_sn_status_uppercase = mp_sn_status.capitalize()

# Определение смещения подбородка
if chin_displacement > 0:
    chin_displacement_status = f"влево на {format_with_comma(round(chin_displacement, 2))} мм"
elif chin_displacement < 0:
    chin_displacement_status = f"вправо на {format_with_comma(round(chin_displacement, 2))} мм"
else:
    chin_displacement_status = "не выявлено"

# Определяем класс в зависимости от значения (N-ANS) / (ANS-Gn)
if ans_quotient > 0.89:
    ans_quotient_status = "негармоничное"
elif ans_quotient < 0.71:
    ans_quotient_status = "негармоничное"
else:
    ans_quotient_status = "гармоничное"

# Определяем класс в зависимости от значения ANS-Xi-Pm
ans_xi_pm_upper_limit = round(ws1['N8'].value, 1) + 5.5
ans_xi_pm_lower_limit = round(ws1['N8'].value, 1) - 5.5

ans_xi_pm_status = ""
if ans_xi_pm > ans_xi_pm_upper_limit:
    ans_xi_pm_status = "увеличению"
elif ans_xi_pm < ans_xi_pm_lower_limit:
    ans_xi_pm_status = "уменьшению"
else:
    ans_xi_pm_status = "норме"

# Определяем класс в зависимости от значения ODI
if odi_value > 79.5:
    odi_value_status = "к глубокой резцовой окклюзии"
elif odi_value < 69.5:
    odi_value_status = "вертикальной резцовой дизокклюзии"
else:
    odi_value_status = "норме"


def get_tooth_status(slant_value, difference, upper_threshold, lower_threshold, tooth_num):
    if slant_value is None:
        return f"Нормальное положение зуба {tooth_num}"
    elif slant_value > upper_threshold:
        return f"Протрузия зуба  {tooth_num} на {difference}˚"
    elif slant_value < lower_threshold:
        return f"Ретрузия зуба  {tooth_num} на {difference}˚"
    else:
        return f"Нормальное положение зуба  {tooth_num}"


slant_r1_1 = ws1['C33'].value
slant_l2_1 = ws1['C34'].value
slant_l3_1 = ws1['C36'].value
slant_r4_1 = ws1['C35'].value

slant_r1_1_dif = format_with_comma(round(ws1['G33'].value, 1)) if ws1['G33'].value is not None else None
slant_l2_1_dif = format_with_comma(round(ws1['G34'].value, 1)) if ws1['G34'].value is not None else None
slant_l3_1_dif = format_with_comma(round(ws1['G36'].value, 1)) if ws1['G36'].value is not None else None
slant_r4_1_dif = format_with_comma(round(ws1['G35'].value, 1)) if ws1['G35'].value is not None else None

r1_1_value_status = get_tooth_status(slant_r1_1, slant_r1_1_dif, 115, 105, 1.1)
l2_1_value_status = get_tooth_status(slant_l2_1, slant_l2_1_dif, 115, 105, 2.1)
l3_1_value_status = get_tooth_status(slant_l3_1, slant_l3_1_dif, 100, 90, 3.1)
r4_1_value_status = get_tooth_status(slant_r4_1, slant_r4_1_dif, 100, 90, 4.1)

# Формирование строки с динамическими данными
u1_pp = f"{r1_1_value_status}. {l2_1_value_status}"
l1_pp = f"{l3_1_value_status}. {r4_1_value_status}"

# Формируем текст, вставляя значения переменных
resume_text2 = f"""
Длина тела нижней челюсти (Go-Me): справа – {format_with_comma(go_me_r_value)} мм, слева – {format_with_comma(go_me_l_value)}  мм (N = {format_with_comma(round(ws1['M59'].value, 1))} мм ± 5,0 мм).
Длина тела нижней челюсти справа {go_me_status}, чем слева на {format_with_comma(round(abs(go_me_r_value - go_me_l_value), 2))} мм.
Длина ветви нижней челюсти (Co-Go) : справа – {format_with_comma(go_go_r_value)}  мм,  слева – {format_with_comma(go_go_l_value)}  мм (N = {format_with_comma(round(ws1['D19'].value, 1))} мм ± 4,0 мм).
Длина ветви нижней челюсти справа {go_go_status}, чем слева на {format_with_comma(round(abs(go_go_r_value - go_go_l_value), 2))} мм.
Гониальный угол (<Ar-Go-Me): справа –  {format_with_comma(ar_go_r_value)}˚,  слева – {format_with_comma(ar_go_l_value)}˚ (N = {format_with_comma(round(ws1['D21'].value, 1))}˚ ± 5,0˚).
Гониальный угол справа {ar_go_status}, чем слева на {format_with_comma(round(abs(ar_go_r_value - ar_go_l_value), 2))}˚.
Ширина базиса нижней челюсти (Md-Md) – {format_with_comma(md_md_value)} мм, что соответствует {md_status} (N = {format_with_comma(round(ws1['D23'].value, 1))} мм ± 3,0 мм).
Положение нижней челюсти по сагиттали  (<SNB) – {format_with_comma(snb_value)}˚, что соответствует {snb_status} (N = 80,0˚ ± 3,0˚).
Положение нижней челюсти по вертикали (<MP-SN) – {format_with_comma(mp_sn_value)}˚, что соответствует {mp_sn_status} (N = 32,0˚ ± 4,0˚).
Смещение подбородка {chin_displacement_status}, \ за счет скелетной ассиметрии.
Roll ротация отсутствует \  вправо (по часовой стрелке) \ влево (против часовой стрелки).
Yaw ротация отсутствует \ вправо  (по часовой стрелке) \ влево (против часовой стрелки).
"""
resume_text3 = f"""
Вертикальное лицевое соотношение (N-ANS/ANS-Gn) {ans_quotient_status} – {format_with_comma(round(ans_quotient, 2))} (N = 0,8 ± 0,09).
Отношение задней высоты лица к передней (S-Go/N-Gn) – {format_with_comma(assessment_growth_type)}% (N = 63,0% ± 2,0%).
Высота нижней трети лица по Ricketts (<ANS-Xi-Pm) – {format_with_comma(ans_xi_pm)}˚, что соответствует {ans_xi_pm_status} (N = {format_with_comma(round(ws1['N8'].value, 1))}˚ ± 5,5˚).
Параметр ODI – {format_with_comma(odi_value)}˚, что соответствует {odi_value_status} (N = 74,5˚ ±  5,0˚).
"""

resume_text4 = f"""
Межрезцовый угол: справа – {format_with_comma(u1_l1_r)}˚, слева – {format_with_comma(u1_l1_l)}˚ (N = 130,0˚ ± 6,0˚).
{u1_pp} (N = 110,0˚± 5,0˚).
{l1_pp} (N = 95,0˚ ± 5,0˚).
"""

width_lower_jaw = ws1['C23'].value
width_upper_jaw = ws1['C10'].value
width_dif_jaw = width_lower_jaw + 5
jaw_dif = abs(width_dif_jaw - width_upper_jaw)
jaw_status = f"составляет {format_with_comma(round(jaw_dif, 1))} мм." if jaw_dif > 0 else "отсутствует."

resume_text5 = f"""
Ширина базиса нижней челюсти – {format_with_comma(width_lower_jaw)} мм. Фактическая ширина базиса верхней челюсти – {format_with_comma(width_upper_jaw)} мм. 
Требуемая ширина базиса верхней челюсти = {format_with_comma(width_dif_jaw)} мм. 
Дефицит ширины скелетного базиса верхней челюсти {jaw_status}
"""

# Добавляем текст на слайд
text_left_18_1 = Inches(0.5)
text_top_18_1 = Inches(0.9)
name_textbox_18_1 = prs.slides[18].shapes.add_textbox(text_left_18_1, text_top_18_1, text_width_17, text_height_17)
text_frame = name_textbox_18_1.text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(10.5)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = resume_text2

# Добавляем текст на слайд
text_left_18_2 = Inches(0.5)
text_top_18_2 = Inches(4.4)
name_textbox_18_2 = prs.slides[18].shapes.add_textbox(text_left_18_2, text_top_18_2, text_width_17, text_height_17)
text_frame = name_textbox_18_2.text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(10.5)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = resume_text3

# Добавляем текст на слайд
text_left_18_3 = Inches(0.5)
text_top_18_3 = Inches(5.73)
name_textbox_18_3 = prs.slides[18].shapes.add_textbox(text_left_18_3, text_top_18_3, text_width_17, text_height_17)
text_frame = name_textbox_18_3.text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(10.5)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = resume_text4

# Добавляем текст на слайд
name_textbox_18_4 = prs.slides[18].shapes.add_textbox(Inches(0.5), Inches(6.9), text_width_17, text_height_17)
text_frame = name_textbox_18_4.text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(10.5)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = resume_text5

print(f"{passed} Слайд №18 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

# Слайд №19
# Верхняя челюсть: пункт 1
if jaw_dif == 0:
    width_basis_lower_jaw = "норме"
elif width_upper_jaw > width_dif_jaw:
    width_basis_lower_jaw = f"расширении базиса верхней челюсти на {format_with_comma(round(jaw_dif, 2))} мм"
else:
    width_basis_lower_jaw = f"сужении базиса верхней челюсти на {format_with_comma(round(jaw_dif, 2))} мм"

# Нижняя челюсть: пункт 3
if go_me_r_value > ws1['D17'].value + ws1['E17'].value:
    go_me_r_status = "увеличена"
elif go_me_r_value < ws1['D17'].value + ws1['E17'].value:
    go_me_r_status = "уменьшена"
else:
    go_me_r_status = "в норме"

# Нижняя челюсть: пункт 3
if go_me_l_value > ws1['D18'].value + ws1['E18'].value:
    go_me_l_status = "увеличена"
elif go_me_l_value < ws1['D18'].value + ws1['E18'].value:
    go_me_l_status = "уменьшена"
else:
    go_me_l_status = "в норме"

# Нижняя челюсть: пункт 4
if go_go_r_value > ws1['D19'].value + ws1['E19'].value:
    go_go_r_status = "увеличена"
elif go_go_r_value < ws1['D19'].value + ws1['E19'].value:
    go_go_r_status = "уменьшена"
else:
    go_go_r_status = "в норме"

# Нижняя челюсть: пункт 4
# 59
if go_go_l_value > ws1['D20'].value + ws1['E20'].value:
    go_go_l_status = "увеличена"
elif go_go_l_value < ws1['D20'].value - ws1['E20'].value:
    go_go_l_status = "уменьшена"
else:
    go_go_l_status = "в норме"

snb_value_finish = ws1['c24'].value

# Определяем класс в зависимости от значения <SNB
if snb_value_finish > 83:
    snb_status_finish = "прогнатия"
elif snb_value_finish < 77:
    snb_status_finish = "ретрогнатия"
else:
    snb_status_finish = "нормогнатия"

snb_status_uppercase = snb_status.capitalize()

incisor_tilt_r1_1 = r1_1_value_status.split("на")[0].strip()
incisor_tilt_l2_1 = l2_1_value_status.split("на")[0].strip()
incisor_tilt_l3_1 = l3_1_value_status.split("на")[0].strip()
incisor_tilt_r4_1 = r4_1_value_status.split("на")[0].strip()

overbite_value = ws1['C29'].value
overjet_value = ws1['C30'].value

# Определяем класс в зависимости от значения Overbite
if overbite_value > 4.4:
    overbite_value_status = f"Глубокая резцовая окклюзия. Вертикальное резцовое перекрытие увеличено до {overbite_value} мм (N = 2,5 мм ± 2,0 мм)."
elif overbite_value < 0.5:
    overbite_value_status = f"Вертикальная резцовая дизокклюзия – {format_with_comma(overbite_value)} мм (N = 2,5 мм ± 2,0 мм)."
else:
    overbite_value_status = f"Вертикальное резцовое перекрытие в норме – {format_with_comma(overbite_value)} мм (N = 2,5 мм ± 2,0 мм)."

# Определяем класс в зависимости от значения Overjet
if overjet_value > 5:
    overjet_value_status = f"Сагиттальная щель – {format_with_comma(round(overjet_value, 1))} мм (N = 2,5 мм ± 2,5 мм)."
elif overjet_value < 0:
    overjet_value_status = f"Обратная сагиттальная щель {format_with_comma(round(overjet_value, 1))} мм (от -0,1 и выше) (N = 2,5 мм ± 2,5 мм)."
else:
    overjet_value_status = f"Сагиттальное резцовое перекрытие в норме – {format_with_comma(round(overjet_value, 1))} мм (N = 2,5 мм ± 2,5 мм)."

slide20_text1 = f"""
1. Скелетный III класс обусловленный диспропорцией расположения апикальных 
    базисов челюстей в сагиттальном направлении. Зубоальвеолярная форма 
    дистальной \ мезиальной окклюзии.
2. Мезофациальный тип строения лицевого отдела черепа. 
3. Нейтральный тип роста с тенденцией к вертикальному\ горизонтальному росту.
4. Высота нижней трети лица по Ricketts  в {ans_xi_pm_status}.
5. Профиль лица  выпуклый. 
6. Ретроположение верхней и нижней губы относительно 
    эстетической плоскости Ricketts. 
7. Сужение и уменьшение объема воздухоносных путей. Сужения и уменьшения 
    объема воздухоносных путей не выявлено. 
8. Нормальное \ Переднее \ Заднее положение правой \ левой суставной головки
    височно-нижнечелюстного сустава.
9. Скелетный возраст соответствует IIIVI стадии созревания шейных позвонков.
"""

slide20_text2 = f"""
1. Ширина базиса верхней челюсти в {width_basis_lower_jaw} 
    (Penn анализ).
2. {sna_status_uppercase} верхней челюсти. {ppsn_status_uppercase} верхней челюсти.
3. Ротация верхней челюсти в Roll \Yaw плоскости вправо (по часовой стрелке)
    \влево (против часовой стрелки).
"""
slide20_text3 = f"""
1. Ширина базиса нижней челюсти в {md_status} базиса нижней челюсти
    относительно возрастной нормы.
2. {snb_status_uppercase} нижней челюсти. {mp_sn_status_uppercase} нижней челюсти.
3. Длина тела нижней челюсти справа {go_me_r_status}. Длина тела нижней челюсти слева
    {go_me_l_status}.
4. Длина ветви нижней челюсти справа {go_go_r_status}. Длина ветви нижней челюсти 
    слева {go_go_l_status}.
5. Смещение подбородка {chin_displacement_status}.
6. Ротация нижней челюсти в Roll \Yaw плоскости вправо (по часовой стрелке) 
    \влево (против часовой стрелки).
"""

slide20_text4 = f"""
1. Межрезцовая линия на верхней челюсти смещена относительно
    срединно сагиттальной линии на 1,2 мм вправо, на нижней челюсти смещена 
    на 1,2 мм влево.
2. Сужение верхнего зубного ряда в области клыков, премоляров, моляров.
    Сужение нижнего зубного ряда в области клыков, моляров, премоляров.
3. Длина фронтального участка верхнего зубного ряда в норме , нижнего зубного 
    ряда в норме.
4. {incisor_tilt_r1_1}. {incisor_tilt_l2_1}. {incisor_tilt_l3_1}.
    {incisor_tilt_r4_1}
5. {overbite_value_status}
6. {overjet_value_status}
7. Глубина кривой Шпее в увеличена справа \ слева.
"""

# Добавляем текст на слайд
text_frame = prs.slides[19].shapes.add_textbox(Inches(0.6), Inches(0.6), Inches(7), Inches(5)).text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(11)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = slide20_text1

# Добавляем текст на слайд
text_frame = prs.slides[19].shapes.add_textbox(Inches(0.6), Inches(3.6), Inches(7), Inches(5)).text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(11)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = slide20_text2

# Добавляем текст на слайд
text_frame = prs.slides[19].shapes.add_textbox(Inches(0.6), Inches(5), Inches(7), Inches(5)).text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(11)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = slide20_text3

# Добавляем текст на слайд
text_frame = prs.slides[19].shapes.add_textbox(Inches(0.6), Inches(7.35), Inches(7), Inches(5)).text_frame
text_frame.word_wrap = True
paragraph = text_frame.add_paragraph()
paragraph.font.size = Pt(11)
paragraph.font.bold = False
paragraph.font.name = "Montserrat"
paragraph.text = slide20_text4

print(f"{passed} Слайд №19 сформирован 🎉")
print("<------------------------------------------------------------------------------------------->")

if folder_name:
    save_folder = os.path.join(work_folder, folder_name)
    prs.save(os.path.join(save_folder, f"{folder_name}.pptx"))


def extract_text_from_slides(prs, slide_indices):
    """
    Функция для извлечения текста с указанных слайдов презентации.
    Args:
        prs (Presentation): Объект презентации PowerPoint.
        slide_indices (list): Список индексов слайдов для извлечения текста.
    Returns:
        str: Текст с указанных слайдов.
    """
    text = ""
    for idx in slide_indices:
        slide = prs.slides[idx]
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text


# Индексы слайдов, с которых нужно извлечь текст
slide_indices_to_extract = [17, 18, 19]  # Пример: извлечение текста с первых трех слайдов

# Извлечение текста с указанных слайдов
extracted_text = extract_text_from_slides(prs, slide_indices_to_extract)

# Путь к файлу, в который будет сохранен текст
output_file_path = os.path.join(os.path.join(work_folder, folder_name), f"{folder_name}.txt")


def save_text_to_file(text, file_path):
    """
    Функция для сохранения текста в файл.
    Args:
        text (str): Текст для сохранения.
        file_path (str): Путь к файлу, в который будет сохранен текст.
    """
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(text)

# Сохранение текста в файл
save_text_to_file(extracted_text, output_file_path)

print(f"😊 Текст успешно извлечен с выбранных слайдов и сохранен в файл {output_file_path}.")
