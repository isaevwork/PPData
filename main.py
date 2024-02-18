import os
import warnings
from openpyxl import load_workbook
from pptx import Presentation

# Путь к папке work
work_folder = os.path.join(os.getenv('USERPROFILE'), 'Downloads', 'work')


# Список для хранения путей к файлам Excel
excel_files = []

# Глобальная переменная для хранения имени папки
folder_name = None


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
        book = load_workbook(filename=excel_file_path)
        print(f"Файл Excel {os.path.basename(excel_file_path)} найден в папке {os.path.dirname(excel_file_path)}.")

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

if folder_name:
    prs.save(os.path.join(work_folder, f"{folder_name}.pptx"))

