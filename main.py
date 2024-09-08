import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Загрузка данных из Google Sheets
def load_google_sheet(s_id, s_range):
    creds = Credentials.from_service_account_file('service.json')
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=s_id, range=s_range).execute()
    return result.get('values', [])

# Установка стиля для документа
def set_document_style(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    paragraph_format = style.paragraph_format
    paragraph_format.line_spacing = 1.5
    paragraph_format.first_line_indent = Cm(1)
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.sections[0].left_margin = Cm(2)
    doc.sections[0].right_margin = Cm(2)
    doc.sections[0].top_margin = Cm(2)
    doc.sections[0].bottom_margin = Cm(2)

# Функция для конвертации полного ФИО в формат "Фамилия И.О."
def convert_to_initials(full_name):
    parts = full_name.split()
    if len(parts) == 3:
        surname, name, patronymic = parts
        return f"{surname} {name[0]}.{patronymic[0]}."
    elif len(parts) == 2:
        surname, name = parts
        return f"{surname} {name[0]}."
    else:
        return full_name

# Генерация программы конференции с форматированием
def generate_conference_program(student_data, tech_data):
    doc = docx.Document()
    set_document_style(doc)

    # Первая строка (жирный шрифт)
    first_paragraph = doc.add_paragraph(
        'Форма представления материалов для программы 78 МСНК ГУАП',
        style='Normal'
    )
    first_paragraph.runs[0].bold = True
    first_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Секция кафедры
    section_heading = doc.add_paragraph()
    run1 = section_heading.add_run(f" {' ' * 4} Секция каф. ")
    run1.bold = True
    run1.italic = True
    run2 = section_heading.add_run(f"{tech_data[0][0]}. {tech_data[0][1]}")
    run2.bold = True
    run2.italic = True

    # Научный руководитель
    doc.add_paragraph(
        f" {' ' * 10} Научный руководитель секции - {tech_data[0][2]}",
        style='Normal'
    )

    doc.add_paragraph(
        f" {' ' * 10} {tech_data[0][3]}",
        style='Normal'
    )

    # Зам. научного руководителя
    doc.add_paragraph(
        f" {' ' * 10} Зам. научного руководителя секции - {tech_data[0][6]}",
        style='Normal'
    )

    doc.add_paragraph(
        f" {' ' * 10} {tech_data[0][7]}",
        style='Normal'
    )
    
    # Преобразуем значения в столбце 8 в числа и находим максимальное
    max_value = max([int(row[8]) for row in student_data if row[8].isdigit()])
    
    for cur_num in range(1, max_value + 1):
        
        # Заседание
        session_heading = doc.add_paragraph(f'Заседание {str(cur_num)}', style='Normal')
        session_heading.runs[0].bold = True
        
        # Дата и время + адрес (в одной строке)
        session_info = doc.add_paragraph()
        session_info.add_run(f"{tech_data[cur_num - 1][11]}{' ' * 35}Санкт-Петербург, ул. Большая Морская, д. 67,")

        # Лит. А, ауд. у правого края (в новой строке)
        room_info = doc.add_paragraph()
        run = room_info.add_run(f"{' ' * 75} лит. А, ауд. {tech_data[cur_num - 1][12]}")

        # Список участников с темами
        participant_num = 1
        for row in student_data:
            if int(row[8]) == cur_num:
                initials = convert_to_initials(row[1])  # Фамилия с инициалами
                doc.add_paragraph(f'{participant_num}. {initials}', style='Normal')
                doc.add_paragraph(f'{row[2]}', style='Normal')  # Тема
                participant_num += 1
                
    doc.save('Программа конференции.docx')

def generate_conference_report(student_data, tech_data):
    doc = docx.Document()
    set_document_style(doc)

    # Первая строка (жирный шрифт)
    first_paragraph = doc.add_paragraph(
        'Отчёт о конференции 78 МСНК ГУАП',
        style='Normal'
    )
    first_paragraph.runs[0].bold = True
    first_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Секция кафедры
    section_heading = doc.add_paragraph()
    run1 = section_heading.add_run(f" {' ' * 4} Секция каф. ")
    run1.bold = True
    run1.italic = True
    run2 = section_heading.add_run(f"{tech_data[0][0]}. {tech_data[0][1]}")
    run2.bold = True
    run2.italic = True

    # Преобразуем значения в столбце 8 в числа и находим максимальное
    max_value = max([int(row[8]) for row in student_data if row[8].isdigit()])
    
    for cur_num in range(1, max_value + 1):
        
        # Заседание
        session_heading = doc.add_paragraph(f'Заседание {str(cur_num)}', style='Normal')
        session_heading.runs[0].bold = True
        
        # Дата и время + адрес (в одной строке)
        session_info = doc.add_paragraph()
        session_info.add_run(f"{tech_data[cur_num - 1][11]}{' ' * 35}Санкт-Петербург, ул. Большая Морская, д. 67,")

        # Лит. А, ауд. у правого края (в новой строке)
        room_info = doc.add_paragraph()
        run = room_info.add_run(f"{' ' * 75} лит. А, ауд. {tech_data[cur_num - 1][12]}")

        doc.add_paragraph(
            f"Научный руководитель секции - {tech_data[0][3]} {convert_to_initials(tech_data[0][2])}",
            style='Normal'
        )

        doc.add_paragraph(
            f"Список докладов",
            style='Normal'
        )

        # Таблица для списка докладов
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '№ п/п'
        hdr_cells[1].text = 'ФИО докладчика, название доклада'
        hdr_cells[2].text = 'Статус (магистр/студент)'
        hdr_cells[3].text = 'Решение'

        # Выравнивание текста в заголовке таблицы по центру
        for cell in hdr_cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.first_line_indent = Cm(0)  # Убираем красную строку

        # Заполнение таблицы
        participant_num = 1
        for row in student_data:
            if int(row[8]) == cur_num:
                initials = row[1]  # Фамилия с инициалами
                status = f"{row[4]} Гр. № {row[5]}"
                recommendation = "опубликовать доклад в сборнике МСНК" if row[9] == "1" else "доклад плохо подготовлен"

                row_cells = table.add_row().cells
                row_cells[0].text = str(participant_num)
                row_cells[1].text = f"{initials}\n{row[2]}"  # ФИО и название доклада
                row_cells[2].text = status  # Статус
                row_cells[3].text = recommendation  # Рекомендация

                # Выравнивание текста в ячейках по левому краю
                for paragraph in row_cells[1].paragraphs + row_cells[2].paragraphs + row_cells[3].paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    paragraph.paragraph_format.first_line_indent = Cm(0)  # Убираем красную строку

                participant_num += 1

        # Добавление пустой строки после таблицы
        doc.add_paragraph()

    # Добавление строки "Подпись научного руководителя секции"
    doc.add_paragraph("Подпись научного руководителя секции", style='Normal')

    doc.save('Отчёт о конференции.docx')


def generate_conference_list(student_data, tech_data):
    doc = docx.Document()
    set_document_style(doc)

    # Первая строка (жирный шрифт)
    first_paragraph = doc.add_paragraph(
        'Список представляемых к публикации докладов',
        style='Normal'
    )
    first_paragraph.runs[0].bold = True
    first_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Секция кафедры
    doc.add_paragraph(f"Кафедра {tech_data[0][1]}")
    doc.add_paragraph(tech_data[0][2])
    doc.add_paragraph(f"e-mail: {tech_data[0][4]}")
    doc.add_paragraph(f"тел.: {tech_data[0][5]}")

    # Добавление списка студентов с оценкой "1" в столбце J
    for row in student_data:
        if row[9] == "1":  # Если в столбце J (9-й столбец) стоит "1"
            # Создание абзаца для фамилии с инициалами и темы
            combined_paragraph = doc.add_paragraph()
            
            # Добавляем фамилию с инициалами курсивом
            name_run = combined_paragraph.add_run(f"{convert_to_initials(row[1])} ")
            name_run.italic = True
            
            # Добавляем тему обычным текстом в ту же строку
            combined_paragraph.add_run(f"{row[2]}")

    # Добавление строки "Руководитель УНИДС"
    doc.add_paragraph("\n" * 2)  # Добавление двух пустых строк перед подписью
    doc.add_paragraph(f"Руководитель УНИДС {' ' * 40}{convert_to_initials(tech_data[0][2])}")


    doc.save('Список представляемых к публикации докладов.docx')

# Пример использования
if __name__ == "__main__":
    student_sheet_id = '1szrnHnYN1LOLG9V8eJeye3YTZY4Ka_FPPZfWjQ5_zlw'
    tech_sheet_id = '1MROr3Pw7nMG2vYW_AeqIy2q9FTF7URD3b24tyrBYWgE'
    student_range = 'Sheet1!A2:L'  # Диапазон данных в Google Sheets, включая столбцы для всех заседаний
    tech_range = 'Sheet1!A2:M'
    
    student_data = load_google_sheet(student_sheet_id, student_range)
    tech_data = load_google_sheet(tech_sheet_id, tech_range)
    
    # CLI для выбора типа документа
    print("Какой документ хотите составить?")
    print("1. Программа конференции")
    print("2. Отчёт о конференции")
    print("3. Список представляемых к публикации докладов")
    print("0. Выйти")
    while True:
        document_type = input("Введите номер документа (1, 2 или 3): ")
        if document_type == '1':
            generate_conference_program(student_data, tech_data)
            print("Сгенерирована программа конференции.")
        elif document_type == '2':
            generate_conference_report(student_data, tech_data)
            print("Сгенерирован отчет о конференции.")
        elif document_type == '3':
            generate_conference_list(student_data, tech_data)
            print("Сгенерирован список представляемых к публикации докладов")
        elif document_type == '0':
            print("Завершение программы")
            break
        else:
            print("Ошибка: Неправильный формат ввода.")

