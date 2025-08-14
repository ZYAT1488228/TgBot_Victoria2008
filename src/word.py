from docx import Document
from datetime import date
from docx.shared import Pt

from io import BytesIO
import os

def make_pass(
    names: list[str],
    birth_dates: list[str],
    passport_numbers: list[str],
    start_date: date,
    end_date: date,
    created_at: date,
    auto_model: str,
    auto_plates: str,
    num: str,
    passport_issued_by: str = '',
    passport_id_code: str = ''
):
    doc = Document('./templates/template.docx')
    
    start_date = start_date.strftime("%d.%m.%Y")
    end_date = end_date.strftime("%d.%m.%Y")
    created_at = created_at.strftime("%d.%m.%Y")

    person_data = ""
    for i in range(len(names)):
        person_data += f"ПІБ: {names[i]}\n"
        person_data += f"Дата народження: {birth_dates[i]}\n"
        person_data += f"№ Паспорта: {passport_numbers[i]}\n\n"

    # Заменяем плейсхолдер <person_data>
    for paragraph in doc.paragraphs:
        if '<person_data>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<person_data>', person_data.strip())

    # Заменяем другие плейсхолдеры
    for paragraph in doc.paragraphs:
        if '<date_start>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<date_start>', start_date)
        if '<date_end>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<date_end>', end_date)
        if '<todaysdate>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<todaysdate>', created_at)
        if '<automodel>' in paragraph.text:
            if auto_model:
                paragraph.text = paragraph.text.replace('<automodel>', f'Авто: {auto_model}')
            else:
                paragraph.text = paragraph.text.replace('<automodel>', '')
        if '<autoplates>' in paragraph.text:
            if auto_plates:
                paragraph.text = paragraph.text.replace('<autoplates>', f'Гос.Номер: {auto_plates}')
            else:
                paragraph.text = paragraph.text.replace('<autoplates>', '')
        if '<number>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<number>', num)
        if '<passport_issued_by>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<passport_issued_by>', passport_issued_by)
        if '<passport_id_code>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<passport_id_code>', passport_id_code)

        for run in paragraph.runs:
            run.font.size = Pt(16)
            run.font.name = 'Times New Roman'

    new_word_file = fr'{num}_{names[0]}_{start_date}_{end_date}.docx'
    
    file_stream = BytesIO()
    
    doc.save(file_stream)

    return file_stream, new_word_file