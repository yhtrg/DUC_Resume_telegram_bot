import requests
import json
import os
import textract
import logging
import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from docx import Document as DocxDocument
from aiogram.types.input_file import FSInputFile
from docx.shared import Pt
from striprtf.striprtf import rtf_to_text
from spire.doc import *
from spire.doc.common import *


logging.basicConfig(level=logging.INFO)

encoding = 'UTF-8'
API_URL = "http://178.212.132.7:3000/api/v1/prediction/1398cdef-26c1-42c8-8f0e-3715b905d0ac"
TOKEN = '7096081921:AAHX23mpdT1pe4yZJfzBxnNM10xroSkB8HI'

DOWNLOAD_FOLDER = 'downloads'
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

dp = Dispatcher()
bot = Bot(token=TOKEN)

def query(payload):
    response = requests.post(API_URL, json=payload)
    return response.json()

def download_word(data):
    doc = DocxDocument('resume.docx')

    for paragraph in doc.paragraphs:
        if '[Имя Фамилия]' in paragraph.text:
            run = paragraph.runs[0]
            run.text = run.text.replace('[Имя Фамилия]', data["name"])
            run.font.size = Pt(22)
            run.font.name = 'Times New Roman'
            run.bold = True

        if '[Должность]' in paragraph.text:
            run = paragraph.runs[0]
            run.text = run.text.replace('[Должность]', data['position'])
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run.bold = True

        if 'Пол:' in paragraph.text:
            run = paragraph.runs[0]
            run = paragraph.add_run(f" {data['gender']}")
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'

        if 'Возраст:' in paragraph.text:
            run = paragraph.runs[0]
            run = paragraph.add_run(f' {data["age"]}')
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'

        if '[Город проживания]' in paragraph.text:
            run = paragraph.runs[0]
            if data['city']:
                run.text = run.text.replace('[Город проживания]', data['city'])
            else:
                run.text = run.text.replace('[Город проживания]', '')
                run.text = run.text.replace('Проживает: г.', '')
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'

        if '[Список технологий]' in paragraph.text:
            run = paragraph.runs[0]
            run.text = run.text.replace('[Список технологий]', (data['tech_stack']))
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'

        if '[Количество лет и месяцев]' in paragraph.text:
            run = paragraph.runs[0]
            run.text = run.text.replace('[Количество лет и месяцев]', data['all_work_experience'])
            run.font.size = Pt(14)
            run.font.name = 'Times New Roman'
            run.bold = True

            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            for experience in data['work_experience']:
                row_cells = table.add_row().cells
                row_cells[0].text = f"{experience['start_date']} - {experience['end_date']}"
                row_cells[1].text = f"{experience['current_position']}\n\n{experience['function']}"
                
            for row in table.rows:
                for cell in row.cells:
                    paragraphs = cell.paragraphs
                    for paragraph in paragraphs:
                        for run in paragraph.runs:
                            font = run.font
                            font.size= Pt(12)
                            font.name = 'Times New Roman'

        if 'Образование' in paragraph.text:
            run = paragraph.runs[0]
            for edu in data['education']:
                run = paragraph.add_run(f"\n- {edu['year']}, {edu['university']}, {edu['faculty']}, {edu['specialty']};")
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'

        if '[Информация о себе]' in paragraph.text:
            run = paragraph.runs[0]
            if data["about"]:
                run.text = run.text.replace('[Информация о себе]', data["about"])
            else:
                run.text = run.text.replace('[Информация о себе]', '')
                run.text = run.text.replace('О себе:', '')
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'

    file_name = f"{data['name'].replace(' ', '_')}_Резюме.docx"
    file_path = os.path.join(DOWNLOAD_FOLDER, file_name)
    doc.save(file_path)
    return file_name

@dp.message(CommandStart())
async def start_command(message: types.Message):
    await message.reply("Привет! Отправь мне документ, и я создам резюме с ним.")

def process_document(file_path):
    response = ''
    if '.pdf' in file_path:
        response = textract.process(file_path).decode(encoding)

    elif '.docx' in file_path:
        doc = DocxDocument(file_path)
        for paragraph in doc.paragraphs:
            response += paragraph.text + '\n'

        if doc.tables:
            for table in doc.tables:
                for row in table.rows:
                    row_text = [cell.text for cell in row.cells]
                    response += ' '.join(row_text)

    elif '.rtf' in file_path:
        with open(file_path) as infile:
            content = infile.read()
            response = rtf_to_text(content)

    elif '.doc' in file_path:
        output_file = file_path[:-5] + '.docx'
        document = Document()
        document.LoadFromFile(file_path)
        document.SaveToFile(output_file, FileFormat.Docx2016)
        response = textract.process(output_file).decode(encoding)
        os.remove(output_file)
    return response

@dp.message()
async def handle_document(message: types.Message):
    user_id = message.from_user.id
    input_document = message.document
    input_file_name = input_document.file_name
    input_file_path = os.path.join(DOWNLOAD_FOLDER, input_file_name)
    input_file_id = input_document.file_id
    input_file = await bot.get_file(input_file_id)
    await bot.download_file(input_file.file_path, input_file_path)
    await message.reply("Документ получен! Обрабатываю файл...")
    response = process_document(input_file_path)
    output_flowise = query({"question": response})
    data = output_flowise['text'][7:-3]
    data = json.loads(data)
    output_file_name = download_word(data)
    output_document = FSInputFile(f'downloads/{output_file_name}')
    await bot.send_document(user_id, output_document)
    os.remove(os.path.join(DOWNLOAD_FOLDER, input_file_name))
    os.remove(os.path.join(DOWNLOAD_FOLDER, output_file_name))

async def main() -> None:
    await dp.start_polling(bot)

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    asyncio.run(main())
