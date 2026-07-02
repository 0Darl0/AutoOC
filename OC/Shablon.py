import openpyxl
import textwrap
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from io import BytesIO
from reportlab.pdfbase.ttfonts import TTFont
import os

pdfmetrics.registerFont(TTFont('NunitoB', 'Font/Nunito-Bold.ttf'))
pdfmetrics.registerFont(TTFont('Nunito', 'Font/Nunito-Regular.ttf'))


def write_pdf(name_shablon, fio, count_web, less, addict_text=None):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(1250, 1404))
    c.setFont("NunitoB", 32)
    if count_web[0] == count_web[1]:
        c.drawString(802, 952, f"{count_web[0]}/{count_web[1]}")
    else:
        c.drawString(802, 952, f"{count_web[0]}")
    c.setFont("Nunito", 34)
    c.drawString(505, 1100, f"{fio}")
    c.setFont("Nunito", 32)

    if addict_text:
        text = addict_text
    elif count_web[0] > 10:
        text = "Ты отлично поработал, результат хороший - это уже победа. Но до твоего реального максимума ещё есть пара шагов, и этот месяц как раз для них. Не расслабляйся, добивай те самые последние слабые места - и тогда будет не просто «хорошо», а по-настоящему круто."
    elif 6 <= count_web[0] <= 9:
        text = "Молодец, что начал и не бросил, но реально сделана только половина, и это честный срез, а не приговор. За месяц предстоит выложиться по полной: бери оставшиеся темы и проходи их в ускоренном темпе, без поблажек. Просто теперь без пропусков и всё получится."
    else:
        text = "Ты старался, это плюс, но результат пока ниже среднего - пробелы серьёзные, и это факт. За месяц предстоит выложиться на максимум: никаких послаблений, но и без паники - просто каждый день разбирай по одной сложной теме, пока не начнёт щёлкать. Ты можешь больше, чем показываешь, просто теперь без пропусков"
    text = "Курс закончился - и это уже победа. Ты прошёл этот путь, даже когда было трудно. Сделал большое дело. Сейчас главное - не выгореть и не сломать себя паникой. Высыпайся, решай по чуть-чуть для уверенности, честно признавай свои слабые места без страха. Ты уже готов намного больше, чем кажется. Просто продолжай в своём темпе и дойди до финиша спокойно. С твоим завершением курса!!"
    stroki = [""]
    inumerator_stroki = 0
    for i in text.split():
        if len(stroki[inumerator_stroki] + i + " ") <= 60:
            stroki[inumerator_stroki] += i + " "
        else:
            inumerator_stroki += 1
            stroki.append("")
            stroki[inumerator_stroki] += i + " "
    y = 550
    for i in stroki:
        c.drawString(150, y, f"{i}")
        y -= 45
    c.save()
    packet.seek(0)
    reader = PdfReader(name_shablon)
    writer = PdfWriter()
    watermark = PdfReader(packet)
    for i in range(len(reader.pages)):
        if i == 1 and count_web[0] == 0:
            ...
        else:
            p = reader.pages[i]
            if i == 1:
                p.merge_page(watermark.pages[0])
            writer.add_page(p)
    with open(f"Result/{fio}_{less}.pdf", "wb") as output_file:
        writer.write(output_file)


def read_file(name):
    wb = openpyxl.load_workbook(name)
    page = wb.active
    for row in page.iter_rows(min_row=2, values_only=True):

        if "БИО ГК" in row[5]:
            balls = [int(x) for x in row[3].split()]
            write_pdf("Shablon/ГК БИО.pdf", row[0], balls, "био", addict_text=row[6])
        if "БИО ПГК" in row[5]:
            balls = [int(x) for x in row[3].split()]
            write_pdf("Shablon/ПГК БИО.pdf", row[0], balls, "био", addict_text=row[6])


    wb.close()


read_file("Data.xlsx")
