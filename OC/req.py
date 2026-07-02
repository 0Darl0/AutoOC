from requests import *
from OC import Create_Data_xlsx
import openpyxl
import datetime
import datas

dict_utf = {
'u0410': 'А', 'u0430': 'а',
'u0411': 'Б', 'u0431': 'б',
'u0412': 'В', 'u0432': 'в',
'u0413': 'Г', 'u0433': 'г',
'u0414': 'Д', 'u0434': 'д',
'u0415': 'Е', 'u0435': 'е',
'u0401': 'Ё', 'u0451': 'ё',
'u0416': 'Ж', 'u0436': 'ж',
'u0417': 'З', 'u0437': 'з',
'u0418': 'И', 'u0438': 'и',
'u0419': 'Й', 'u0439': 'й',
'u041a': 'К', 'u043a': 'к',
'u041b': 'Л', 'u043b': 'л',
'u041c': 'М', 'u043c': 'м',
'u041d': 'Н', 'u043d': 'н',
'u041e': 'О', 'u043e': 'о',
'u041f': 'П', 'u043f': 'п',
'u0420': 'Р', 'u0440': 'р',
'u0421': 'С', 'u0441': 'с',
'u0422': 'Т', 'u0442': 'т',
'u0423': 'У', 'u0443': 'у',
'u0424': 'Ф', 'u0444': 'ф',
'u0425': 'Х', 'u0445': 'х',
'u0426': 'Ц', 'u0446': 'ц',
'u0427': 'Ч', 'u0447': 'ч',
'u0428': 'Ш', 'u0448': 'ш',
'u0429': 'Щ', 'u0449': 'щ',
'u042a': 'Ъ', 'u044a': 'ъ',
'u042b': 'Ы', 'u044b': 'ы',
'u042c': 'Ь', 'u044c': 'ь',
'u042d': 'Э', 'u044d': 'э',
'u042e': 'Ю', 'u044e': 'ю',
'u042f': 'Я', 'u044f': 'я',
}

courses_id = {"БиологияПолугодовойкурс": 169,"БиологияГодовойкурс": 159,
              "БиологияЭкспресскурс": 179,"РусскийязыкПолугодовойкурс": 171, "РусскийязыкГодовойкурс": 161,
              "РусскийязыкЭкспресскурс": 180, "ПрофильнаяматематикаГодовойкурс": 162, "ХимияГодовойкурс": 160,
              "ХимияПолугодовойкурс": 0, "ХимияЭкспресскурс": 0}


person = []
session = Session()
link = datas.links
data = datas.data
header = datas.headers
responce = session.post(link, data=data, headers=header).text
link = "https://backend.neofamily.ru/api/v2/curator/students?limit=100&offset=1&sort[latest_contact_at]=asc&"
header = datas.headers_2
responce = session.get(link, headers=header).text
data = responce.split('"data":')[1].split(',"pagination"')[0]
fio_curs = []
while "first_name" in data:
    fio_curs.append([])
    name = data.split('"first_name":"')[1]
    name = name[:name.index('"')]

    if 'u0' in name:
        name = name.replace("\\u", " ")
        name = name.split()
        norm_name = ''
        for i in range(len(name)):
            norm_name += dict_utf["u"+name[i]]
        name = norm_name
    fio_curs[-1].append(name)
    data = data.replace('first_name":"', " ",1)

for i in range(len(fio_curs)):
    last_name = data.split('"last_name":"')[1]
    last_name = last_name[:last_name.index('"')]

    if 'u0' in last_name:
        last_name = last_name.replace("\\u", " ")
        last_name = last_name.split()
        norm_last_name = ''
        for j in range(len(last_name)):
            norm_last_name += dict_utf["u"+last_name[j]]
        last_name = norm_last_name
    fio_curs[i].append(last_name)
    data = data.replace('last_name":"', " ",1)
for i in range(len(fio_curs)):
    fio_curs[i].append([])
    courses = data.split('"name":"')[1]
    if "]" not in courses:
        flag = True
    else: flag = False
    courses = courses[:courses.index('"')]

    if 'u0' in courses:
        courses = courses.replace("\\u", " ")
        courses = courses.replace('.', '').replace('-', '')
        courses = courses.split()
        norm_courses = ''
        for j in range(len(courses)-1):
            norm_courses += dict_utf["u"+courses[j]]
        courses = norm_courses
        fio_curs[i][2].append(courses)
        data = data.replace('name":"', " ", 1)
    if flag:
        courses = data.split('"name":"')[1]
        if "]" not in courses:
            flag = True
        else:
            flag = False
        courses = courses[:courses.index('"')]
        if 'u0' in courses:
            courses = courses.replace("\\u", " ")
            courses = courses.replace('.', '').replace('-', '')
            courses = courses.split()
            norm_courses = ''
            for j in range(len(courses) - 1):
                norm_courses += dict_utf["u" + courses[j]]
            courses = norm_courses
            fio_curs[i][2].append(courses)
            data = data.replace('name":"', " ", 1)
        if flag:
            courses = data.split('"name":"')[1]
            courses = courses[:courses.index('"')]
            if 'u0' in courses:
                courses = courses.replace("\\u", " ")
                courses = courses.replace('.', '').replace('-', '')
                courses = courses.split()
                norm_courses = ''
                for j in range(len(courses) - 1):
                    norm_courses += dict_utf["u" + courses[j]]
                courses = norm_courses
            data = data.replace('name":"', " ", 1)
            fio_curs[i][2].append(courses)
for i in range(len(fio_curs)):
    id_stud = data.split('"id":')[1]
    fio_curs[i].append(id_stud[:id_stud.index('"')-1])
    data = data.replace('"id":', " ",1)

for i in range(len(fio_curs)):
    fio_curs[i].append(["","","",""])
    id_stud = fio_curs[i][3]
    for j in fio_curs[i][2]:
        link = f"https://backend.neofamily.ru/api/v2/curator/students/{id_stud}/courses/{courses_id[j]}/webinars?limit=8&offset=1&sort[date]=desc&filter[date_from]={datetime.date.today() - datetime.timedelta(days=20)}+00:00:00&filter[date_to]={datetime.date.today()}+23:59:59"
        responce = session.get(link, headers=header).text
        try:
            count_web = responce.split('"counters":')[1].split('"completed_webinars":')[1][:responce.split('"counters":')[1].split('"completed_webinars":')[1].index('}')]
            total = responce.split('"pagination":')[1].split('"total":')[1][:responce.split('"pagination":')[1].split('"total":')[1].index(',')]
        except:
            count_web,total = 0,0
        if "Биология" in j:
            fio_curs[i][4][2] = f'{count_web} {total}'
        if "Русский" in j:
            fio_curs[i][4][0] = f'{count_web} {total}'
        if "математика" in j:
            fio_curs[i][4][1] = f'{count_web} {total}'
        if "Химия" in j:
            fio_curs[i][4][3] = f'{count_web} {total}'

wb = openpyxl.Workbook()
page = wb.active
page.append(["Ник на платформе", "Вебы р", "Вебы м", "Вебы б", "Вебы по х", "Курсы", "Уникальный текст"])
for i in range(len(fio_curs)):
    Create_Data_xlsx.add(page, (fio_curs[i][0] + " " +fio_curs[i][1]), fio_curs[i][2], web_r=f"{fio_curs[i][4][0]}",web_b=fio_curs[i][4][2], web_h=fio_curs[i][4][3], web_m=fio_curs[i][4][1])
wb.save("Data.xlsx")


