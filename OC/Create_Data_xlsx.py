import openpyxl


def add(page, name, *curses ,web_r="", web_m="", web_b="", web_h=""):
    row = [name, web_r, web_m, web_b, web_h]
    cell = ""
    for i in curses:
        if "БиологияГодовойкурс" in i: cell+= "БИО ГК,"
        if "БиологияПолугодовойкурс" in i: cell+= "БИО ПГК,"
        if "БиологияЭкспресскурс" in i: cell += "БИО ЭКС,"
        if "РусскийязыкПолугодовойкурс" in i: cell += "РУС ПГК,"
        if "РусскийязыкГодовойкурс" in i: cell += "РУС ГК,"
        if "РусскийязыкЭкспресскурс" in i: cell += "РУС ЭКС,"
        if "ХимияГодовойкурс" in i: cell += "ХИМ ГК,"
        if "ХимияПолугодовойкурс" in i: cell += "ХИМ ПГК,"
        if "ХимияЭкспресскурс" in i: cell += "ХИМ ЭКС,"
        if "ПрофильнаяматематикаГодовойкурс" in i: cell += "МАТ ГК,"
    row.append(cell[:-1])
    page.append(row)