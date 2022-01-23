import os.path
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
import configparser


def number_to_month(number):
    month = (
        'января',
        'февраля',
        'марта',
        'апреля',
        'мая',
        'июня',
        'июля',
        'августа',
        'сентября',
        'октября',
        'ноября',
        'декабря'
    )
    return month[int(number) - 1]


def order_file_coose():
    """процедура указателя имени файла с номерами приказов"""
    messagebox.showinfo(title="Отсутсвуют №приказов", message="Укажите фаил приказов")
    number_filename = askopenfilename(
        defaultextension=".xls", filetypes=[("номера приказов ", ".xls*")]
    )
    return number_filename


def template_file_chooser():
    """процедура указателя на имя файла шаблона"""
    messagebox.showinfo(title="Отсутсвует шаблон", message="Укажите фаил шаблона")
    doc_template_filename = askopenfilename(
        defaultextension= ".docx", filetypes= [("шаблон талона ", ".docx*")]
    )
    return doc_template_filename


# вызываем тк для отрисовки
Tk().withdraw()

# проверяем наличие файла конфига при наличие его загружаем
config_file_name = "config.ini"
if os.path.exists(config_file_name):
    config = configparser.ConfigParser()
    config.read(config_file_name)
    number_filename = config["File names"]["number_filename"]
    doc_template = config["File names"]["doc_template_filename"]
# при отсутсвии создаем фаил опрашиваем пользователя
else:
    number_filename = order_file_coose()
    doc_template = template_file_chooser()

    config = configparser.ConfigParser()
    config["File names"] = {}
    config["File names"]["number_filename"] = number_filename
    config["File names"]["doc_template_filename"] = doc_template
    with open(config_file_name, "w") as config_file:
        config.write(config_file)

# проверяем наличие файла шаблона
if not os.path.exists(doc_template):
    doc_template = template_file_chooser()

    config["File names"]["doc_template_filename"] = doc_template
    with open(config_file_name, "w") as config_file:
        config.write(config_file)

# проверяем наличие файла приказов
if not os.path.exists(number_filename):
    number_filename = order_file_coose()

    config["File names"]["number_filename"] = number_filename
    with open(config_file_name, "w") as config_file:
        config.write(config_file)

# открыаем фаил сверхурочных
messagebox.showinfo(title="Фаил сверхурочных", message="Укажите фаил сверхурочных")
filename = askopenfilename(
    defaultextension=".xls", filetypes=[("exel files ", ".xls*")]
)
# парсим эксель обрезаем 13 первых строк
data = pd.ExcelFile(filename).parse(sheet_name=0, skiprows=13)
# переименовываем колонки и отбрасываем все кроме имени табельного и времени
data.columns = list("ABCDEFGHIJKLMNOP")
data.drop(
    ["A", "B", "D", "E", "I", "J", "F", "G", "L", "M", "N", "O", "P"],
    axis="columns",
    inplace=True,
)
data.columns = ["ФИО работника", "Табельный №", "Количество часов"]
# отбрасываем пустые строки
data.dropna(inplace=True)
data = data[data["ФИО работника"] != "3"]
# переводим часы и табельный в числа отбрасываем нули
data["Количество часов"] = data["Количество часов"].apply(int)
data["Табельный №"] = data["Табельный №"].apply(int)

# подгружам фаил с номерами приказов отбрасываем все лишнее
now = datetime.datetime.now()
month = now.month
# отсееваем пустой лист (да блин там теперь есть такие)
while month > 0:
    month = month - 1
    datan = pd.ExcelFile(number_filename).parse(sheet_name=month, header=None, index=False)
    datan.dropna(inplace=True)
    if len(datan.values) > 0:
        break

# номер приказа в самом низу
date_and_number = datan.iloc[-1:]
pnumber = date_and_number[1].values[0] if '/' in str(date_and_number[1].values[0]) else date_and_number[2].values[0]

# определяем текушую дату
full_date_data = datetime.datetime.today().isoformat()
full_date_data = full_date_data.split("-")
full_date_data[2] = full_date_data[2][:2]
year = int(full_date_data[0]) - 2000
date = "{}. {} ".format(full_date_data[2], number_to_month(full_date_data[1]))
weekday = datetime.datetime.today().weekday()
full_date = "{}. {}. {}".format(full_date_data[2], full_date_data[1], full_date_data[0])

# перебераем все строки выдаем по две на шаблон
for number in range(0, len(data.values), 2):
    doc = DocxTemplate(doc_template)
    FIO, tabn, term = data.values[number]
    # на случай если трок нечетное количество
    if number + 1 < len(data.values):
        FIO2, tabn2, term2 = data.values[number + 1]
    else:
        FIO2, tabn2, term2 = " " * 25, " " * 10, 2
    if weekday == 4:
        minutes = 45
        start = 14
    else:
        minutes = 0
        start = 16
    finish = start + term
    finish2 = start + term2
    context = {
        "tabn": tabn,
        "FIO": FIO,
        "date": date,
        "year": year,
        "term": term,
        "full_date": full_date,
        "prnumber": pnumber,
        "start": start,
        "finish": finish,
        "minut": minutes,
        "tabn2": tabn2,
        "FIO2": FIO2,
        "date": date,
        "year": year,
        "term2": term2,
        "full_date": full_date,
        "prnumber": pnumber,
        "start": start,
        "finish2": finish2,
        "minut": minutes,
    }
    doc.render(context)
    doc.save("талон на печать {}.docx".format(number))
