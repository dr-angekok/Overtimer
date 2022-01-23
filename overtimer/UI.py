from overtimer import config_crud, xl_parsers, formaters, core
from os.path import isfile
from sys import exit
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox


def main():
    Tk().withdraw()
    config = config_crud.CrudConfig()

    if config.template_path == '/' or not isfile(config.template_path):
        messagebox.showerror(
            "Отсутсвует шаблон",
            "Нужно выбрать фаил шаблона")
        doc_template_filename = askopenfilename(
            defaultextension=".docx", filetypes=[("шаблон талона ", ".docx*")]
        )
        if not doc_template_filename:
            config.template_path_set('/')
            messagebox.showerror(
                "Ошибка",
                "Не возможно завершить без шаблона")
            exit()
        config.template_path_set(doc_template_filename)

    if config.person_list_path == '/' or not isfile(config.person_list_path):
        messagebox.showerror(
            'Отсутсвует список контингента',
            'Нужно выбрать фаил списка сверхурочных'
        )
        person_filename = askopenfilename(
            defaultextension=".xls", filetypes=[("exel files ", ".xls*")]
        )
        if not person_filename:
            config.template_path_set('/')
            messagebox.showerror(
                "Ошибка",
                "Не возможно завершить без списка сверхурочных")
            exit()
        config.person_list_path_set(person_filename)

    if config.orders_path == '/' or not isfile(config.orders_path):
        messagebox.showerror(
            'Отсутсвует фаил приказов',
            'Нужно выбрать фаил списка приказов'
        )
        order_filename = askopenfilename(
            defaultextension=".xls", filetypes=[("exel files ", ".xls*")]
        )
        if not order_filename:
            config.orders_path_set('/')
            messagebox.showerror(
                "Ошибка",
                "Не возможно завершить без списка номеров приказов")
            exit()
        config.orders_path_set(order_filename)

    order_number = xl_parsers.get_order_number(config.orders_path)
    persons_data = xl_parsers.get_personal_df(config.person_list_path)

    formaters.make_docx(core.make_contexts(order_number, persons_data, config.chief), config.template_path)


if __name__ == '__main__':
    main()
