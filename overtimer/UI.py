from overtimer import config_crud, xl_parsers, formaters, core
from os.path import isfile
from sys import exit
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox


def do_error_exit(command, message):
    command('/')
    messagebox.showerror(
        'Ошибка',
        message,
    )
    exit()


def main():
    Tk().withdraw()
    config = config_crud.CrudConfig()

    if config.template_path == '/' or not isfile(config.template_path):
        messagebox.showerror(
            'Отсутсвует шаблон',
            'Нужно выбрать фаил шаблона')
        doc_template_filename = askopenfilename(
            defaultextension=".docx", filetypes=[('шаблон талона ', '.docx*')]
        )
        if not doc_template_filename:
            config.template_path_set('/')
            messagebox.showerror(
                'Ошибка',
                'Не возможно завершить без шаблона')
            exit()

        config.template_path_set(doc_template_filename)

    if config.person_list_path == '/' or not isfile(config.person_list_path):
        messagebox.showerror(
            'Отсутсвует список контингента',
            'Нужно выбрать фаил списка сверхурочных'
        )
        person_filename = askopenfilename(
            defaultextension=".xls", filetypes=[('exel files ', '.xls*')]
        )
        if not person_filename:
            do_error_exit(config.template_path_set, 'Не возможно завершить без списка сверхурочных')
        config.person_list_path_set(person_filename)

    try:
        persons_data = xl_parsers.get_personal_df(config.person_list_path)
    except IsADirectoryError:
        do_error_exit(config.person_list_path_set, 'Список недоступен, продолжение не возможно')
    except FileNotFoundError:
        do_error_exit(config.person_list_path_set, 'Список недоступен, продолжение не возможно')
    except IndexError:
        do_error_exit(config.person_list_path_set, 'Повреждена структура файла, останов')
    except KeyError:
        do_error_exit(config.person_list_path_set, 'Нет нужных колонок, останов')
    except ValueError:
        do_error_exit(config.person_list_path_set, 'Нет нужных колонок, останов')
    except Exception as e:
        do_error_exit(config.person_list_path_set, e.message)

    if config.orders_path == '/' or not isfile(config.orders_path):
        messagebox.showerror(
            'Отсутсвует фаил приказов',
            'Нужно выбрать фаил списка приказов'
        )
        order_filename = askopenfilename(
            defaultextension=".xls", filetypes=[('exel files ', '.xls*')]
        )
        if not order_filename:
            do_error_exit(config.orders_path_set, 'Не возможно завершить без списка номеров приказов')
        config.orders_path_set(order_filename)

    try:
        order_number = xl_parsers.get_order_number(config.orders_path)
    except IsADirectoryError:
        do_error_exit(config.orders_path_set, 'Список недоступен, продолжение не возможно')
    except FileNotFoundError:
        do_error_exit(config.orders_path_set, 'Список недоступен, продолжение не возможно')
    except IndexError:
        do_error_exit(config.orders_path_set, 'Повреждена структура файла, останов')
    except KeyError:
        do_error_exit(config.orders_path_set, 'Нет нужных колонок, останов')
    except ValueError:
        do_error_exit(config.orders_path_set, 'Нет нужных колонок, останов')
    except Exception as e:
        do_error_exit(config.orders_path_set, e.message)

    formaters.make_docx(core.make_contexts(order_number, persons_data, config.chief), config.template_path)


if __name__ == '__main__':
    main()
