import fill_docx
import tkinter as tk


def start():

    # При нажатии на кнопку "Отправить" срабатывает скрипт, который создает словарь с данными из полей, и передает в
    # модуль, который редактирует файл docx
    def submit():

        # словарь с данными
        data = {
            # СВЯЗЬ
            'phone': phone.get(),
            'email': email.get(),
            # информация о контакте
            'surname': surname.get(),
            'name': name.get(),
            'middlename': middlename.get(),
            'sex': sex.get(),
            'date_of_birth': date_of_birth.get(),
            'town_of_birth': town_of_birth.get(),
            'citizenship': citizenship.get(),
            # данные паспорта
            'passport_series': passport_series.get(),
            'passport_number': passport_number.get(),
            'passport_office': passport_office.get(),
            'passport_date': passport_date.get(),
            'division_code': division_code.get(),
            # адрес регистрации
            'region': region.get(),
            'district': district.get(),
            'town': town.get(),
            'locality': locality.get(),
            'street': street.get(),
            'house': house.get(),
            'corpus': corpus.get(),
            'flat': flat.get(),
            'address_index': address_index.get(),
        }

        # Очистка полей после нажатия на кнопку "Отправить"
        phone.delete(0, 'end')
        email.delete(0, 'end')
        surname.delete(0, 'end')
        name.delete(0, 'end')
        middlename.delete(0, 'end')
        sex.delete(0, 'end')
        date_of_birth.delete(0, 'end')
        town_of_birth.delete(0, 'end')
        citizenship.delete(0, 'end')
        passport_series.delete(0, 'end')
        passport_number.delete(0, 'end')
        passport_office.delete(0, 'end')
        passport_date.delete(0, 'end')
        division_code.delete(0, 'end')
        region.delete(0, 'end')
        district.delete(0, 'end')
        town.delete(0, 'end')
        locality.delete(0, 'end')
        street.delete(0, 'end')
        house.delete(0, 'end')
        corpus.delete(0, 'end')
        flat.delete(0, 'end')
        address_index.delete(0, 'end')

        # отправляем на редактирование/создание данные из полей и очищение их
        if create_docx_path['state'] == 'normal':
            filename = create_docx_path.get()
            fill_docx.write_docx(data, filename)
            create_docx_path.delete(0, 'end')
        else:
            path = edit_docx_path.get()
            fill_docx.refactor_docx(path, data)
            edit_docx_path.delete(0, 'end')

    # срабатывает при выборе radio кнопки "Создаём"
    # делаем неактивным поле ввода "Путь к файлу"
    def create_docx_path():
        create_docx_path.config(state='normal')
        edit_docx_path.config(state='disabled')

    # срабатывает при выборе radio кнопки "Редактируем"
    # делаем неактивным поле ввода "Название файла"
    def edit_docx_path():
        create_docx_path.config(state='disabled')
        edit_docx_path.config(state='normal')

    # Создание и отображение окна приложения
    win = tk.Tk()
    win.title('Данные паспорта в docx')
    win.geometry('500x650')
    win.resizable(False, False)

    # Разметка лейблов по окну
    # лицевая страница паспорта
    tk.Label(win, text='ВВОД ДАННЫХ', bg='green').grid(row=0, column=0, sticky='w')
    tk.Label(win, text='ЛИЦЕВАЯ СТОРОНА ПАСПОРТА', bg='yellow').grid(row=1, column=0, sticky='w')
    tk.Label(win, text='Телефон').grid(row=2, column=0, sticky='w')
    tk.Label(win, text='E-mail').grid(row=3, column=0, sticky='w')
    tk.Label(win, text='Фамилия').grid(row=4, column=0, sticky='w')
    tk.Label(win, text='Имя').grid(row=5, column=0, sticky='w')
    tk.Label(win, text='Отчество').grid(row=6, column=0, sticky='w')
    tk.Label(win, text='Пол').grid(row=7, column=0, sticky='w')
    tk.Label(win, text='Дата рождения').grid(row=8, column=0, sticky='w')
    tk.Label(win, text='Место рождения').grid(row=9, column=0, sticky='w')
    tk.Label(win, text='Гражданство').grid(row=10, column=0, sticky='w')
    tk.Label(win, text='Серия паспорта').grid(row=11, column=0, sticky='w')
    tk.Label(win, text='Номер паспорта').grid(row=12, column=0, sticky='w')
    tk.Label(win, text='Выдан').grid(row=13, column=0, sticky='w')
    tk.Label(win, text='Дата выдачи').grid(row=14, column=0, sticky='w')
    tk.Label(win, text='Код подразделения').grid(row=15, column=0, sticky='w')

    # страница с пропиской
    tk.Label(win, text='СТРАНИЦА С РЕГИСТРАЦИЕЙ', bg='yellow').grid(row=16, column=0, sticky='w')
    tk.Label(win, text='Регион').grid(row=17, column=0, sticky='w')
    tk.Label(win, text='Район').grid(row=18, column=0, sticky='w')
    tk.Label(win, text='Город').grid(row=19, column=0, sticky='w')
    tk.Label(win, text='Населенный пункт').grid(row=20, column=0, sticky='w')
    tk.Label(win, text='Улица').grid(row=21, column=0, sticky='w')
    tk.Label(win, text='Дом').grid(row=22, column=0, sticky='w')
    tk.Label(win, text='Корпус/Строение').grid(row=23, column=0, sticky='w')
    tk.Label(win, text='Квартира').grid(row=24, column=0, sticky='w')
    tk.Label(win, text='Почтовый индекс').grid(row=25, column=0, sticky='w')

    # создание полей для ввода данных ЛИЦЕВОЙ СТОРОНЫ и СВЯЗИ
    phone = tk.Entry(win)
    email = tk.Entry(win)
    surname = tk.Entry(win)
    name = tk.Entry(win)
    middlename = tk.Entry(win)
    sex = tk.Entry(win)
    date_of_birth = tk.Entry(win)
    town_of_birth = tk.Entry(win)
    citizenship = tk.Entry(win)
    passport_series = tk.Entry(win)
    passport_number = tk.Entry(win)
    passport_office = tk.Entry(win)
    passport_date = tk.Entry(win)
    division_code = tk.Entry(win)

    # создание полей для ввода данных СТРАНИЦЫ С РЕГИСТРАЦИЕЙ
    region = tk.Entry(win)
    district = tk.Entry(win)
    town = tk.Entry(win)
    locality = tk.Entry(win)
    street = tk.Entry(win)
    house = tk.Entry(win)
    corpus = tk.Entry(win)
    flat = tk.Entry(win)
    address_index = tk.Entry(win)

    # разметка полей для ввода данных ЛИЦЕВОЙ СТОРОНЫ И СВЯЗИ
    phone.grid(row=2, column=1)
    email.grid(row=3, column=1)
    surname.grid(row=4, column=1)
    name.grid(row=5, column=1)
    middlename.grid(row=6, column=1)
    sex.grid(row=7, column=1)
    date_of_birth.grid(row=8, column=1)
    town_of_birth.grid(row=9, column=1)
    citizenship.grid(row=10, column=1)
    passport_series.grid(row=11, column=1)
    passport_number.grid(row=12, column=1)
    passport_office.grid(row=13, column=1)
    passport_date.grid(row=14, column=1)
    division_code.grid(row=15, column=1)

    # разметка полей для ввода СТРАНИЦЫ С РЕГИСТРАЦИЕЙ
    region.grid(row=17, column=1)
    district.grid(row=18, column=1)
    town.grid(row=19, column=1)
    locality.grid(row=20, column=1)
    street.grid(row=21, column=1)
    house.grid(row=22, column=1)
    corpus.grid(row=23, column=1)
    flat.grid(row=24, column=1)
    address_index.grid(row=25, column=1)

    # выпадающий список для выбора редактируем или создаём файл
    choice_var = tk.StringVar(value="new")

    # создание radio сущности для СОЗДАНИЯ файла docx
    radio_create = tk.Radiobutton(
        win,
        text="Создаём",
        variable=choice_var,
        value="new",
        command=create_docx_path
    )
    radio_create.grid(row=27, column=0)

    # создание radio сущности для РЕДАКТИРОВАНИЯ файла docx
    radio_edit = tk.Radiobutton(
        win,
        text="Редактируем",
        variable=choice_var,
        value="edit",
        command=edit_docx_path,
    )
    radio_edit.grid(row=27, column=1)

    # разметка лейбла "название файла" для создания
    tk.Label(win, text='Название файла, чтобы создать').grid(row=28, column=0, sticky='w')

    # разметка лейбла "путь к файлу для редактирования" для редактирования
    tk.Label(win, text='Путь к файлу на редактирование').grid(row=28, column=1, sticky='w')

    # разметка формы для заполнения filename создания docx
    create_docx_path = tk.Entry(win)
    create_docx_path.grid(row=29, column=0, sticky='we')

    # разметка формы для заполнения path для редактирования docx
    edit_docx_path = tk.Entry(win, state='disabled')
    edit_docx_path.grid(row=29, column=1, sticky='we')

    # ЗАВЕРШАЮЩАЯ КНОПКА "ОТПРАВИТЬ"
    submit_btn = tk.Button(win, text='Отправить', command=submit)
    submit_btn.grid(row=30, column=1)

    win.mainloop()


if __name__ == '__main__':
    start()
