import pandas as pd
import pywhatkit as pw
import time
import pyautogui as pg
import keyboard as k # закоментируй, если запускаешь на linux
import random


# import ctypes
# import py_win_keyboard_layout


# """ Для корректной работы библиотеки pywhatkit необходима en раскладка клавиатуры до запуска кода!!!"""
# # Функция get_layout проверяет раскладку и возвращает её значение
# def get_layout():
#     u = ctypes.windll.LoadLibrary("user32.dll")
#     pf = getattr(u, "GetKeyboardLayout")
#     if hex(pf(0)) == '0x4190419':
#         return 'ru'
#     if hex(pf(0)) == '0x4090409':
#         return 'en'


# # Проверяем раскладку клавиатуры и меняем её, если нужно
# if get_layout() == 'ru':
#     py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)


def read_txt_file():
    """Читает файл с текстом сообщения"""
    file = '<your path>/whatsend/media/text.txt'
    with open(file, 'r', encoding='utf-8') as f:
        text_msg = ''
        for row in f:
            text_msg += row
    return text_msg


def format_phone(phone):
    """Форматирует номер телефона"""
    if len(phone) == 10 and phone[0] == '9':
        phone = '7' + phone
        return phone
    elif phone[0] == '8':
        phone = phone.replace('8', '7', 1)
        return phone
    for i in phone:
        if i in '-() +;':
            phone = phone.replace(i, '')
            return phone
        else:
            return phone


def format_name_txt(name, msg, index_of_name):
    """Работа с именами"""
    # Проверяем наличие имени в строке
    if isinstance(name, float):
        text_msg = f'{msg}'
        return text_msg
    # Если указано только имя
    elif len(name.split(' ')) == 1:
        text_msg = f'{name.title().strip()}, {msg.replace("Добрый", "добрый")}'
        return text_msg
    else:
        # Получаем имя, если указана ещё и фамилия
        first_name = name.split(' ')
        # Форматируем текст сообщения для имени
        # Если в столбце указана фамилия, затем имя, то нужно поменять индекс на [1]
        text_msg = f'{first_name[index_of_name].title().strip()}, {msg.replace("Добрый", "добрый")}'
        return text_msg


def send_msg_image(name_excel_file, name_img, index):
    """
    Принимает 3 аригумента:
    name_excel_file - Имя файла таблицы
    name_img - Имя файла картинки
    index - индекс имени, зависит от формата указания Имя/Фамилия или Фамилия/Имя
    """
    index_of_name = index
    msg = read_txt_file()
    excel_data = pd.read_excel(f'<your path>/whatsend/media/{name_excel_file}.xlsx')  # Открываем файл

    img = f'<your path>/whatsend/media/{name_img}.jpeg'  # Путь к картинке
    
    for i, row in excel_data.iterrows():
        sec = random.randint(20, 30)

        name = row['Name']
        phone = str(row['Phone'])
        if name == 'stop':  # Флаг для ограничения рассылки
            break

        phone = format_phone(phone)
        phone = '+' + phone.replace('.0', '')
        message = format_name_txt(name, msg, index_of_name)

        # Отправка сообщения с картинкой
        pw.sendwhats_image(phone, img, message, wait_time=sec, tab_close=True, close_time=3)

        # Автоматизация нажатия на кнопку "отправить"
        # pg.click(1050, 950) # Для Full Hd монитора
        # pg.click(1963, 990)
        # time.sleep(4)
        # k.press_and_release('enter')


def send_msg(name_excel_file, index):
    """
    Принимает 2 аригумента:
    name_excel_file - Имя файла таблицы
    index - индекс имени, зависит от формата указания Имя/Фамилия или Фамилия/Имя
    """
    index_of_name = index
    msg = read_txt_file()
    excel_data = pd.read_excel(f'<your path>/whatsend/media/{name_excel_file}.xlsx')  # Открываем файл

    
    for _, row in excel_data.iterrows():
        sec = random.randint(20, 30)

        name = row['Name']
        phone = str(row['Phone'])
        if name == 'stop':  # Флаг для ограничения рассылки
            break

        phone = format_phone(phone)
        phone = '+' + phone.replace('.0', '')
        message = format_name_txt(name, msg, index_of_name)

        # Отправка сообщения без картинки
        pw.sendwhatmsg_instantly(phone, message, wait_time=sec, tab_close=True, close_time=4)

        # Автоматизация нажатия на кнопку "отправить"
        # pg.click(1050, 950) # Для Full Hd монитора
        # pg.click(1963, 990)
        # time.sleep(4)
        # k.press_and_release('enter')
