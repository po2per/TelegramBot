import os
import telebot
from telebot import types
import openpyxl

# Работа с Ботом

bot = telebot.TeleBot('5456238940:AAEcPWG4cG-Oom_FyGD2aJMZE7mkeOJ4CGY')


@bot.message_handler(commands=['start'])
def start(message):
    mess = f'Здравствуйте! Спасибо, что выбрали Skink!' \
           f'\nМеня зовут <b>SkinkBot</b>, я создан, для того, чтобы помочь вам!' \
           f'\n\n<b><u>"Привет"</u></b>, чтобы поздороваться со мной' \
           f'\n<b><u>"Номер заказа"</u></b>, для того чтобы узнать номер вашего заказа' \
           f'\n<b><u>"Статус"</u></b>, для того, чтобы узнать статус вашего заказа' \
           f'\n<b><u>"Макет"</u></b>, чтобы увидеть макет своего чехла, если он готов' \
           f'\n<b><u>"Сайт"</u></b>, перейти на наш сайт и оформить заказ!' \
           f'\nЕсли вам необходимо загрузить свою картинку, просто перенесите ее ко мне' \
           f'\nЛюбую другую фразу, для связи с горячей линией поддержки'
    bot.send_message(message.chat.id, mess, parse_mode='html')


@bot.message_handler(content_types=['text'])
def get_user_text(message):
    try:
        number = int(message.text)
        bot.send_message(message.chat.id, "Чтобы узнать статус заказа, сначала отправьте слово 'Статус' в сообщении")

    except:
        if message.text == "Привет":
            bot.send_message(message.chat.id,
                             f'Здравствуйте, <b>{message.from_user.first_name} {message.from_user.last_name}</b>! '
                             f'Чем могу быть полезен?',
                             parse_mode='html')
        elif message.text == "Статус":
            msg = bot.send_message(message.chat.id, f'Введите номер заказа', parse_mode='html')
            bot.register_next_step_handler(msg, user_number)
        elif message.text == "Номер заказа":
            bot.send_message(message.chat.id, f'К сожалению я пока не умею искать заказы, но очень скоро научусь!',
                             parse_mode='html')
        elif message.text == "Сайт":
            markup = types.InlineKeyboardMarkup()
            markup.add(types.InlineKeyboardButton("Skink Shop", url="https://skink.shop/"))
            bot.send_message(message.chat.id,
                             f'Посетите наш сайт, для оформления заказа!', reply_markup=markup)
        elif message.text == "Макет":
            msg = bot.send_message(message.chat.id, f'Введите номер заказа', parse_mode='html')
            bot.register_next_step_handler(msg, user_footage)
        else:
            bot.send_message(message.chat.id, "К сожалению я вас не понимаю. Если у вас остались вопросы, "
                                              "напишите в поддержку @skink_shop, вам обязательно ответят!")


def user_number(message):
    print(type(message.text))
    number = int(message.text)
    searcher = False
    wb = openpyxl.reader.excel.load_workbook(filename="1.xlsx")
    wb.active = 0
    sheet = wb.active
    for i in range(1, 1000):
        if int(number) == sheet['C' + str(i)].value and sheet['B' + str(i)].value is not None and \
                sheet['H' + str(i)].value is not None:
            searcher = True
            bot.send_message(message.chat.id, f"Заказ № {sheet['C' + str(i)].value} {sheet['B' + str(i)].value} "
                                              f"трек номер: {sheet['H' + str(i)].value}")

        elif int(number) == sheet['C' + str(i)].value and sheet['B' + str(i)].value is not None and \
                sheet['H' + str(i)].value is None:
            searcher = True
            bot.send_message(message.chat.id, f"Заказ № {sheet['C' + str(i)].value} {sheet['B' + str(i)].value}"
                                              f" трек номер скоро будет доступен")

        elif int(number) == sheet['C' + str(i)].value and sheet['B' + str(i)].value is None:
            searcher = True
            bot.send_message(message.chat.id, f"Ваш заказ находится в очереди для производства")

    if not searcher:
        bot.send_message(message.chat.id, f"Заказ не обработан или не существует")


def user_footage(message):
    try:
        src = f'pic/{message.text}/footage3.png'
        photo = open(src, 'rb')
        bot.send_photo(message.chat.id, photo)
        src_info = f'pic/{message.text}/1.txt'
        file = open(src_info, encoding='utf-8')
        phone_model = file.readline()
        phone_model = file.readline()
        phone_model = file.readline()
        bot.send_message(message.chat.id, f'Модель телефона: {phone_model}')
        src = f'pic/{message.text}/footage2.png'
        photo = open(src, 'rb')
        bot.send_photo(message.chat.id, photo)
        file.seek(0)
        phone_model = file.readline()
        phone_model = file.readline()
        bot.send_message(message.chat.id, f'Модель телефона: {phone_model}')
        src = f'pic/{message.text}/footage.png'
        photo = open(src, 'rb')
        bot.send_photo(message.chat.id, photo)
        file.seek(0)
        phone_model = file.readline()
        bot.send_message(message.chat.id, f'Модель телефона: {phone_model}')
        file.close()
    except:
        try:
            src = f'pic/{message.text}/footage2.png'
            photo = open(src, 'rb')
            bot.send_photo(message.chat.id, photo)
            src_info = f'pic/{message.text}/1.txt'
            file = open(src_info, encoding='utf-8')
            phone_model = file.readline()
            phone_model = file.readline()
            bot.send_message(message.chat.id, f'Модель телефона: {phone_model}')
            src = f'pic/{message.text}/footage.png'
            photo = open(src, 'rb')
            bot.send_photo(message.chat.id, photo)
            file.seek(0)
            phone_model = file.readline()
            bot.send_message(message.chat.id, f'Модель телефона: {phone_model}')
            file.close()
        except:
            try:
                src = f'pic/{message.text}/footage.png'
                photo = open(src, 'rb')
                bot.send_photo(message.chat.id, photo)
                src_info = f'pic/{message.text}/1.txt'
                file = open(src_info, encoding='utf-8')
                phone_model = file.readline()
                bot.send_message(message.chat.id, f'Модель телефона: {phone_model}')
                file.close()
            except:
                bot.send_message(message.chat.id, f'Макет не готов')


@bot.message_handler(content_types=['photo'])
def get_user_photo(message):
    file_photo = bot.get_file(message.photo[-1].file_id)
    filename, file_expansion = os.path.splitext(file_photo.file_path)

    downloaded_file_photo = bot.download_file(file_photo.file_path)

    src = 'photos/' + message.from_user.first_name + file_expansion
    with open(src, 'wb') as new_file:
        new_file.write(downloaded_file_photo)
    bot.send_message(message.chat.id, 'Пожалуйста отправьте картинку файлом без сжатия, как показано в инструкции:')


bot.polling(none_stop=True)
