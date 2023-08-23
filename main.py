import os
import telebot
from telebot import types
from num2words import num2words
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from msgs import *
from t0ken import *

bot = telebot.TeleBot(TOKEN)

print("run in progress...")
numero = 0
placeholders = {}
subs = ["{{номер}}", "{{дата}}", "{{ок_дата}}", "{{стоимость}}", "{{фио}}", "{{нум_ту_вордс}}", "{{ндфл}}",
        "{{ндфл_ту_вордс}}", "{{серияномер}}", "{{код}}", "{{др}}", "{{дата_выд}}", "{{кем}}", "{{адрес}}",
        "{{р/с}}", "{{банк}}", "{{бик}}", "{{@исполн}}", "{{{№исполн}}", "{{фио_кор}}"]


@bot.message_handler(commands=['start'])
def start(message):
    try:
        markup = types.InlineKeyboardMarkup(row_width=1)
        item1 = types.InlineKeyboardButton('Договор ГПХ', callback_data='gph')
        markup.add(item1)
        bot.send_message(message.chat.id, text="Какой документ заполняем?",
                         parse_mode='Markdown', reply_markup=markup)
    except Exception as e:
        print(e)


@bot.message_handler(content_types=['text'])
def answer(message):
    markup = types.InlineKeyboardMarkup(row_width=1)
    placeholders[subs[numero]] = message.text
    print(placeholders[subs[numero]])
    bot.reply_to(message, 'Записал, нажмите *"Далее"* для продолжения', parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "gph")
def fill(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Приступить к заполнению', callback_data='nomer')
    markup.add(item1)
    bot.send_message(call.message.chat.id, text="Чтобы приступить к заполнению анкеты нажмите на кнопку ниже",
                     reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "nomer")
def nom_dog(call):
    global numero
    numero = 0
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Далее', callback_data='date')
    markup.add(item1)
    bot.send_message(call.message.chat.id, text="Введите номер договора",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "date")
def date(call):
    global placeholders, subs, numero
    numero = 1
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='nomer')
    item2 = types.InlineKeyboardButton('Далее', callback_data='date_ok')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text='Введите дату заключения договора (в формате\n(в формате дд.мм.гггг))',
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "date_ok")
def date_ok(call):
    global placeholders, subs, numero
    numero = 2
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='date')
    item2 = types.InlineKeyboardButton('Далее', callback_data='stoim')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text='Введите дату окончания действия договора\n(в формате дд.мм.гггг)',
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "stoim")
def stoim(call):
    global placeholders, subs, numero
    numero = 3
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='date_ok')
    item2 = types.InlineKeyboardButton('Далее', callback_data='fio')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите стоимость услуг исполнителя\n(в формате 75000.00)",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "fio")
def fio(call):
    global placeholders, subs, numero
    numero = 4
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='stoim')
    item2 = types.InlineKeyboardButton('Далее', callback_data='pasp')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите ФИО исполнителя",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "pasp")
def pasp(call):
    global placeholders, subs, numero
    numero = 5
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='fio')
    item2 = types.InlineKeyboardButton('Далее', callback_data='kod')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите серию, номер паспорта",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "kod")
def kod(call):
    global placeholders, subs, numero
    numero = 6
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='pasp')
    item2 = types.InlineKeyboardButton('Далее', callback_data='dr')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите код подразделения",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "dr")
def dr(call):
    global placeholders, subs, numero
    numero = 7
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='kod')
    item2 = types.InlineKeyboardButton('Далее', callback_data='vyd')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите дату роджения\n(в формате дд.мм.гггг)",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "vyd")
def vyd(call):
    global placeholders, subs, numero
    numero = 8
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='dr')
    item2 = types.InlineKeyboardButton('Далее', callback_data='kem')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите дату выдачи паспорта\n(в формате дд.мм.гггг)",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "kem")
def kem(call):
    global placeholders, subs, numero
    numero = 9
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='vyd')
    item2 = types.InlineKeyboardButton('Далее', callback_data='adr')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите кем выдан паспорт",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "adr")
def adr(call):
    global placeholders, subs, numero
    numero = 10
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='kem')
    item2 = types.InlineKeyboardButton('Далее', callback_data='rs')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите адрес прописки исполнителя",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "rs")
def rs(call):
    global placeholders, subs, numero
    numero = 11
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='adr')
    item2 = types.InlineKeyboardButton('Далее', callback_data='bank')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите рассчетный счет",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "bank")
def bank(call):
    global placeholders, subs, numero
    numero = 12
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='rs')
    item2 = types.InlineKeyboardButton('Далее', callback_data='bik')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите название банка исполнителя",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "bik")
def bik(call):
    global placeholders, subs, numero
    numero = 13
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='bank')
    item2 = types.InlineKeyboardButton('Далее', callback_data='mail')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите БИК",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "mail")
def mail(call):
    global placeholders, subs, numero
    numero = 14
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='bik')
    item2 = types.InlineKeyboardButton('Далее', callback_data='tel')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите эл. почту исполнителя",
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "tel")
def tel(call):
    global placeholders, subs, numero
    numero = 15
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='mail')
    item2 = types.InlineKeyboardButton('Готово!', callback_data='send')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text="Введите номер телефона исполнителя\n(в формате 7 (9xx) xxx xx xx)",
                     parse_mode='Markdown', reply_markup=markup)

bot.edit_message_reply_markup(chat_id=call.message.chat.id, message_id=call.message.message_id)

@bot.callback_query_handler(func=lambda call: call.data == "send")
def docx(callback_query):
    try:
        global placeholders
        placeholders["{{нум_ту_вордс}}"] = num2words(int(placeholders["{{стоимость}}"]), lang='ru')
        placeholders["{{ндфл}}"] = str(0.13 * int(placeholders["{{стоимость}}"]))
        placeholders["{{ндфл_ту_вордс}}"] = num2words(int(placeholders["{{ндфл}}"]), lang='ru')
        secondname, firstname, otch = placeholders["{{фио}}"].split()
        placeholders["{{фио_кор}}"] = firstname[0] + '. ' + otch[0] + '. ' + secondname

        print(placeholders)
        doc = Document('Договор ГПХ.docx')

        for paragraph in doc.paragraphs:
            for placeholder, value in placeholders.items():
                if placeholder in paragraph.text:
                    # Заменяем текст и устанавливаем шрифт
                    if value:
                        paragraph.text = paragraph.text.replace(placeholder, value)
                        run = paragraph.runs[0]
                        font = run.font
                        font.name = "Times New Roman"
                        font.size = Pt(12)

        target_text = "Гражданин(ка) Российской Федерации " + placeholders["{{фио}}"]

        for paragraph in doc.paragraphs:
            if target_text in paragraph.text:
                run = paragraph.runs[0]  # Берем первую текстовую руну параграфа
                print(run.text)
                if target_text in run.text:
                    # Разделяем текстовую руну на две части
                    parts = [target_text, run.text.split(target_text)[1]]
                    run.clear()
                    # Добавляем текст перед выделенной частью без жирного начертания
                    new_run = paragraph.add_run(target_text)
                    new_run.bold = True  # Устанавливаем жирное начертание для новой руны
                    font = new_run.font
                    font.name = "Times New Roman"
                    font.size = Pt(12)
                    # Создаем новую руну для выделенной части
                    new_run = paragraph.add_run(parts[1])
                    new_run.bold = False  # Устанавливаем жирное начертание для новой руны

                    font = new_run.font
                    font.name = "Times New Roman"
                    font.size = Pt(12)
                    # Добавляем текст после выделенной части без жирного начертания
                    # new_run.add_text(parts[1])
                    doc.save('Договор ГПХ.docx')
        document_path = 'Договор ГПХ.docx'
        with open(document_path, 'rb') as docx_file:
            bot.send_document(callback_query.message.chat.id, docx_file)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    bot.polling()
