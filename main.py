import os
import telebot
from telebot import types
from msgs import *
from t0ken import *

bot = telebot.TeleBot(TOKEN)

print("run in progress...")
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


@bot.callback_query_handler(func=lambda call: call.data == "gph")
def fill(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Приступить к заполнению', callback_data='nomer')
    markup.add(item1)
    bot.send_message(call.message.chat.id, text="Чтобы приступить к заполнению анкеты нажмите на кнопку ниже",
                     reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "nomer")
def nom_dog(call):
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Далее', callback_data='date')
    markup.add(item1)
    bot.send_message(call.message.chat.id, text=nomer,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "date")
def date(call):
    global placeholders, subs
    placeholders[subs[0]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[0]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='nomer')
    item2 = types.InlineKeyboardButton('Далее', callback_data='date_ok')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=date,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "date_ok")
def date_ok(call):
    global placeholders, subs
    placeholders[subs[1]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[1]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='date')
    item2 = types.InlineKeyboardButton('Далее', callback_data='stoim')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=date_ok,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "stoim")
def stoim(call):
    global placeholders, subs
    placeholders[subs[2]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[2]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='date_ok')
    item2 = types.InlineKeyboardButton('Далее', callback_data='fio')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=stoim,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "fio")
def fio(call):
    global placeholders, subs
    placeholders[subs[3]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[3]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='stoim')
    item2 = types.InlineKeyboardButton('Далее', callback_data='pasp')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=fio,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "pasp")
def pasp(call):
    global placeholders, subs
    placeholders[subs[4]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[4]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='fio')
    item2 = types.InlineKeyboardButton('Далее', callback_data='kod')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=pasp,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "kod")
def kod(call):
    global placeholders, subs
    placeholders[subs[5]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[5]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='pasp')
    item2 = types.InlineKeyboardButton('Далее', callback_data='dr')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=kod,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "dr")
def dr(call):
    global placeholders, subs
    placeholders[subs[6]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[6]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='kod')
    item2 = types.InlineKeyboardButton('Далее', callback_data='vyd')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=dr,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "vyd")
def vyd(call):
    global placeholders, subs
    placeholders[subs[7]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[7]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='dr')
    item2 = types.InlineKeyboardButton('Далее', callback_data='kem')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=vyd,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "kem")
def kem(call):
    global placeholders, subs
    placeholders[subs[8]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[8]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='vyd')
    item2 = types.InlineKeyboardButton('Далее', callback_data='adr')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=kem,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "adr")
def adr(call):
    global placeholders, subs
    placeholders[subs[9]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[9]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='kem')
    item2 = types.InlineKeyboardButton('Далее', callback_data='rs')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=adr,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "rs")
def rs(call):
    global placeholders, subs
    placeholders[subs[10]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[10]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='adr')
    item2 = types.InlineKeyboardButton('Далее', callback_data='bank')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=rs,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "bank")
def bank(call):
    global placeholders, subs
    placeholders[subs[11]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[11]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='rs')
    item2 = types.InlineKeyboardButton('Далее', callback_data='bik')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=bank,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "bik")
def bik(call):
    global placeholders, subs
    placeholders[subs[12]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[12]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='bank')
    item2 = types.InlineKeyboardButton('Далее', callback_data='mail')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=bik,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "mail")
def mail(call):
    global placeholders, subs
    placeholders[subs[13]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[13]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='bik')
    item2 = types.InlineKeyboardButton('Далее', callback_data='tel')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=mail,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "tel")
def tel(call):
    global placeholders, subs
    placeholders[subs[14]] = call.message.text
    print(call.message.text, '\tvs\t', placeholders[subs[14]])
    markup = types.InlineKeyboardMarkup(row_width=1)
    item1 = types.InlineKeyboardButton('Назад', callback_data='mail')
    item2 = types.InlineKeyboardButton('Готово!', callback_data='send')
    markup.add(item1, item2)
    bot.send_message(call.message.chat.id, text=tel,
                     parse_mode='Markdown', reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data == "send")
def docx(call):
    try:
        global placeholders, subs
        placeholders[subs[-1]] = call.message.text
        print(call.message.text, '\tvs\t', placeholders[subs[-1]])
        placeholders["{{нум_ту_вордс}}"] = num2words(int(placeholders["{{стоимость}}"]), lang='ru')
        placeholders["{{ндфл}}"] = 0.13 * int(placeholders["{{стоимость}}"])
        placeholders["{{ндфл_ту_вордс}}"] = num2words(int(placeholders["{{ндфл}}"]), lang='ru')
        secondname, firstname, otch = placeholders["{{фио}}"].split()
        placeholders["{{фио_кор}}"] = firstname[0] + '. ' + otch[0] + '. ' + secondname

        print(placeholders)

        target_text = placeholders["{{фио}}"]
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
                    doc.save('ГПХ' + placeholders["{{фио_кор}}"] + '.docx')
        document_path = 'ГПХ' + placeholders["{{фио_кор}}"] + '.docx'
        with open(document_path, 'rb') as docx_file:
            bot.send_document(callback_query.message.chat.id, docx_file)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    bot.polling()
