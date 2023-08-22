import os
import telebot
from telebot import types
from t0ken import *

bot = telebot.TeleBot("6313360423:AAFnAi6FglZYbnU6hglG04PKLy7Ee9d-Lfw")
zap = False


@bot.message_handler(commands=['start'])
def start(message):
    try:
        # Создаем объект клавиатуры для inline кнопки
        keyboard = types.InlineKeyboardMarkup()
        button = types.InlineKeyboardButton(text="Договор ГПХ", callback_data="gph")
        keyboard.add(button)

        # Отправляем пользователю сообщение с inline кнопкой
        bot.send_message(message.chat.id, "Какой договор заполняем?", reply_markup=keyboard)
    except Exception as e:
        print(e)


@bot.message_handler(content_types=['text'])
def answer(message):
    global zap
    if zap:
        markup = types.InlineKeyboardMarkup(row_width=1)
        print(callbkz[s])
        if 0 <= s <= len(callbkz) - 1:
            bot.send_message(message.chat.id, text='Введите ' + msgs[callbkz[s]])
            callbkz[s] =
        else:
            s = 0
    else:
        msg1 = bot.send_message(message.chat.id, 'Пока что бот не умеет отвечать на сообщения.')
        time.sleep(10)
        bot.delete_message(chat_id=message.chat.id, message_id=msg1.message_id)


@bot.callback_query_handler(func=lambda call: True)
def download_docx(callback_query):
    try:
        callbkz = [nomer, date, fio, endate, num, num2, ndfl, ndfl2, pas, kod, dr, vyd, kem, adr, rs, bank, bik, mail,
                   tel]
        for i in spis:
            i = ''
        callbkz.pop(num2, ndfl, ndfl2)

        if call.data == "gph":
            doc = Document("Договор ГПХ.docx")
            markup = types.InlineKeyboardMarkup(row_width=1)
            item1 = types.InlineKeyboardButton("Начать заполнение шаблона", callback_data='nach')
            msgs = {nomer: "номер договора", date: 'дату заключения договора (в формате "14" марта 2023)',
                    fio: "ФИО исполнителя", endate: "дату окончания действия договора",
                    num: "цену услуг исполнителя", pas: "серию, номер паспорта",
                    kod: "код подразделения", dr: "дату роджения",
                    vyd: "дату выдачи паспорта", kem: "кем выдан паспорт",
                    adr: "адрес прописки исполнителя",
                    rs: "рассчетный счет", bank: "банк исполнителя",
                    bik: "БИК", mail: "эл. почту исполнителя", tel: "номер телефона исполнителя"
                    }
            bot.send_message(call.message.chat.id, text="Чтобы начать заполнять нажмите кнопку ниже",
                             reply_markup=markup)

            # заполнение
        elif call.data == "nach":
            zap = True

        placeholders = {"{{номер}}": nomer, "{{дата}}": date,
                        "{{фио}}": fio, "{{ок_дата}}": endate,
                        "{{стоимость}}": num, "{{нум_ту_вордс}}": num2,
                        "{{ндфл}}": ndfl, "{{ндфл_ту_вордс}}": ndfl2,
                        "{{серияномер}}": pas,
                        "{{код}}": kod, "{{др}}": dr,
                        "{{дата_выд}}": vyd, "{{кем}}": kem,
                        "{{адрес}}": adr, "{{р/с}}": rs,
                        "{{банк}}": bank, "{{бик}}": bik,
                        "{{@исполн}}": mail, "{{{№исполн}}": tel}

        placeholders["{{нум_ту_вордс}}"] = num2words(int(placeholders["{{стоимость}}"]), lang='ru')
        placeholders["{{ндфл}}"] = 0.13 * int(placeholders["{{стоимость}}"])
        placeholders["{{ндфл_ту_вордс}}"] = num2words(int(placeholders["{{ндфл}}"]), lang='ru')
        secondname, firstname, otch = placeholders["{{фио}}"].split()
        res = firstname[0] + '. ' + otch[0] + '. ' + secondname

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
