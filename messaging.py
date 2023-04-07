import telebot
import logging
import xlwt
import xlrd
from time import sleep

ADMIN_ID = 
bot = telebot.TeleBot('', skip_pending=True)
logging.basicConfig(filename="logs.log", level=logging.INFO, format=' %(asctime)s - %(levelname)s - %(message)s')
allowedusers = [[0, 0]]
ur = xlrd.open_workbook("users.xls")
sheetr = ur.sheet_by_index(0)
userslist = sheetr.col_values(0)
grouplist = sheetr.col_values(1)
for i in range(len(userslist)):
    a = int(userslist[i])
    b = int(grouplist[i])
    allowedusers.append([a, b])
logging.info("start bot")


@bot.message_handler(commands=['start'])
def start(message):
    if not (([message.from_user.id, 1] in allowedusers) or ([message.from_user.id, 2] in allowedusers)):
        keyboard = telebot.types.InlineKeyboardMarkup()
        key_g1 = telebot.types.InlineKeyboardButton(text='1', callback_data='change group 1')
        key_g2 = telebot.types.InlineKeyboardButton(text='2', callback_data='change group 2')
        key_cancel = telebot.types.InlineKeyboardButton(text='отмена', callback_data='cancel')
        keyboard.add(key_g1, key_g2)
        keyboard.add(key_cancel)
        bot.send_message(message.from_user.id,
                         text="Привет, когда закончится пара этот бот напомнит, какая следующая, чтобы тебе не "
                              "пришлось искать расписание. Кроме того, бот постоянно обновляет расписание, "
                              "поэтому ты будешь в курсе изменений.\n" + "Если хочешь зарегестрироваться, укажи группу",
                         reply_markup=keyboard)
    else:
        bot.send_message(message.from_user.id, "Вы уже зарегестрированы")


@bot.message_handler(commands=['users'])
def users(message):
    if message.from_user.id == ADMIN_ID:
        bot.send_document(ADMIN_ID, open("users.xls", 'rb'))


@bot.message_handler(commands=['logs'])
def logs(message):
    if message.from_user.id == ADMIN_ID:
        bot.send_document(ADMIN_ID, open("logs.log", 'rb'))


@bot.message_handler(commands=['settings'])
def settings(message):
    if not (([message.from_user.id, 1] in allowedusers) or ([message.from_user.id, 2] in allowedusers)):
        bot.send_message(message.from_user.id, "Вы еще не зарегестрированы")
    else:
        u = "0"
        for i in allowedusers:
            if i[0] == message.from_user.id:
                u = str(i[1])
                break
        keyboard = telebot.types.InlineKeyboardMarkup()
        key_change = telebot.types.InlineKeyboardButton(text='изменить группу', callback_data='change')
        key_delete = telebot.types.InlineKeyboardButton(text='удалить', callback_data='delete')
        key_cancel = telebot.types.InlineKeyboardButton(text='отмена', callback_data='cancel')
        keyboard.add(key_change)
        keyboard.add(key_delete)
        keyboard.add(key_cancel)
        bot.send_message(message.from_user.id, reply_markup=keyboard,
                         text="Ваша группа - " + u + "\n" + "Хотите изменить группу или удалить себя из списка?")


@bot.callback_query_handler(func=lambda call: True)
def callback_worker(call):
    if call.data == "change":
        bot.delete_message(call.message.chat.id, call.message.id)
        keyboard = telebot.types.InlineKeyboardMarkup()
        key_g1 = telebot.types.InlineKeyboardButton(text='1', callback_data='change group 1')
        key_g2 = telebot.types.InlineKeyboardButton(text='2', callback_data='change group 2')
        key_cancel = telebot.types.InlineKeyboardButton(text='отмена', callback_data='cancel')
        keyboard.add(key_g1, key_g2)
        keyboard.add(key_cancel)
        bot.send_message(call.message.chat.id, text="Выберите группу", reply_markup=keyboard)
        logging.info(
            "user changing:" + str(call.message.chat.first_name) + " " + str(call.message.chat.last_name) + ":" + str(
                call.message.chat.username) + ":" + str(call.message.chat.id))
    elif call.data == "delete":
        bot.delete_message(call.message.chat.id, call.message.id)
        r = 0
        for i in range(len(allowedusers)):
            if allowedusers[i][0] == call.message.chat.id:
                r = i
                break
        allowedusers[r][1] = -2
        rewrite_users()
        bot.send_message(call.message.chat.id, "Готово")
        logging.info(
            "user deleted:" + str(call.message.chat.first_name) + " " + str(call.message.chat.last_name) + ":" + str(
                call.message.chat.username) + ":" + str(call.message.chat.id))
    elif call.data == "cancel":
        bot.delete_message(call.message.chat.id, call.message.id)
        bot.send_photo(call.message.chat.id, open("ok.png", 'rb'))
        logging.info(
            "user canceled:" + str(call.message.chat.first_name) + " " + str(call.message.chat.last_name) + ":" + str(
                call.message.chat.username) + ":" + str(call.message.chat.id))
    elif call.data == "change group 1":
        bot.delete_message(call.message.chat.id, call.message.id)
        r = -1
        for i in range(len(allowedusers)):
            if allowedusers[i][0] == call.message.chat.id:
                r = i
                break
        if r == -1:
            allowedusers.append([call.message.chat.id, 1])
        else:
            allowedusers[r][1] = 1
        rewrite_users()
        bot.send_message(call.message.chat.id, "Готово")
        logging.info("user chose group 1:" + str(call.message.chat.first_name) + " " + str(
            call.message.chat.last_name) + ":" + str(call.message.chat.username) + ":" + str(call.message.chat.id))
    elif call.data == "change group 2":
        bot.delete_message(call.message.chat.id, call.message.id)
        r = -1
        for i in range(len(allowedusers)):
            if allowedusers[i][0] == call.message.chat.id:
                r = i
                break
        if r == -1:
            allowedusers.append([call.message.chat.id, 2])
        else:
            allowedusers[r][1] = 2
        rewrite_users()
        bot.send_message(call.message.chat.id, "Готово")
        logging.info(
            "user chose group 2:" + str(call.message.chat.first_name) + " " + str(
                call.message.chat.last_name) + ":" + str(
                call.message.chat.username) + ":" + str(call.message.chat.id))


def rewrite_users():
    l = len(allowedusers)
    i = 1
    uw = xlwt.Workbook(encoding="utf8")
    sheetw = uw.add_sheet("users")
    while i < l:
        if allowedusers[i][1] != -2:
            sheetw.write(i - 1, 0, allowedusers[i][0])
            sheetw.write(i - 1, 1, allowedusers[i][1])
            i += 1
        else:
            del allowedusers[i]
            l -= 1
    uw.save("users.xls")


while True:
    try:
        bot.polling(none_stop=True, interval=1)
    except Exception:
        logging.error("network trouble")
        sleep(15)
