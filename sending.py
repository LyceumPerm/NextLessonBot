import telebot
import xlrd
import os
from time import sleep
import datetime
import wget
import logging

t = open("TOKEN.txt")
TOKEN = t.readline().strip()
bot = telebot.TeleBot(TOKEN)
t.close()
logging.basicConfig(filename="logs.log", level=logging.INFO, format=' %(asctime)s - %(levelname)s - %(message)s',
                    encoding="utf8")

LINK = "https://docs.google.com/spreadsheets/d/1tGbeevMu_7_n_pKDFjH3cNFNigClVW3v/export?format=xlsx&id=1tGbeevMu_7_n_pKDFjH3cNFNigClVW3v"
MONDAY = "17.04"

A1 = [["" for i in range(2)] for j in range(4)]
A2 = [["" for k in range(2)] for l in range(4)]
allowedusers = [[0, 0]]
del_array = []
del_array2 = []


def find_index(workbook):
    s = workbook.sheet_names()
    index = -1
    for i in range(len(s)):
        if MONDAY in s[i]:
            index = i
            break
    if index != -1:
        return index


def delete_message():
    for i in del_array:
        try:
            bot.delete_message(i[0], i[1])
        except:
            pass
    del del_array[0:]
    logging.info("messages deleted")


def delete_message2():
    for i in del_array2:
        try:
            bot.delete_message(i[0], i[1])
        except:
            pass
    del del_array2[0:]
    logging.info("morning messages deleted")


def update_schedule():
    if os.path.isfile("Schedule.xlsx"):
        os.remove("Schedule.xlsx")
    t = True
    while t:
        try:
            wget.download(LINK, "Schedule.xlsx")
            logging.info("schedule downloaded")
            t = False
        except Exception:
            logging.error("schedule download failed")
            sleep(5)
            t = True
    tmp = os.listdir()
    for item in tmp:
        if item.endswith(".tmp"):
            os.remove(item)


def update_users():
    ur = xlrd.open_workbook("users.xls")
    sheet_r = ur.sheet_by_index(0)
    n = sheet_r.nrows
    del allowedusers[1:]
    for i in range(n):
        a = int(sheet_r.row_values(i)[0])
        b = int(sheet_r.row_values(i)[1])
        allowedusers.append([a, b])
    logging.info("users updated")


def is_merged(r, c):
    wb1 = xlrd.open_workbook("Schedule.xlsx")
    sheet1 = wb1.sheet_by_index(find_index(wb1))
    m = sheet1.merged_cells
    return (r, r + 1, c - 1, c + 1) in m


def get_schedule(weekday):
    wb = xlrd.open_workbook("Schedule.xlsx")
    sheet = wb.sheet_by_index(find_index(wb))
    if weekday > 0:
        x = 1
    else:
        x = 0
    for i in range(4):
        k = sheet.row_values(3 + 5 * weekday + i + x)[60]
        if is_merged(3 + weekday * 5 + i + x, 59):
            l = sheet.row_values(3 + weekday * 5 + i + x)[58]
            if l.find("(") == -1:
                A1[i][0] = l.strip()
                A2[i][0] = l.strip()
            else:
                A1[i][0] = l[:l.find("(")].strip()
                A2[i][0] = l[:l.find("(")].strip()
            if type(k) == float:
                k = str(k)[:-2]
            A1[i][1] = k
            A2[i][1] = k
        else:
            l1 = sheet.row_values(3 + weekday * 5 + i + x)[58]
            l2 = sheet.row_values(3 + weekday * 5 + i + x)[59]
            if l1.find("(") == -1:
                A1[i][0] = l1.strip()
            else:
                A1[i][0] = l1[:l1.find("(")].strip()
            if l2.find("(") == -1:
                A2[i][0] = l2.strip()
            else:
                A2[i][0] = l2[:l2.find("(")].strip()
            if type(k) == float:
                k = str(k)[:-2]
            if k.find("/") != -1:
                A1[i][1] = k[:k.find("/")]
                A2[i][1] = k[k.find("/") + 1:]
            elif A1[i][0] != "" and A2[i][0] == "":
                A1[i][1] = k
                A2[i][1] = ""
            elif A1[i][0] == "" and A2[i][0] != "":
                A1[i][1] = ""
                A2[i][1] = k
            else:
                A1[i][1] = k
                A2[i][1] = k
        if A1[i][0] == "":
            A1[i][1] = "---"
        if A2[i][0] == "":
            A2[i][1] = "---"

    logging.info("Schedule updated")


def send_next_lesson(g1, g2):
    delete_message()
    if g1[0] != "":
        for i in allowedusers:
            if i[1] == 1:
                m1 = bot.send_message(i[0], g1[0] + " в " + g1[1]).message_id
                logging.info("lesson sended to " + str(i[0]))
                del_array.append([i[0], m1])
    if g2[0] != "":
        for i in allowedusers:
            if i[1] == 2:
                m2 = bot.send_message(i[0], g2[0] + " в " + g2[1]).message_id
                logging.info("lesson sended to " + str(i[0]))
                del_array.append([i[0], m2])


def send_schedule():
    now = datetime.datetime.now()
    current_time = now.strftime("%H:%M")
    weekday = now.weekday()
    if weekday > 4:
        sleep(3600 * 5)
    else:
        if current_time == "06:30":
            update_schedule()
            get_schedule(weekday)
            update_users()
            for i in allowedusers[1:]:
                if i[1] == 1:
                    m1 = bot.send_message(i[0],
                                          "Расписание на сегодня:\n1. " + A1[0][0] + " [" + A1[0][
                                              1] + "]\n2. " +
                                          A1[1][0] + " [" + A1[1][1] + "]\n3. " + A1[2][0] + " [" +
                                          A1[2][1] + "]\n4. "
                                          + A1[3][0] + " [" + A1[3][1] + "]\n").message_id
                    logging.info("morning schedule sended to " + str(i[0]))
                    del_array2.append([i[0], m1])
                elif i[1] == 2:
                    m2 = bot.send_message(i[0],
                                          "Расписание на сегодня:\n1. " + A2[0][0] + " [" + A2[0][
                                              1] + "]\n2. " +
                                          A2[1][0] + " [" + A2[1][1] + "]\n3. " + A2[2][0] + " [" +
                                          A2[2][1] + "]\n4. "
                                          + A2[3][0] + " [" + A2[3][1] + "]\n").message_id
                    logging.info("morning schedule sended to " + str(i[0]))
                    del_array2.append([i[0], m2])
            sleep(60)
        elif current_time == "08:50":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[0]
            g2 = A2[0]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "10:25":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[1]
            g2 = A2[1]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "12:15":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[2]
            g2 = A2[2]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "14:05":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[3]
            g2 = A2[3]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "16:00":
            delete_message()
            delete_message2()
            sleep(60)


while True:
    try:
        send_schedule()
        sleep(30)
    except Exception as e:
        logging.error("sending error " + str(e))
        sleep(20)
