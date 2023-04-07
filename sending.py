import telebot
import xlrd
import os
from time import sleep
import datetime
import wget
import logging

# test comment
bot = telebot.TeleBot('')
logging.basicConfig(filename="logs.log", level=logging.INFO, format=' %(asctime)s - %(levelname)s - %(message)s')

LINK = ""

A1 = [["" for i in range(2)] for j in range(4)]
A2 = [["" for k in range(2)] for l in range(4)]
allowedusers = [[0, 0]]


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


def isMerged(r, c):
    wb1 = xlrd.open_workbook("Schedule.xlsx")
    sheet1 = wb1.sheet_by_index(0)
    m = sheet1.merged_cells
    return (r, r + 1, c - 1, c + 1) in m


def get_schedule(weekday):
    wb = xlrd.open_workbook("Schedule.xlsx")
    sheet = wb.sheet_by_index(0)
    if weekday >= 1:
        x = 1
    else:
        x = 0
    for i in range(4):
        k = sheet.row_values(3 + 5 * weekday + i + x)[60]
        if isMerged(3 + weekday * 5 + i + x, 59):
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
                A1[i][0] = l1[:l2.find("(")].strip()
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
    if g1[0] != "":
        for i in allowedusers:
            if i[1] == 1:
                bot.send_message(i[0], g1[0] + " в " + g1[1])
            logging.info("lesson sended to" + str(i[0]))
    if g2[0] != "":
        for i in allowedusers:
            if i[1] == 2:
                bot.send_message(i[0], g2[0] + " в " + g2[1])
            logging.info("lesson sended to" + str(i[0]))


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
                    bot.send_message(i[0],
                                     "Расписание на сегодня:\n1. " + A1[0][0] + " [" + A1[0][
                                         1] + "]\n2. " +
                                     A1[1][0] + " [" + A1[1][1] + "]\n3. " + A1[2][0] + " [" +
                                     A1[2][1] + "]\n4. "
                                     + A1[3][0] + " [" + A1[3][1] + "]\n")
                elif i[1] == 2:
                    bot.send_message(i[0],
                                     "Расписание на сегодня:\n1. " + A2[0][0] + " [" + A2[0][
                                         1] + "]\n2. " +
                                     A2[1][0] + " [" + A2[1][1] + "]\n3. " + A2[2][0] + " [" +
                                     A2[2][1] + "]\n4. "
                                     + A2[3][0] + " [" + A2[3][1] + "]\n")
                logging.info("morning schedule sended to" + str(i[0]))
            sleep(60)
        elif current_time == "08:50":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[0]
            g2 = A2[0]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "10:30":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[1]
            g2 = A2[1]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "12:20":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[2]
            g2 = A2[2]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "14:10":
            update_schedule()
            get_schedule(weekday)
            update_users()
            g1 = A1[3]
            g2 = A2[3]
            send_next_lesson(g1, g2)
            sleep(60)


while True:
    send_schedule()
    sleep(30)
