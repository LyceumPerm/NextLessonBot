import telebot
import os
from time import sleep
import datetime
import wget
import logging
import openpyxl

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
allowedusers = []
del_array = []
del_array2 = []


def find_sheet(wb):
    sheets = wb.sheetnames
    for sh in sheets:
        if MONDAY in sh:
            sheet = sh
            break
    return sheet


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
    ur = openpyxl.load_workbook("users.xlsx")
    sheet_r = ur.active
    n = sheet_r.max_row + 1
    del allowedusers[0:]
    for i in range(1, n):
        a = int(sheet_r.cell(i, 1).value)
        b = int(sheet_r.cell(i, 2).value)
        allowedusers.append([a, b])
    logging.info("users updated")


def is_merged(r, c, sheet1):
    m = sheet1.merged_cells
    merged = False
    for i in m:
        if str(i).endswith(sheet1.cell(r, c).coordinate):
            merged = True
    return merged


def get_schedule(date):
    wb = openpyxl.load_workbook("Schedule.xlsx")
    sheet = wb[find_sheet(wb)]
    for row in range(1, sheet.max_row + 1):
        if str(sheet.cell(row, 1).value)[:10] == date:
            start_row = row
    for column in range(1, sheet.max_column + 1):
        if str(sheet.cell(2, column).value)[:10] == "11мат1":
            start_column = column
    for i in range(4):
        k = sheet.cell(start_row + i, start_column + 2).value
        if k == None or sheet.cell(start_row + i, start_column + 2).font.strike:
            k = ""
        if is_merged(start_row + i, start_column + 1, sheet):
            l = sheet.cell(start_row + i, start_column).value
            if l == None or sheet.cell(start_row + i, start_column).font.strike:
                l = ""
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
            l1 = sheet.cell(start_row + i, start_column).value
            l2 = sheet.cell(start_row + i, start_column + 1).value
            if l1 == None or sheet.cell(start_row + i, start_column).font.strike:
                l1 = ""
            if l2 == None or sheet.cell(start_row + i, start_column + 1).font.strike:
                l2 = ""
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
    current_date = now.strftime("%Y-%m-%d")
    if weekday > 4:
        sleep(3600 * 5)
    else:
        if current_time == "06:30":
            update_schedule()
            get_schedule(current_date)
            update_users()
            for i in allowedusers:
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
            get_schedule(current_date)
            update_users()
            g1 = A1[0]
            g2 = A2[0]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "10:25":
            update_schedule()
            get_schedule(current_date)
            update_users()
            g1 = A1[1]
            g2 = A2[1]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "12:15":
            update_schedule()
            get_schedule(current_date)
            update_users()
            g1 = A1[2]
            g2 = A2[2]
            send_next_lesson(g1, g2)
            sleep(60)
        elif current_time == "14:05":
            update_schedule()
            get_schedule(current_date)
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
