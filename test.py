import xlrd

A1 = [["" for i in range(2)] for j in range(4)]
A2 = [["" for k in range(2)] for n in range(4)]


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
            if A1[i][0]=="":
                A1[i][1] = "---"
            if A2[i][0]=="":
                A2[i][1] = "---"




get_schedule(0)
print(A1)
print(A2)
