import xlwt
book = xlwt.Workbook(encoding="utf-8")

# Add a sheet to the workbook
sheetw = book.add_sheet("users")

# Write to the sheet of the workbook
sheetw.write(1, 0, "This is the First Cell of the First Sheet")

# Save the workbook
book.save("spreadsheet.xls")