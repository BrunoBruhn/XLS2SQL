def getting(file_name):
    pointSheetObj = []
    import xlrd as xl
    TeamPointWorkbook = xl.open_workbook(file_name)
    pointSheets = TeamPointWorkbook.sheet_names()

    for i in pointSheets:
        pointSheetObj.append(tuple((TeamPointWorkbook.sheet_by_name(i),i)))

    return pointSheetObj

