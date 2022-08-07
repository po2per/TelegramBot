import openpyxl

wb = openpyxl.reader.excel.load_workbook(filename="1.xlsx")
print(wb.sheetnames)
wb.active = 0
sheet = wb.active


# print(sheet['A1'].value)


def find_package(a):
    if type(a) is not int:
        return print("Не корректный ввод")
    searcher = False
    for i in range(1, 1000):
        if int(a) == sheet['C' + str(i)].value and sheet['B' + str(i)].value is not None and \
                sheet['H' + str(i)].value is not None:
            searcher = True
            return print("№", sheet['C' + str(i)].value, sheet['B' + str(i)].value, " трек номер:",
                         sheet['H' + str(i)].value)

        elif int(a) == sheet['C' + str(i)].value and sheet['B' + str(i)].value is not None and \
                sheet['H' + str(i)].value is None:
            searcher = True
            return print("№", sheet['C' + str(i)].value, sheet['B' + str(i)].value, "отправка состоится в ближайшее "
                                                                                    "время")

        elif int(a) == sheet['C' + str(i)].value and sheet['B' + str(i)].value is None:
            searcher = True
            return print("Ваш заказ находится в очереди для производства")

    if not searcher:
        return print("Заказ не обработан или не существует")


find_package(2015)
