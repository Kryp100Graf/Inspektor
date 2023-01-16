# This is a sample Python script.
import morph as morph
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


# def print_hi(name):
# Use a breakpoint in the code line below to debug your script.
# print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
from docxtpl import DocxTemplate
import openpyxl
import pymorphy2

wb = openpyxl.load_workbook('!источник.xlsx')
ws = wb['Лист1']
table = ws.tables['Таблица1']
arrayOfValues = []
morph = pymorphy2.MorphAnalyzer()

for row in ws.values:
    listKeys = {'ogrn': row[0], 'inn': row[1], 'lastName': row[2], 'firstName': row[3], 'patronymic': row[4],
                'startOfActivities': row[5],
                'endOfActivities': row[6], 'okved': row[7], 'longOkved': row[8]}
    arrayOfValues.append(listKeys)
arrayOfValues.pop(0)
# name = (listKeys['lastName'] + ' ' + listKeys['firstName'] + ' ' + listKeys['patronymic']).title()


# print(arrayOfValues[0]['lastName'])
for listKeys in arrayOfValues:


    try:
        nameGent = (
                morph.parse(listKeys['lastName'])[0].inflect({'gent'}).word + " " + \
                morph.parse(listKeys['firstName'])[0].inflect({'gent'}).word + " " + \
                morph.parse(listKeys['patronymic'])[0].inflect({'gent'}).word).title()
    except AttributeError:
        nameGent = (listKeys['lastName'] + " " + listKeys['firstName'] + " " + listKeys['patronymic']).title()

    context = {'inn': listKeys['inn'], 'nameGent': nameGent, 'name':
        (listKeys['lastName'] + " " + listKeys['firstName'] + " " + listKeys['patronymic']).title()}

    doc = DocxTemplate("!МП шаблон.docx")
    doc.render(context)
    doc.save("МП " + context['name'] + ".docx")

    doc = DocxTemplate("!Решение шаблон.docx")
    doc.render(context)
    doc.save("Решение " + context['name'] + ".docx")

    doc = DocxTemplate("!Уведомление шаблон.docx")
    doc.render(context)
    doc.save("Уведомление " + context['name'] + ".docx")
