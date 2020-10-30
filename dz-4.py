# Импортируем библиотеку pandas
import pandas, openpyxl, collections
#from openpyxl import load_workbook

def popular(listing, pop_name, quantity):
    pop_name_listing = collections.Counter(listing).most_common(quantity)
    pop_list=[]
    for element in pop_name_listing:
        m=[0]*13
        m[0]=element[1]
        for i in range(len(excel_data_dict)):
            if element[0] in excel_data_dict[i][pop_name]:
                jear_month_day = str(excel_data_dict[i]['Дата посещения'])
                m[int(jear_month_day[5:7])] +=1
                
        x = [element[0], m[0], m[1], m[2], m[3], m[4], m[5], m[6], m[7], m[8], m[9], m[10], m[11], m[12]]
        pop_list.append(x)
        print(x)
    return pop_list


def excel_write(text, row_begin_records):
    wb = openpyxl.load_workbook(filename='report.xlsx')
    sheet = wb['Лист1']
    for i in range(len(text)):
        for j in range(14):
            sheet.cell(row=i+row_begin_records, column=j+1).value = text[i][j]
    wb.save(filename='report.xlsx')
    return


# Читаем файл ексль и результат передаем в переменную excel_data
# Переменная excel_data имеет тип <class 'pandas.core.frame.DataFrame'>
excel_data = pandas.read_excel('logs.xlsx', sheet_name='log2')
print(excel_data)

# Преобразуем переменную excel_data в список словарей с помощью метода to_dict()
# Результат передаем в переменную excel_data_dict
excel_data_dict = excel_data.to_dict(orient='records')
browsers = []
items = []
for element in excel_data_dict:
    browsers.append(element['Браузер'])
    s = element['Купленные товары'].split(',')
    for k in s:
        items.append(k.strip())
       
browser_list = popular(browsers, 'Браузер', 7)
excel_write(browser_list, 5)

item_list = popular(items, 'Купленные товары', 7)
excel_write(item_list, 19)

print('Всё хорошо!')
