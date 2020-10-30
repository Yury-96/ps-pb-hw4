# Импортируем библиотеку pandas
import pandas, openpyxl, collections
#from openpyxl import load_workbook

# Читаем файл ексль и результат передаем в переменную excel_data
# Переменная excel_data имеет тип <class 'pandas.core.frame.DataFrame'>
excel_data = pandas.read_excel('logs.xlsx', sheet_name='log2')
print(excel_data)

browsers = []
# Преобразуем переменную excel_data в список словарей с помощью метода to_dict()
# Результат передаем в переменную excel_data_dict
excel_data_dict = excel_data.to_dict(orient='records')

for element in excel_data_dict:
    browsers.append(element['Браузер'])

popular_browser = collections.Counter(browsers).most_common(7)
browser_list=[]
for element in popular_browser:
    m=[0]*13
    m[0]=element[1]
    for i in range(len(excel_data_dict)):
        if excel_data_dict[i]['Браузер']==element[0]:
            jear_month_day = str(excel_data_dict[i]['Дата посещения'])
            m[int(jear_month_day[5:7])] +=1
            
    browser = element[0]       
    x = [browser, m[0], m[1], m[2], m[3], m[4], m[5], m[6], m[7], m[8], m[9], m[10], m[11], m[12]]
    browser_list.append(x)
    #print(x)

wb = openpyxl.load_workbook(filename='report.xlsx')
sheet = wb['Лист1']
for i in range(len(browser_list)):
    for j in range(14):
        sheet.cell(row=i+5, column=j+1).value = browser_list[i][j]
wb.save(filename='report.xlsx')

print('Всё хорошо!')
