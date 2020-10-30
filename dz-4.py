# Импортируем библиотеку pandas
import pandas, openpyxl, collections, time
#from openpyxl import load_workbook

# Читаем файл ексль и результат передаем в переменную excel_data
# Переменная excel_data имеет тип <class 'pandas.core.frame.DataFrame'>
excel_data = pandas.read_excel('logs.xlsx', sheet_name='log2')
print(excel_data)

browsers = []
# Преобразуем переменную excel_data в список словарей с помощью метода to_dict()
# Результат передаем в переменную excel_data_dict
excel_data_dict = excel_data.to_dict(orient='records')
# print(excel_data_dict)
# print()
for element in excel_data_dict:
    browsers.append(element['Браузер'])
print(browsers)
popular_browser = collections.Counter(browsers).most_common(7)
print()
print(popular_browser)
browser_list=[]
#browser = collections.namedtuple('browser', trend) #, month2, month3, month4, month5, month6, month7, month8, month9, month10, month11, month12)
for element in popular_browser:
    m1=0
    m2=0
    m3=0
    m4=0
    m5=0
    m6=0
    m7=0
    m8=0
    m9=0
    m10=0
    m11=0
    m12=0
    for i in range(len(excel_data_dict)):
        if excel_data_dict[i]['Браузер']==element[0]:
            jear_month_day = str(excel_data_dict[i]['Дата посещения'])
            if  jear_month_day[5]=='0' and jear_month_day[6]=='1': m1 +=1
            elif jear_month_day[5]=='0' and jear_month_day[6]=='2': m2 +=1
            elif jear_month_day[6]=='3': m3 +=1
            elif jear_month_day[6]=='4': m4 +=1
            elif jear_month_day[6]=='5': m5 +=1
            elif jear_month_day[6]=='6': m6 +=1
            elif jear_month_day[6]=='7': m7 +=1
            elif jear_month_day[6]=='8': m8 +=1
            elif jear_month_day[6]=='9': m9 +=1
            elif jear_month_day[6]=='0': m10 +=1
            elif jear_month_day[5]=='1' and jear_month_day[6]=='1': m11 +=1
            elif jear_month_day[5]=='2' and jear_month_day[6]=='2': m12 +=1
    browser = element[0]       
    trend = str(element[1])          
    x = [browser, trend, str(m1), str(m2), str(m3), str(m4), str(m5), str(m6), str(m7), str(m8), str(m9), str(m10), str(m11), str(m12)]
        #for j in range(len(x)):
    browser_list.append(x) #list(browser, trend, str(m1), str(m2), str(m3), str(m4), str(m5), str(m6), str(m7), str(m8), str(m9), str(m10), str(m11), str(m12)))
#print(browser_list)

wb = openpyxl.load_workbook(filename='report.xlsx')
sheet = wb['Лист1']
for i in range(len(browser_list)):
    for j in range(14):
        sheet.cell(row=i+5, column=j+1).value = browser_list[i][j]
wb.save(filename='report.xlsx')

print('Всё хорошо!')
