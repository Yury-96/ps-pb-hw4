def get_value_from_list(object_list, separator, key):
    """Функция находит значение ключа key из списка object_list
    по разделителю separator
    :param object_list: список строк
    :param separator: разделитель
    :param key: искомый ключ"""

    # Объявляем переменную для хранения найденного значения
    value = None
    for element in object_list:
        # Итерируемся по каждому элементу из переданного списка object_list.
        # Каждый элемент разделяем на части, используя разделитель separator. 
        # В итоге получим, что первый элемент - это ключ, второй элемент - это значение.
        words = element.split(separator)
        if words[0] == key:
            # Если первый элемент равен искому ключу key, то обновляем значение value и выходим из цикла
            value = words[1]
            break
    
    # Возвращаем найденное значение
    return value
    

# Журнал регистрации
log = """name:Иван;gender:m;item:Часы;item_cost:9800
name:Иван;gender:m;item:Фитнес-браслет;item_cost:12300
name:Иван;gender:m;item:Кофемашина;item_cost:23500
name:Петр;gender:m;item:Часы;item_cost:9800
name:Петр;gender:m;item:Фитнес-браслет;item_cost:12300
name:Петр;gender:m;item:Айфон;item_cost:77900
name:Петр;gender:m;item:Чехол для телефона;item_cost:350
name:Петр;gender:m;item:Кофемашина;item_cost:23500
name:Дарья;gender:m;item:Айфон;item_cost:77900
name:Марья;gender:m;item:Кофемашина;item_cost:23500
name:Юлия;gender:m;item:Фитнес-браслет;item_cost:12300"""

# Создаем итоговый список, в котором будем хранить разделенные записи
log_list = []

# Делим журнал регистрации log на список строк по разделителю - перенос строки.
# Полученный результат сохраняем в список records
records = log.split('\n')

# Обходим полученный список records
for log_record in records:
    
    # Определяем структуру записи, которую будем помещать в список log_list
    record_dict = {
        'name': '',
        'gender': '',
        'item': '',
        'item_cost': 0,
    }

    # Разделяем запись журнала log_record по точке с запятой - получаем список строк elements
    elements = log_record.split(';')

    # В списке elements ищем элемент по ключу и получаем значение этого элемента
    user_name = get_value_from_list(elements, ':', 'name')
    user_gender = get_value_from_list(elements, ':', 'gender')
    item_title = get_value_from_list(elements, ':', 'item')
    item_cost = get_value_from_list(elements, ':', 'item_cost')

    # Обновляем значения ключей в словаре record_dict
    record_dict['name'] = user_name
    record_dict['gender'] = user_gender
    record_dict['item'] = item_title
    record_dict['item_cost'] = item_cost

    # Добавляем полученный словарь record_dict в итоговый список log_list
    log_list.append(record_dict)
# Выводим итоговый список log_list
#print(log_list)

item_list = []
for k in log_list:
    if int(k['item_cost']) < 13000:
        if k['item'] not in item_list:
            item_list.append(k['item'])
            #print(f"{k['item']}: стоимость {k['item_cost']}")
print(f'Список товаров дешевле 13000: {item_list}')

