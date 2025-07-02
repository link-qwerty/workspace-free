# Imports
import csv
import json

# Defines
json_content = dict()

# Code

# Читаем файл csv
with open("export.csv", "r", encoding="utf-8") as csv_file:
    csv_content = csv.DictReader(csv_file)

    # Заполняем словарь для записи в json
    for row in csv_content:
        json_content[row['Почта'].strip()] = {'Название компании': row['Название компании'].strip().replace('\"', "'"),
                                      'Имя': row['Имя'].strip(), 'Проект': row['Проект'].strip().replace('\"', "'"),
                                      'Ссылка': row['Название компании'].replace(' ', '%20').strip().replace('\"', "'")}
csv_file.close()

# Записываем данные в json
with open("persons.json", "w", encoding="utf-8") as json_file:
    json.dump(json_content, json_file, ensure_ascii=False, indent=4)
json_file.close()