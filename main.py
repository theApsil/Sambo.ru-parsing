import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl


def fix_img(tag):
    str_tag = str(tag)
    if 'img' in str_tag:
        start_index = str_tag.find('<td')
        end_index = str_tag.find('</td>') + len('</td>')
        output_text = str_tag[:start_index] + str_tag[end_index:]
        str_tag = output_text

    if '<a' in str_tag:
        soup = BeautifulSoup(str_tag, 'html.parser')
        str_tag = soup.find('a').text
    else:
        soup = BeautifulSoup(str_tag, 'html.parser')
        str_tag = soup.find('td').text.strip()

    return str_tag


def parse_sambo_events(city, start_year, end_year):
    data_collection = []

    for year in range(start_year, end_year + 1):
        url = f'https://www.sambo.ru/events/{year}/'
        response = requests.get(url)

        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            events = soup.find_all('tr', class_='item')

            for event in events:
                event_date = event.find('td', class_='date').nobr.text.strip()
                event_name_tag = fix_img(event.find('td', class_='name'))

                event_location = event.find('td', class_='location').nobr.text.strip()
                if city.lower() in event_location.lower():
                    # print(f'{year}: {event_name_tag}, {event_date}, {event_location}')
                    data_collection.append(
                        {'Year': year, 'Дата проведения': event_date, 'Название мероприятия': event_name_tag,
                         'Расположение': event_location})

    return data_collection


city_input = input('Введите город: ')
start_year_input = int(input('Введите начальный год (2007 и выше): '))
end_year_input = int(input('Введите конечный год (2007 и выше): '))

data = parse_sambo_events(city_input, start_year_input, end_year_input)

excel_file = pd.ExcelWriter('sambo_events.xlsx', engine='openpyxl')

for year in range(start_year_input, end_year_input + 1):
    year_data = [event for event in data if event['Year'] == year]
    df = pd.DataFrame(year_data, columns=['Дата проведения', 'Название мероприятия', 'Расположение'])
    df.to_excel(excel_file, sheet_name=str(year), index=False)

excel_file._save()
print('Данные успешно сохранены в файл sambo_events.xlsx.')
