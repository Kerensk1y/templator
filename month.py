import json
from bs4 import BeautifulSoup

months = {"01": 'января',
          "02": 'февраля',
          "03": "марта",
          "04": "апреля",
          "05": "мая",
          "06": "июня",
          "07": "июля",
          "08": "августа",
          "09": "сентября",
          "10": "октября",
          "11": "ноября",
          "12": "декабря"}


def xml2dict2json():
    with open('bik_base.xml', 'r', encoding='utf-8') as file:
        xml_data = file.read()

    soup = BeautifulSoup(xml_data, 'xml')

    elements = soup.find_all('bik')
    base = {}
    for element in elements:
        bik_value = element.get('bik')
        name_value = element.get('name')
        base[bik_value] = name_value

    print("Словарь составлен")
    with open('bik_base.json', 'w') as json_file:
        json.dump(base, json_file)
    print('json создан')


def json2dict(obj):
    with open('bik_base.json', 'r') as json_file:
        data = json.load(json_file)
    if data[obj]:
        return data[obj]
    else:
        pass


# xml2dict2json()
# json2dict('042520607')
