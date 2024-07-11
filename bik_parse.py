import re
import json
from bs4 import BeautifulSoup
import pickle

with open('bik_base.xml', 'r', encoding="utf8") as f:
	file = f.read()

soup = BeautifulSoup(file, 'xml')
names = soup.findAll('bik')

with open('bik_base.pickle', 'wb') as fp:
	pickle.dump(dict(re.findall(r'bik="(.*?)".*?name="(.*?)"', str(names))), fp)

bik = "200000304"

with open('bik_base.pickle', 'rb') as fp:
	bik_mapper = pickle.load(fp)
	print(bik_mapper[bik])