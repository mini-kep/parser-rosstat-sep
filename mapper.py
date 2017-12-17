from getter import InterimCSV

year, month = 2017, 10

COLUMNS_DEFAULT = {1:'bln_rub', 2:'yoy', 3:'rog'}

d1 = dict(
    substring = 'Индекс промышленного производства',
    name = 'INDPRO'
)

d2 = dict(
    substring = 'Продукция сельского хозяйства',
    name = 'AGROPROD'
)

rows = InterimCSV(year, month, 'main').from_csv()

# TODO: write pseudocode for mapper()
def mapper(year, month, d):
    pass

assert mapper(year, month, d1) == [
    {'name': 'INDPRO_yoy', 'date': '2017-10', 'value': 100.0},         
    {'name': 'INDPRO_rog', 'date': '2017-10', 'value': 105.7},         
]

assert mapper(year, month, d2) == [
    {'name': 'AGROPROD_bln_rub', 'date': '2017-10', 'value': 733.8},         
    {'name': 'AGROPROD_yoy', 'date': '2017-10', 'value': 97.5},         
    {'name': 'AGROPROD_rog', 'date': '2017-10', 'value': 64.7},
]         
