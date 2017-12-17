"""
Download .doc files from publication and convert tables to CSV.
   
Publication home:

    http://www.gks.ru/wps/wcm/connect/rosstat_main/rosstat/ru/statistics/publications/catalog/doc_1140086922125
    
Sample url:    

    http://www.gks.ru/bgd/regl/b17_01/IssWWW.exe/Stg/d10/1-0.doc
"""

import arrow    
import requests
from pathlib import Path

from word import doc2csv, from_csv

def download(url, path):
    path = str(path)
    r = requests.get(url.strip(), stream=True)
    with open(path, 'wb') as f:
        for chunk in r.iter_content(chunk_size=1024):
            # filter out keep-alive new chunks
            if chunk:
                f.write(chunk)

def url(year: int, month: int, pub: str):
    """
    Args:
       year(int)
       month(int)
       pub(str) - file base name like '1-0.doc'
    """    
    last_digits = year-2000
    month = str(month).zfill(2)
    return ('http://www.gks.ru/bgd/regl/' +             
            'b{}_01/'.format(last_digits) +            
            'IssWWW.exe/Stg/d{}/'.format(month) + 
            '{}.doc'.format(pub))
    
assert url(2017, 10, '1-0') == \
    'http://www.gks.ru/bgd/regl/b17_01/IssWWW.exe/Stg/d10/1-0.doc'

class Folder:
    root = Path(__file__).parent
    
    @staticmethod               
    def md(path):
        if not path.exists():
            path.mkdir(parents=True)
        return path    
               
    def __init__(self, year: int, month: int):
        self.path = self.root / 'data' / str(year) / str(month) 
    
    @property    
    def interim(self):
        return self.subfolder('interim')
    
    @property    
    def raw(self):
        return self.subfolder('raw')
        
    def subfolder(self, subfolder):
        return self.md(self.path / subfolder)
     
        
class DocFile:        
    def __init__(self, year: int, month: int, pub: str):
        self.year, self.month = year, month
        self.url = url(year, month, pub)
        self.path = Folder(year, month).raw / f'{pub}.doc'        

    @property
    def size(self):
        if self.path.exists():
            return int(round(self.path.stat().st_size / 1024, 0))

    def download(self):
        download(self.url, self.path)
        
    def to_csv(self, target):
        doc2csv(doc_path = self.path,
                csv_path = InterimCSV(self.year, self.month, target).path)                     
        

class InterimCSV:        
    def __init__(self, year: int, month: int, target: str):
        self.path = Folder(year, month).interim / f'{target}.csv'
    
    def from_csv(self):
        return from_csv(self.path)

#TODO:  change 'stable_files' dictionary 
    
#class File:
#    stable_files=dict(
#        main=('1-0', 'Основные экономические и социальные показатели'),
#        retail=('3-1', 'Розничная торговля'),
#        bop=('3-2', 'Внешняя торговля'),
#        pri=('4-0', 'Индекс цен и тарифов'),
#        cpi=('4-1', 'Потребительские цены'),
#        ppi=('4-2', 'Цены производителей'),
#        odue=('5-0', 'Просроченная кредиторская задолженность организаций'),
#        soc=('6-0', 'Уровень жизни населения'),
#        lab=('7-0', 'Занятость и безработица'),
#        dem=('8-0', 'Демография')
#)    
#    
#    section2_a = dict(
#        ip=('2-1-0', 'Индекс промышленного производства'),
#        mng=('2-1-1', 'Добыча полезных ископаемых'),
#        mnf=('2-1-2', 'Обрабатывающие производства'),
#        pwr=('2-1-3', 'Обеспечение электрической энергией, газом и паром;'
#                      'кондиционирование воздуха'),
#        wat=('2-1-4', 'Водоснабжение; водоотведение, организация сбора ' 
#                      'и утилизации отходов, деятельность по ликвидации '
#                      'загрязнений'),
#        agro=('2-2-1', 'Сельское хозяйство'),
#        wood=('2-2-2', 'Лесозаготовки'),
#        constr=('2-3', 'Строительство'),
#        trans=('2-4', 'Транспорт'),
#    )
#        
#    section2_b = dict(
#        gdp=('2-1', 'Валовой внутренний продукт'),
#        ip=('2-2-0', 'Индекс промышленного производства'),
#        mng=('2-2-1', 'Добыча полезных ископаемых'),
#        mnf=('2-2-2', 'Обрабатывающие производства'),
#        pwr=('2-2-3', 'Обеспечение электрической энергией, газом и паром;'
#                      'кондиционирование воздуха'),
#        wat=('2-2-4', 'Водоснабжение; водоотведение, организация сбора ' 
#                      'и утилизации отходов, деятельность по ликвидации '
#                      'загрязнений'),
#        agro=('2-3-1', 'Сельское хозяйство'),
#        wood=('2-3-2', 'Лесозаготовки'),
#        constr=('2-4', 'Строительство'),
#        trans=('2-5', 'Транспорт'),
#    )
#
#    def __init__(self, year: int, month: int, target: str):
#        self.target = target
#        if target in self.stable_files.keys():
#            self.postfix = self.stable_files[target][0]
#        elif month in [9, 8, 6, 5]:
#            self.postfix = self.section2_b[target][0]
#        else:
#            self.postfix = self.section2_a[target][0]
#        self.doc = DocFile(year, month, self.postfix)
#        self.csv = InterimCSV(year, month, self.target)
#        
#    def download(self):
#        self.doc.download()
#        
#    def to_csv(self):
#        self.doc.to_csv(self.target)
#
#    def from_csv(self):
#        return list(self.csv.from_csv())

def official_dates(): 
    # QUESTION: how far back in time can we run?
    start = arrow.get(2016, 1, 1)
    end = arrow.get(2017, 10, 1)
    for r in arrow.Arrow.range('month', start, end):
        yield r.year, r.month

if __name__ == "__main__":
    year, month = 2017, 10 
    d = DocFile(year, month, '1-0')
    d.download()
    d.to_csv('main')
