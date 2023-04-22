# # für den Download folgendermaßen vorgehen
#import urllib.request
import requests
from bs4 import BeautifulSoup
from pathlib import Path
from dataclasses import dataclass, field
import pandas as pd
from datetime import date, datetime
import re

@dataclass
class DataClassEdiDoc:
    domain: str = ""
    description: str = ""
    format: str = ""
    type: str = ""
    from_date: date = None
    to_date: date = None
    url: str = ""
    filename: Path = None
    view: str = ""

    def download(self, subdirectory):
        f_path = Path.cwd() / "EDI_ENERGY" / subdirectory / self.domain / self.filename
        r = requests.get(self.url)
        print(f_path)
        Path(f_path.parent).mkdir(parents=True, exist_ok=True)
        with open(f_path, 'wb') as f:
            f.write(r.content)

filter_values = DataClassEdiDoc

def get_view(view):
    directory = []
    url = 'https://www.edi-energy.de/index.php'
    data = {'tx_bdew_bdew[view]': view, 'tx_bdew_bdew[group]' : '1', 'id': '38'}

    r = requests.post(url, data=data)

    idx = BeautifulSoup(r.text, 'html.parser')
    table = idx.find('table')
    for row in table.find_all('tr'):
        cols = row.find_all('td')
        if len(cols) != 4:
            if len(cols) >= 1:
                dom = str(cols[0].string)
            continue

        ediDoc = DataClassEdiDoc(view=view)
        ediDoc.domain = dom

        ediDoc.description = " ".join(cols[0].string.split())
        ediDoc.from_date = cols[1].string.strip()
        ediDoc.to_date = cols[2].string.strip()
        if ediDoc.to_date == 'Offen':
            ediDoc.to_date = '31.12.9999'
        href = cols[3].find('a')
        ediDoc.url = 'https://www.edi-energy.de' + cols[3].find('a')['href']
        r = requests.head(ediDoc.url)
        dispo = r.headers['Content-Disposition']
        ediDoc.filename = dispo[dispo.find('=') + 2:-1]
        directory.append(ediDoc)
        if re.match('^[A-Z]{6} ', ediDoc.description):
            ediDoc.format = " ".join(re.findall('[A-Z]{6} ', ediDoc.description)).strip()
            search = re.search('MIG|AHB', ediDoc.description)
            if search:
                ediDoc.type = search.group()
        else:
            ediDoc.format = " "

    return directory

def load_web_directory_to_excel():
    directory = []
    directory.extend(get_view('now'))
    directory.extend(get_view('future'))
    directory.extend(get_view('archive'))

    path = Path('/EDI@ENERGY/EDI_ENERGY')
    today = datetime.today().strftime('%Y-%m-%d')
    file_name = path / f'EDI_Directory_{today}.xlsx'
    pd.DataFrame(directory).to_excel(file_name, sheet_name='EDI_Directory', index=False)
    # get_directory_from_web().to_excel(file_name, sheet_name='EDI_Directory', index=False)

def conv_date(date):
    """
    Wenn man die CSV Datei mit Excel speichert werden die Datumsfelder im Format dd.mm.yy gespeichert.
    Das wird hier bei Bedarf wieder in das Format dd.mm.yyyy zurück verwandelt. Dabei wird davon ausgegangen, dass
    das Jahr 99 im 99 Jahrhundert liegt (31.12.9999)
    :param date:
    :return:
    """
    if len(date) == 8:
        if date[6:8] == '99':
            cc = '99'
        else:
            cc = '20'
        date = date[:6] + cc + date[6:8]
    return date

def get_directory_from_excel(file):
    dir = []
    df = pd.read_excel(file, sheet_name='EDI_Directory')

    for ediDoc in df.to_dict(orient='records'):
        ediDoc['from_date'] = conv_date(ediDoc['from_date'])
        ediDoc['to_date'] = conv_date(ediDoc['to_date'])
        dir.append(DataClassEdiDoc(**ediDoc))
    return dir

def fltr(ediDoc):
    if filter_values.from_date != None:
        from_date = datetime.strptime(ediDoc.from_date, '%d.%m.%Y').date()
        to_date = datetime.strptime(ediDoc.to_date, '%d.%m.%Y').date()
        if from_date > filter_values.from_date:
            return False

        if to_date < filter_values.from_date:
            return False

    if filter_values.view != "":
        if ediDoc.view != filter_values.view:
            return False
    return True

def down_load_files(file):
    dir = get_directory_from_excel(file)
    # filter_values.from_date = date(2022,10,1)
    filter_values.view = 'future'

    for ediDoc in filter(fltr, dir):
        # ediDoc.download(filter_values.from_date.strftime('%Y%m%d'))
        ediDoc.download('20231001')

def main():
    #load_web_directory_to_excel()
    #down_load_files('/EDI@ENERGY/EDI_ENERGY/EDI_Directory_2023-04-12.xlsx')
    down_load_files('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/EDI_Directory_2023-04-12.xlsx')

if __name__ == '__main__':
    main()