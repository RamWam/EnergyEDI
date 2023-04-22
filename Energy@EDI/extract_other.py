from bs4 import BeautifulSoup
import re
import pandas as pd
from pathlib import Path
import logging
import util

def get_table_index(edi_doc):
    change_hist = []
    for t in edi_doc.find_all('table'):
        first_row = t.find('tr')
        first_col = first_row.find('td')
        style_dict_first_col = util.get_style_dict(first_col)

        first_text = util.get_text(first_col).replace('\n', '').strip()

        # Prüfen, ob die erste Spalte grau hinterlegt ist
        if style_dict_first_col.get('background', ' ') == '#D8DFE4':
            # Lesen des ersten Textes in der ersten Spalte
            if first_text.startswith('Änd-ID'):
                change_hist.append(t)

    return change_hist

def extract_other(path):
    with open(path, 'r') as f:
        data = f.read()

    edi_doc = BeautifulSoup(data, 'html.parser')
    table_index = get_table_index(edi_doc)

    new_file = path.parent / f'{path.stem}.xlsx'
    writer = pd.ExcelWriter(new_file)
    util.change_history_to_excel(table_index, writer)
    if writer.sheets == {}:
        return

    writer.close()

def main():
    paths = []

    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20221001'))
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230401'))
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20231001'))
    # paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230414'))
    for path in paths:
        for file in path.rglob('*.HTML'):
            if not '_MIG_' in file.stem \
               and not '_AHB_' in file.stem \
               and not 'EBD_' in file.stem:
                source = file.parent / file.name
                try:
                    extract_other(source)
                except BaseException as err:
                    logging.error(f'File {source}. Fehler {err}')
                    raise err

if __name__ == '__main__':
    main()
