from bs4 import BeautifulSoup
from dataclasses import dataclass, field
import re
import pandas as pd
from pathlib import Path
import logging
import util

@dataclass
class DataClassEBD:
    """ drei ebenen:
    1. Festlegung (GPKE, MABIS, MPES, WIM Strom, Herkunftsnachweis, MeMi, Netzbetreiberwechsel....)
    2. Name Aktivitäts- oder Sequenzdiagramm (AD Kündigung,
    3. EBD oder Codeliste pro Anwendungsfall

    """
    requlation: str = ""
    activity: str = ""
    ident: str = ""
    role: str = ""
    step: str = ""
    step_text: str = ""
    step_level: str = ""
    result: str = ""
    code: str = ""
    note: str = ""
    comment: str = ""

@dataclass
class DataClassCodeList:
    requlation: str = ""
    activity: str = ""
    ident: str = ""
    code_list: str = ""
    code: str = ""
    usage: str = ""
    condition: str = ""
    name: str = ""

@dataclass
class DataClassTableIndex:
    type: str = ""
    h1: str = ""
    h2: str = ""
    h3: str = ""
    # Zwischenüberschrift
    h4: str = ""
    tabs: list = field(default_factory=list)

def has_white_col(row):
    return any([util.get_style_dict(c).get('background', 'white') == 'white' for c in row.find_all('td')])

def get_table_index(edi_doc):
    change_hist = list()
    ebd_tables = list()
    for t in edi_doc.find_all('table'):
        # Suchen nach Tabellen, bei denen in der ersten Spalte der ersten Zeile "Prüfende Rolle"
        # steht. Von dort aus werden die darüberliegenden Header und alle darunter liegenden
        # Tabellen bis zum nächsten Header gelesen.
        first_row = t.find('tr')
        first_col = first_row.find('td')
        head = util.get_string(first_col)

        if head.startswith('Änd-ID'):
            change_hist.append(t)
            continue

        table = DataClassTableIndex()
        if head.startswith('Prüfende Rolle'):
            table.type = 'EBD'
        if head.startswith('Code'):
            table.type = 'Codelist'

        if table.type != '':
            table.tabs.append(t)

            # Darüber liegenden Header lesen. Jeweils der erste h1, h2 und h3
            table.h1 = util.get_string(t.find_previous_sibling('h1'))
            table.h2 = util.get_string(t.find_previous_sibling('h2'))
            table.h3 = util.get_string(t.find_previous_sibling('h3'))

            # alle darunterliegenden Tabellen lesen bis zum nächsten Header
            for next_s in t.next_siblings:
                if next_s.name in ('h1', 'h2', 'h3'):
                    break
                if next_s.name == 'table':
                    # Aufhören, wenn Tabelle mit grau hinterlegte Zeile anfängt, es sich also "logisch" um
                    # eine neue Tabelle handelt.
                    if not has_white_col(next_s.find('tr')):
                        break
                    table.tabs.append(next_s)
            ebd_tables.append(table)

    return ebd_tables, change_hist

def get_offset(cols):
    re_step = re.compile(r'^[0-9]+\*?$')
    for key, col in cols.items():
        if col == None:
            return max(cols.keys())

        if re_step.match(util.get_string(col)):
            return key

def get_ebd(tab_list):
    ebd_list = []

    for tl in filter(lambda t: t.type == 'EBD', tab_list):
        role = ''
        for tab in tl.tabs:

            # Mit der rowspan Option kann sich in HTML eine Spalte über mehrere Zeilen erstrecken.
            # Beim Lesen einer neuen Zeile muss daher berücksichtigt werden,
            # welche Spalten aus der vorherigen Zeile übernommen werden müssen.
            # Das dict row_span enthält für jede Spalte die Anzahl der Zeilen über die es sich erstreckt.
            # Das dict cols merkt sich die Spalteninhalte aus den vorherigen Zeilen
            # Für beide dicts werden maximal 20 Einträge vorgesehen. Die EBD Tabellen haben in der Regel fünf
            # Spalten

            # Initialiseren des row_span dicts
            row_span = {i: 1 for i in range(20)}

            # Initialiseren des cols dicts
            cols = {i:None for i in range(20)}
            for row in tab.find_all('tr'):
                # Überlesen von Überschrifts- und Kommentarzeilen,
                # die komplett farbig (z.B. grau) hinterlegt sind.
                if not has_white_col(row):
                    first_col = row.find('td')
                    first_text = util.get_string(first_col)
                    if first_text.startswith('Prüfende Rolle:'):
                        role = first_text.split(':')[1].strip()
                    continue

                ebd = DataClassEBD()
                ebd.requlation = tl.h1
                ebd.activity = tl.h2
                ebd.ident = tl.h3
                ebd.role = role

                # Bei jeder neuen Zeile die Einträge in row_span um Eins vermindern
                row_span.update({key:value - 1 for key,value in row_span.items() if value > 0})

                # Kopie von row_span erzeugen mit den Spalten, für die in dieser Zeile
                # neue Werte erwartet werden.
                row_span_copy = {key:value for key, value in row_span.items() if value == 0}

                #Spalten initialisieren, die nicht aus Vorzeile übernommen werden sollen
                cols.update({key:None for key in row_span_copy.keys()})

                for col in row.find_all('td'):
                    for key, val in row_span_copy.items():
                        cols[key] = col
                        row_span_copy.pop(key)
                        row_span[key] = int(col.get('rowspan', 1))
                        break

                offset = get_offset(cols)
                # Nur Zeilen verwenden, die mindestens 5 Spalten haben
                if len(list(cols.values())) - list(cols.values()).count(None) - offset > 4:
                    ebd.step = util.get_string(cols[offset])
                    ebd.step_text = "".join(cols[offset+1].strings).strip()
                    ebd.step_text = util.get_string(cols[offset+1])
                    # Die Farbe der ersten Spalte sagt aus, ob die Prüfung auf
                    # Kopf-, Positions- oder Summenebene stattfinden soll.
                    match util.get_style_dict(cols[offset]).get('background', 'white'):
                        # Grau
                        case '#BFBFBF':
                            ebd.step_level = 'Kopf'
                        # Grün
                        case '#92D050':
                            ebd.step_level = 'Position'
                        # Gelb
                        case 'yellow':
                            ebd.step_level = 'Summe'

                    ebd.result = util.get_string(cols[offset+2])
                    ebd.code = util.get_string(cols[offset+3])
                    ebd.note = util.get_string(cols[offset+4])
                    ebd.comment = " ".join([util.get_string(col) for key, col in cols.items() if key < offset])
                    ebd_list.append(ebd)
    return ebd_list

def get_code_lists(tab_list):
    re_code_list = re.compile(r'[GS]{1,2}_\d{3,4}_')
    code_list_list = []

    for tl in filter(lambda t: t.type == 'Codelist', tab_list):
        for tab in tl.tabs:
            for row in tab.find_all('tr'):
                # Überlesen von Überschriftszeilen,
                # die komplett farbig (z.B. grau) hinterlegt sind.
                if not has_white_col(row):
                    # Wenn für die Aktivität schon ein EBD vorgesehen, aber noch nicht durch den BDEW umgesetzt ist,
                    # wird die temporär zu nutzenden Codeliste unmittelbar vor dem Tabellenkopf in einer Zwischenüberschrift
                    # oder in roter Schrift  oder auch ohne besondere Formatierung angegeben.
                    p = row.find_previous('p')
                    if re_code_list.match(util.get_string(p)):
                        cl = util.get_string(p)
                    else:
                        cl = ""

                    continue
                code_list = DataClassCodeList()
                code_list.requlation = tl.h1
                code_list.activity = tl.h2
                code_list.ident = tl.h3
                code_list.code_list = cl

                cols = row.find_all('td')

                code_list.code = util.get_string(cols[0])
                code_list.usage = util.get_string(cols[1])
                # Zwischen den Spalten "Nutzung" und "Name" können mehrere
                # Bedingungsspalten liegen. Diese werden mit den Trennzeichen "###" miteinander
                # verbunden
                code_list.condition = "\n###".join([util.get_string(col) for col in cols[2:-1]])

                code_list.name = util.get_string(cols[-1])

                # Codeliste aus h3 Überschrift extrahieren
                if code_list.code_list == "":
                    match = re_code_list.search(code_list.ident)
                    if match:
                        code_list.code_list = code_list.ident[match.span()[0]:]
                code_list_list.append(code_list)

    return code_list_list

def extract_ebd(path):
    with open(path, 'rb') as f:
        data = f.read()

    edi_doc = BeautifulSoup(data, 'html.parser')
    table_index, change_hist = get_table_index(edi_doc)

    new_file = path.parent / f'{path.stem}.xlsx'
    writer = pd.ExcelWriter(new_file)

    # EBD
    ebd_list = get_ebd(table_index)
    df = pd.DataFrame(ebd_list)
    rn_dict = {'requlation': 'Festlegung',
               'activity': 'Aktivitätsdiagramm',
               'ident': 'EBD',
               'role': 'Prüfende Rolle',
               'step': 'Nr.',
               'step_text': 'Prüfschritt',
               'result': 'Prüfergebnis',
               'code': 'Code',
               'note': 'Hinweis',
               'step_level': 'Prüfebene',
               'comment': 'Kommentar'}
    df.rename(columns=rn_dict, inplace=True)
    df.to_excel(writer, sheet_name='EBD', index=False)

    # Codelisten
    code_lists = get_code_lists(table_index)
    df = pd.DataFrame(code_lists)
    rn_dict = {'requlation': 'Festlegung',
               'activity': 'Aktivitätsdiagramm',
               'ident': 'EBD/Codeliste',
               'code_list': 'Codeliste',
               'code': 'Code',
               'usage': 'Nutzung',
               'condition': 'Bedingung',
               'name': 'Name'}
    df.rename(columns=rn_dict, inplace=True)
    df.to_excel(writer, sheet_name='CodeLists', index=False)

    #  Änderungshistorie
    util.change_history_to_excel(change_hist, writer)
    writer.close()

def main():
    paths = []
    #paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20231001'))
    #paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230401'))
    #paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20221001'))
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230419'))
    for path in paths:
        for file in path.rglob('*.HTML'):
            if 'EBD_' in file.stem:
                source = file.parent / file.name
                try:
                    extract_ebd(source)
                except BaseException as err:
                    logging.error(f'File {source}. Fehler {err}')
                    raise err

if __name__ == '__main__':
    main()