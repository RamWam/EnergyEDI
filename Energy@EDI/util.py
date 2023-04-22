from dataclasses import dataclass, field
import pandas as pd

@dataclass
class DataClassChangeHistory:
    id: str = ""
    location: str = ""
    old: str = ""
    new: str = ""
    reason: str = ""
    status: str = ""

def change_history_to_excel(table_list, excel_writer):
    history = []
    for table in table_list:
        changeHistory = DataClassChangeHistory()
        skip_next = False
        old_id = " "
        for row in table.find_all('tr', recursive=False):
            # Zeile überspringen wenn vorherige Zeile mit "Änd-Id" begann
            if skip_next:
                skip_next = False
                continue
            columns = row.find_all('td',recursive=False)
            # Es gibt manchmal Zeilen mit 7 statt 6 Spalten. Dann ist die erste Spalte leer und kann
            # überlesen werden.
            if len(columns) == 7:
                columns.pop(0)

            id = get_text(columns[0]).strip()
            if id[:6] == 'Änd-ID':
                skip_next = True
                continue

            # Wenn die Id leer ist handelt, es sich um eine Fortsetzung einer
            # auf der vorherigen Seite begonnenen Tabellenzeile.
            # Die einzelnen Spalten werden dann um den neuen Text erweitert.
            # Die Tabellenzeile wird erst der Liste hinzugefügt, wenn die Id nicht leer ist.
            # Zur Sicherheit wird nach der For-Schleife noch ein Listeintrag hinzugefügt. Dieser
            # kann unter Umständen leer sein.
            if id == "":
                changeHistory.location = " ".join([changeHistory.location, get_text(columns[1])])
                changeHistory.old = " ".join([changeHistory.old, get_text(columns[2])])
                changeHistory.new = " ".join([changeHistory.new, get_text(columns[3])])
                changeHistory.reason = " ".join([changeHistory.reason, get_text(columns[4])])
                changeHistory.status = " ".join([changeHistory.status, get_text(columns[5])])
            else:
                if old_id != " ":
                    changeHistory.id = old_id
                    history.append(changeHistory)
                    changeHistory = DataClassChangeHistory()

                old_id = id
                changeHistory.location = get_text(columns[1])
                changeHistory.old = get_text(columns[2])
                changeHistory.new = get_text(columns[3])
                changeHistory.reason = get_text(columns[4])
                changeHistory.status = get_text(columns[5])
        changeHistory.id = old_id
        history.append(changeHistory)

    # Wenn keine Änderungshistorie vorhanden wird ein leerer Satz in die Tabelle geschrieben
    if history == []:
        history.append(DataClassChangeHistory())
    df = pd.DataFrame(history)
    rn_dict = {'id': 'Änd-ID', \
               'location': 'Ort', \
               'old': 'Bisher', \
               'new': 'Neu', \
               'reason': 'Grund der Anpassung', \
               'status': 'Status'}
    df.rename(columns=rn_dict, inplace=True)
    df.to_excel(excel_writer, sheet_name='Änderungshistorie', index=False)

def get_style_dict(tag):
    style_dict = {}
    if 'style' in tag.attrs:
        style_dict = dict([tuple(x) for x in [s.strip().split(':') for s in tag['style'].split(';')]])
    return style_dict

def get_text(column):
    """
    Konkateniert alle Strings in den Paragraphen einer Tabellenspalte
    :param column:
    :param index:
    :return:
    """
    return " ".join([string.replace(u'\xA0', ' ').strip() for para in column.find_all('p') for string in para.strings]).strip()

def get_string(tag):
    str = "".join(tag.strings).strip()
    str = str.replace('\xA0', ' ')
    str = str.replace('\n', ' ')
    str = str.replace('\r', ' ')
    return str

