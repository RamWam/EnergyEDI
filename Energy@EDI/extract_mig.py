from bs4 import BeautifulSoup
import re
from collections import namedtuple
import pandas as pd
from pathlib import Path
import logging
import util

MessageStrct = namedtuple('MessageStruct',
                              'Counter \
                               Number \
                               Qual \
                               StatusStd \
                               StatusBdew \
                               RepetitionStd \
                               RepetitionBdew \
                               Level \
                               Content \
                               Note \
                               Example')

def get_message_struct(tab_list):
    """
    Liest die Nachrichtenstruktur aus MIG Handbuch
    :param edi_doc: bs4 beautiful Soup
    :return: dataframe mit extrahierten Daten
    """
    match_line = re.compile('^        \d\d\d\d')
    messageStrct = []
    for table in tab_list:
        # Die Daten stecken in den P-Tags
        for table_line in table.find_all('p'):
            line_string = "".join(table_line.strings).replace(u'\xa0', u' ').replace(u'\n', u'')
            if match_line.match(line_string):
                elements = line_string.split()

                counter = elements.pop(0)
                if elements[0][0].isnumeric():
                    number = elements.pop(0)
                else:
                    number = 0

                qual = elements.pop(0)
                statusStd = elements.pop(0)
                statusBdew = elements.pop(0)
                repetitionStd = elements.pop(0)
                repetitionBdew = elements.pop(0)
                level = elements.pop(0)
                content = ' '.join(elements)

                messageStrct.append(
                    MessageStrct(counter, number, qual, statusStd, statusBdew, repetitionStd, repetitionBdew, level,
                                 content, "", ""))
    return pd.DataFrame(messageStrct)

def get_intend(column):
    """
    Ermittelt, wie weit der Text in der Spalte eingerückt ist.
    :param column:
    :return: z.B. 7.2pt
    """
    intend_set = set()
    for para in column.find_all('p'):
            for style_option in para['style'].split(';'):
                key, value = style_option.split(':')
                if key == 'margin-left':
                    intend_set.add(value)
    if len(intend_set) == 1:
        return list(intend_set)[0]

def get_desc_as_string_list(col):
    """
    Die Texte der Beschreibungsspalte des Segmentlayout werden als Liste von Tupeln
    zurückgegeben. Die Tupel enthalten Kennzeichen ob der Text fett ist und den Text selber
    :param col:
    :return: Tupel
    """
    bold_flags = []
    lines = []
    for p in col.find_all('p'):
        if p.find('b'):
            bold = True
        else:
            bold = False

        s = "".join(p.strings)
        for line in s.splitlines():
            lines.append(line)
            bold_flags.append(bold)
    return list(zip(bold_flags, lines))

def get_table_no(rows):
    """
    Liest die einleitenden Tabellenzeilen bis zur Tabellennummer und gibt diese zurück
    :param rows:
    :return:
    """
    #  Bevor die Tabelle mit der Datenelementbeschreibung beginnt, wird das betreffende
    #  Segment im Kontext der Nachrichtenstruktur dargestellt. Hier interessiert uns die Nummer in der letzten
    #  Zeile. Mit ihr kann man auf die Nachrichtenstruktur referenzieren.
    # re_segment_ref = re.compile(r'(^.*\d\d\d\d)(.*)(\d+)(.*)')
    re_segment_ref = re.compile(r'(^.*\d\d\d\d)(\s+)(\d+)(.*)')
    for row in rows:
        columns = row.find_all('td')
        first_col = util.get_text(columns[0])

        match = re_segment_ref.match(first_col)
        if match:
            # return int(match.group(3))
            return match.group(3)
            # segment_dict['no'] = int(match.group(3))
            # print(match.group(3))
            break

def is_element_line(row):
    cols = row.find_all('td')
    if len(cols) != 7:
        return False
    if util.get_text(cols[0]).startswith('Bez'):
        return False
    return True

def is_description_line(row):
    cols = row.find_all('td')
    if len(cols) == 1:
        if not ("".join(cols[0].strings).strip().startswith("Standard")):
            return True

def has_values(d):
    # return not (d['de'] == '' and d['deg'] == '')
    return  d['de'] != ''

def extract_description(col, de, qual, quals):
    """
    Liest aus der Beschreibungsspalte (= rechte Spalte) entweder
    die Beschreibung des Datenelements oder die Liste der Qualifier

    Die Texte der Beschreibungspalte sind entweder:
      a) eine Beschreibung des Datenelementes
      b) Qualifier bestehend aus Qualifier und Bezeichnung z.B. UTILTS Netznutzungszeiten-Nachricht
      c) eine Beschreibung des Qualifiers
     Die Qualifier werden am Fettdruck erkannt.
     Die Unterscheidung zwischen Qualifier und Qualifierbezeichnung wird mit
     regulärem Ausdruck getroffen. Eine wesentliche Rolle spielen die drei Leerzeichen
     zwischen beiden Werten.

     Es handelt sich um c) wenn in der Spalte vorher schon ein Qualifier stand, sonst ist es a)

    :param col: Textspalte
    :param de: Segment Dict
    :param qual: Qualifier Dict
    :param quals: Qualifier Liste
    :return: (segment dict, qual dict, qualifier list)
    """
    re_qual = re.compile(r'(\s*)(\w*)(\s\s\s)(.*)', flags=re.DOTALL)
    for bold, string in get_desc_as_string_list(col):
        # Leerstrings überlesen
        if len(string.strip()) == 0:
            continue
        if bold:
            # Qualifier
            match = re_qual.match(string)
            if match:
                if has_values(qual):
                    quals.append(qual)
                    qual = init_qual()
                qual['no'] = de['no']
                if 'deg' in de:
                    qual['deg'] = de['deg']
                qual['de'] = de['de']
                qual['qual'] = match.group(2)
                qual['descr'] = ''
                descr = match.group(4)
                qual['beschreibung'] = ''
            else:
                descr = string
            qual['descr'] = " ".join((qual['descr'], descr.strip())).strip()
        else:
            if has_values(qual):
                qual['beschreibung'] = " ".join((qual['beschreibung'], string.strip())).strip()
            else:
                de['de_beschreibung'] = " ".join((de['de_beschreibung'], string.strip())).strip()

    return (de, qual, quals)

def init_de(table_no):
    return dict([('no', table_no),
                ('deg', ''),
                ('deg_name', ''),
                ('deg_edifact_status', ''),
                ('deg_bdew_status', ''),
                ('de', ''),
                ('de_name', ''),
                ('de_edifact_status', ''),
                ('de_edifact_format', ''),
                ('de_bdew_status', ''),
                ('de_bdew_format', ''),
                ('de_beschreibung', '')
                ])

def init_deg(table_no):
    return dict([('no', table_no),
                 ('deg', ''),
                 ('deg_name', ''),
                 ('deg_edifact_status', ''),
                 ('deg_bdew_status', '')])

def init_qual():
    return dict([('no', ''),
                 ('deg', ''),
                 ('de', ''),
                 ('qual', ''),
                 ('descr', ''),
                 ('beschreibung', ''),
                 ])

def traverse_table(table, message_struct):
    #ToDo besser strukturieren
    """
    Auslesen einer Tabelle mit Beschreibung eines Segmentes
    :param table:
    :return: Tuple mit Liste mit Felder und Liste mit Qualifiern
    """
    # segment_dict['no'] = get_table_no(rows)
    rows = table.find_all('tr').__iter__()
    table_no = get_table_no(rows)

    if table_no == None:
        print('Tabellennummer nicht gefunden')
        return

    segment_fields = []
    qual_fields = []
    deg_dict = init_deg(table_no)
    de_dict = init_de(table_no)

    # Danach Überlesen von Überschriften Zeilen und der einleitenden Zeile.
    rows.__next__()
    rows.__next__()
    rows.__next__()

    qual_dict = init_qual()
    # Alle Datenelementgruppen und Datenelemente lesen
    for row in filter(is_element_line, rows):
        columns = row.find_all('td')
        first_col = util.get_text(columns[0])

        if len(first_col) == 0:
            first_col = ' '

        if first_col[0] == ' ':
            de_dict, qual_dict, qual_fields = extract_description(col=columns[6],
                                                                  de=de_dict,
                                                                  qual=qual_dict,
                                                                  quals=qual_fields)
        elif first_col[0].isnumeric():
            # Datenelement

            # Verzögertes Wegschreiben des segment_dicts, da
            # sich eine Datenelement über mehrere Zeilen erstrecken kann.
            if has_values(de_dict):
                segment_fields.append(de_dict.copy())
                #qual_dict = init_qual()
                if has_values(qual_dict):
                    qual_fields.append(qual_dict)
                    qual_dict = init_qual()
                de_dict = init_de(table_no)
                de_dict.update(deg_dict)

            de_dict['de'] = first_col
            de_dict['de_name'] = util.get_text(columns[1])
            de_dict['de_edifact_status'] = util.get_text(columns[2])
            de_dict['de_bdew_status'] = util.get_text(columns[4])
            de_dict['de_edifact_format'] = util.get_text(columns[3])
            de_dict['de_bdew_format'] = util.get_text(columns[5])
            de_dict['de_beschreibung'] = ' '

            de_dict, qual_dict, qual_fields = extract_description(col = columns[6],
                                                                  de= de_dict,
                                                                  qual = qual_dict,
                                                                  quals=qual_fields)

            # Löschen der Datenelementgruppe, wenn es sich um ein Datenelement handelt, das nicht eingerückt ist
            if get_intend(columns[0]) != '7.2pt':
                de_dict.update(init_deg(table_no))
        else:
            # Datenelementgruppe
            deg_dict = init_deg(table_no)
            deg_dict['deg'] = first_col
            deg_dict['deg_name'] = util.get_text(columns[1])
            deg_dict['deg_edifact_status'] = util.get_text(columns[2])
            deg_dict['deg_bdew_status'] = util.get_text(columns[4])
            if not has_values(de_dict):
                de_dict.update(deg_dict)
    if has_values(qual_dict):
        qual_fields.append(qual_dict)

    if has_values(de_dict):
        segment_fields.append(de_dict.copy())

    rows = table.find_all('tr').__iter__()

    for row in filter(is_description_line, rows):
        col = row.find('td')
        text = "".join(col.strings)
        if 'Bemerkung' in text:
            break
    try:
        note = "".join(rows.__next__().find('td').strings)
        if "Beispiel" in "".join(rows.__next__().find('td').strings):
            example = "".join(rows.__next__().find('td').strings)
            message_struct.loc[table_no].Note = note
            message_struct.loc[table_no].Example = example
    except StopIteration:
        pass

    return (segment_fields, qual_fields)
    # return (segment_fields, qual_fields, message_struct)

def get_segmentlayout(tab_list, message_struct):
    """
    Liest die Segmentlayouts aus MIG Document
    :param edi_doc: BS4
    :return: Tupel mit Dataframe für Felder und Dataframe für Qualifier
    """
    fields = []
    quals = []

    for tab in tab_list:
        new_fields, new_quals = traverse_table(tab, message_struct)
        fields.extend(new_fields)
        quals.extend(new_quals)

    return (pd.DataFrame(fields), pd.DataFrame(quals))

def get_table_index(edi_doc):
    change_hist = []
    segment_layout = []
    message_structure = []

    for t in edi_doc.find_all('table'):
        first_row = t.find('tr')
        first_col = first_row.find('td')
        style_dict_first_col = util.get_style_dict(first_col)

        first_text = util.get_text(first_col).replace('\n', '').strip()

        # Prüfen, ob die erste Spalte grau hinterlegt ist
        if style_dict_first_col.get('background', ' ') == '#D8DFE4':
            if first_text.startswith('Status'):
                message_structure.append(t)
            elif first_text.startswith('Standard'):
                segment_layout.append(t)
            elif first_text.startswith('Änd-ID'):
                change_hist.append(t)

    return  {'MessageStructure': message_structure, 'SegmentLayout': segment_layout,  'ChangeHist': change_hist}

def extract_mig(path):
    with open(path, 'r') as f:
        data = f.read()

    edi_doc = BeautifulSoup(data, 'html.parser')
    table_index = get_table_index(edi_doc)

    new_file = path.parent / f'{path.stem}.xlsx'
    writer = pd.ExcelWriter(new_file)

    message_struct = get_message_struct(table_index['MessageStructure'])
    message_struct.insert(0, 'Format', path.stem[:6])
    message_struct = message_struct.set_index(['Number'], drop=False)

    df1, df2 = get_segmentlayout(table_index['SegmentLayout'], message_struct)

    rn_dict = {'Counter': 'Zähler',
               'Number': 'Nummer',
               'Qual': 'Bez',
               'StatusStd': 'Status Standard',
               'StatusBdew': 'Status BDEW',
               'RepetitionStd': 'MaxWdh Standard',
               'RepetitionBdew': 'MaxWdh BDEW',
               'Level': 'Ebene',
               'Content': 'Inhalt',
               'Note': 'Bemerkung',
               'Example': 'Beispiel'}
    message_struct.rename(columns=rn_dict, inplace=True)
    message_struct.to_excel(writer, sheet_name='Nachrichtenstruktur', index=False)

    df1.insert(0, 'Format', path.stem[:6])
    df2.insert(0, 'Format', path.stem[:6])
    rn_dict = {'no': 'Zähler',
               'deg': 'Datenelementgruppe',
               'deg_name': 'Datenelementgruppe Name',
               'deg_edifact_status': 'Datenelementgruppe EDIFACT Status',
               'deg_bdew_status': 'Datenelementgruppe BDEW Status',
               'de': 'Datenelement',
               'de_name': 'Datenelement Name',
               'de_edifact_status': 'Datenelement EDIFACT Status',
               'de_edifact_format': 'Datenelementgruppe EDIFACT Format',
               'de_bdew_status': 'Datenelementgruppe BDEW Status',
               'de_bdew_format': 'Datenelementgruppe BDEW Format',
               'de_beschreibung': 'Beschreibung'}
    df1.rename(columns=rn_dict, inplace=True)
    df1.to_excel(writer, sheet_name='Segmentlayout', index=False)

    rn_dict = {'no': 'Zähler',
               'deg': 'Datenelementgruppe',
               'de': 'Datenelement',
               'qual': 'Qualifier',
               'descr': 'Qualifier Beschreibung',
               'beschreibung': ''}
    df2.rename(columns=rn_dict, inplace=True)
    df2.to_excel(writer, sheet_name='Segmentlayout_Qual', index=False)

    #  Änderungshistorie
    util.change_history_to_excel(table_index['ChangeHist'], writer)

    writer.close()

def main():
    paths = []
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20231001'))
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230401'))
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20221001'))
    for path in paths:
        for file in path.rglob('*.HTML'):
            if '_MIG_' in file.stem:
                source = file.parent / file.name
                try:
                    extract_mig(source)
                except BaseException as err:
                    logging.error(f'File {source}. Fehler {err}')

if __name__ == '__main__':
    # ToDo Funktionsbeschreibungen überarbeiten. HowTo Lesen und Type Hints hinzufügen
    # ToDo Parsen des Dokumentes wie bei AHB Extrac
    main()



