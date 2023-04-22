from bs4 import BeautifulSoup
import re
import pandas as pd
from pathlib import Path
from dataclasses import dataclass, field
from itertools import filterfalse
import logging
import util

re_ahb_status = re.compile(r'Muss|Kann|Soll')
@dataclass
class DataClassContext:
    path: str = ''
    format: str = ''

# @dataclass
# class DataClassChangeHistory:
#     id: str = ""
#     location: str = ""
#     old: str = ""
#     new: str = ""
#     reason: str = ""
#     status: str = ""

@dataclass
class DataClassSegment:
    name: str = ""
    pids: dict() = field(default_factory=dict)
    cond: dict() = field(default_factory=dict)

@dataclass
class DataClassQualifier:
    qual: str = ""
    qual_descr: str = ""
    operands: list[str] = field(default_factory=list)

@dataclass
class DataClassDataElement:
    format: str = ""
    head: str = ""
    pid: str = ""
    seg_name: str = ""
    sgr: str = ""
    sgr_cond: str = ""
    seg: str = ""
    seg_cond: str = ""
    de: str = ""
    de_descr: str = ""
    qual: str = ""
    qual_descr: str = ""
    condition: str = ""
    raw_condition: str = ""
    error: str = ""
    raw: str = ""

def filter_white_rows(row):
    for col in row.find_all('td'):
        for style_option in col['style'].split(';'):
            key, value = style_option.split(':')
            if key.endswith('background'):
                if value == 'white':
                    return True
                else:
                    return False
    return False

def get_line(column):
    for para in column.find_all('p'):
        str = "".join(para.strings)
        if str != "":
            yield "".join(para.strings)

def get_pids_from_header(row):
    cols_it = iter(row.find_all('td'))
    # erste Spalte überlesen
    cols_it.__next__()
    col = cols_it.__next__()
    lines = util.get_text(col).splitlines()
    if len(lines)  != 0:
        # ToDo ist nicht ganz sauber. Besser wäre, es würden nur die fünfstelligen Zahlen am Ende der Zeile ausgewählt
        pid_tuple = tuple(re.findall(r'(\d{5,5})', lines[-1]))

    if len(pid_tuple) == 0:
        # Wenn keine PIDs vorhanden sind (APERAK, CNTRL) werden Pseudo Pids erstellt
        pid_tuple = ('PID_1', 'PID_2', 'PID_3', 'PID_4', 'PID_5',)
    return pid_tuple

def get_head_str(headers, sourceline):
    _, header = max([(h.sourceline, h) for h in headers if h.sourceline < sourceline])
    if header == None:
        return ""
    else:
        return " ".join(header.strings).strip().replace('\n', '')

def get_table_index(edi_doc):
    change_hist = []
    use_cases = {}
    pids = ()

    headers = list(edi_doc.find_all(re.compile("^h\d")))
    for t in edi_doc.find_all('table'):
        first_row = t.find('tr')
        # Lesen der ersten Spalte, die nicht leer ist.
        # Hintergrund: für die Änderungshistorie, wird manchmal eine zusätzliche Spalte
        # eingschoben, die überlesen werden muss
        col_it = first_row.find_all('td').__iter__()
        try:
            while True:
                first_col = col_it.__next__()
                first_text = util.get_text(first_col).replace('\n', '').strip()
                if first_text != '':
                    break
        except StopIteration:
            continue


        # first_col = first_row.find('td')
        style_dict_first_col = util.get_style_dict(first_col)

        first_text = util.get_text(first_col).replace('\n', '').strip()

        # Prüfen, ob die erste Spalte grau hinterlegt ist
        if style_dict_first_col.get('background', ' ') == '#D8DFE4':
            # Lesen des ersten Textes in der ersten Spalte
            if first_text.startswith('Änd-ID'):
                change_hist.append(t)

            # Header-Tabelle eines Anwendungsfalls
            elif re.match(r'^EDIFACT\s+Struktur', first_text):
                pids = get_pids_from_header(first_row)
                if pids not in use_cases:
                    head_str = get_head_str(headers, t.sourceline)
                    use_cases[pids] = (head_str, [])
                use_cases[pids][1].append(t)

        else:
            # Wenn Spalte nicht grau hinterlegt ist, die Anzahl der Spalten prüfen.
            # Das ist ein schwaches Kriterium, aber ein besseres fällt mir im Moment nicht ein,
            # um die Folgetabellen der Anwendungsfälle zu identifizieren
            if len(first_row.find_all('td')) == 3:
                if pids not in use_cases:
                    head_str = get_head_str(headers, t.sourceline)
                    use_cases[pids] = (head_str, [])
                use_cases[pids][1].append(t)

    return {'ChangeHist': change_hist, 'UseCases': use_cases}

def first_col_item(col):
    # Ordnet aus der ersten Spalte den letzten Bezeichner entweder
    # einer Segmentgruppe, einem Segment oder einen Datenelement zu.
    # Im AHB werden den Datenelementen ein oder zwei Segmentzeilen vorangestellt.
    # z.B. SG2
    #      SG2 NAD
    #      SG2 NAD 3035
    # (SG2=Segmentgruppe, NAD=Segment, 3035=Datenelement.
    # Wenn kein Zeichen fett ist, handelt es sich um die Segmentbeschreibung (z.B. Nachrichten-Kopfsegment)

    # Prüfen, ob es einen fetten Text gib
    has_bold = False
    for p in col.find_all('p'):
        for b in p.find_all('b'):
            has_bold = True
            break

    # Wenn kein fetter Text vorhanden handelt es sich um die Segmentbezeichnung
    if has_bold == False:
        t = util.get_text(col).replace('\n', '')
        if t != '':
            return('Segmentname', util.get_text(col).replace('\n', ''))

    tokens = util.get_text(col).replace('\n', '').split()
    # Prüfen, wie viele Bezeichner in der ersten Spalte stehen.
    match len(tokens):
        case 0:
            return ('', '')

        case 1:
            # Bei einem Bezeichner bezeichnet der letzte eine Segmentgruppe oder ein Segment
            if tokens[0].startswith('SG'):
                return ('Segmentgruppe', tokens[-1])
            else:
                return ('Segment', tokens[-1])

        # Bei zwei Bezeichnern, bezeichnet der letzte ein Segment oder ein Datenelement
        case 2:
            if tokens[0].startswith('SG'):
                return ('Segment', tokens[-1])
            else:
                return ('Dataelement', tokens[-1])

        case 3:
            # Bei drei Bezeichnern, bezeichnet der letzte ein Datenelement
            return ('Dataelement', tokens[-1])

        case _:
            raise Exception('Erste Spalte des Anwendungsfalls kann nicht interpretiert werden')

def segment_pids(pids, start_pos_ref, data_element):
    """
    Problem: Word ersetzt bei der Umwandlung in HTML die Tabulatoren zwischen den PID Bedingungen in Leerzeichen.
    Aufgrund der nichtproportional Schrift kann die Anzahl der Leerzeichen nicht ohne weiteres verwendet werden,
    um zu ermitteln, auf welche PID sich die Bedingung bezieht.

    Lösung:
    Die AHB-Status der Segmente werden ausgewertet, um zu ermitteln, auf welche PID sich die Bedingungen
    der folgenden Datenelemente beziehen. Im Unterschied zu den Datenelementzeilen
    liegt in den Segmentzeilen kein weiterer Beschreibungstext vorliegt, sondern nur die Texte Muss/Soll/Kann.

    Die Funktion liefert ein Tupel von PIDs, das anzeigt, für welche PID in den Datenelementzeilen Bedingungen vorliegen.
    Als Referenz für die Ausrichtung wird das (erste) UNH
    Segment verwendet werden, da hier immer für alle PIDs "Muss" angegeben ist. Für jeden AHB-Status
    werden die Zeichen vom Beginn der Zeile an gezählt und der entsprechende  AHb-Status aus der
    UNH-Referenzzeile ausgewählt, dessen Zeichenanzahl sich von diesem am wenigsten unterscheidet.
    :param pid_tuple_doc:
    :param unh_ref_line:
    :param data_element:
    :return:
    """
    # Startpositionen der AHB-Status in erster Segmentzeile ermitteln
    first_line = get_line(data_element).__next__()
    start_pos = [match.start() for match in re_ahb_status.finditer(first_line)]

    pids_seg = {}
    for sp in start_pos:
        seq = [(abs(pos - sp), pid) for pid, pos in start_pos_ref]
        seq.sort()
        dif = [dif for dif , _ in seq]
        # Pid mit der kleinsten Abweichung nehmen
        pid = seq[0][1]
        pids_seg[pid] = {'start': sp, 'dif': dif}

    return pids_seg

def split_pseudo_tab(str):
    # Sucht nach 3 aufeinanderfolgenden Leerzeichen. Wir gehen davon aus, dass bei der Umwandlung von
    # Word nach HTML die Tabs durch mindestens 3 Leerzeichen ersetzt werden
    re_pseudo_tab = re.compile(r'(\s{3,})')
    start_and_ends = [[match.start(), match.end()] for match in re_pseudo_tab.finditer(str)]
    new = []
    for se in start_and_ends:
        new.extend(se)
    if not new:
        return []
    new.pop(0)
    new.append(len(str))
    beg = new[::2]
    end = new[1::2]

    return [(str[b:e], b) for b, e in zip(new[::2], new[1::2])]

def seg_conditions(dataelement, segment_pids, start_pos_ref):
    """

    :param column: Bedingungen für Segmentgruppen oder Segmente
    :return: Liste mit Strings pro PID
    """

    pids_seg = {}
    for line in get_line(dataelement):
        for column in split_pseudo_tab(line):
            seq = [(abs(pos - column[1]), pid) for pid, pos in start_pos_ref]
            seq.sort()
            pid = seq[0][1]
            if pid in pids_seg:
                pids_seg[pid] = " ".join((pids_seg[pid], column[0]))
            else:
                pids_seg[pid] = column[0]
    return pids_seg

@dataclass
class DataClassCondition:
    name: str = ""
    descr: str = ""

@dataclass
class DataClassToken:
    type: str = ""
    pos: int = 0
    strng: str = ""
    line: int = 0

def get_de_token(dataelement):
    for line_no, p in enumerate(dataelement.find_all('p')):
        pos = 0

        # Style Optionen aus Paragraph in Dictionary schreiben
        style_dict = {'color': '', 'margin-left': '0pt'}
        if 'style' in p.attrs:
            style_dict.update(dict([tuple(x) for x in [s.strip().split(':') for s in p['style'].split(';')]]))

        look_ahead = " ".join(p.strings) + ' '
        look_ahead = look_ahead.replace('\n','')

        for ps in p.strings:
            # Style Optionen aus Parent ( in der Regel SPAN) in Dictionary schreiben
            if 'style' in ps.parent.attrs:
                style_dict.update(dict([tuple(x) for x in [s.strip().split(':') for s in ps.parent['style'].split(';')]]))

            # Fettdruck in Style_Dict aufnehmen
            if ps.parent.parent.name == 'b':
                style_dict['bold'] = True
            else:
                style_dict['bold'] = False

            # Manchmal werden die Operanden nicht durch Strings, sondern innerhalb eines Strings
            # durch Leerzeichen getrennt. Für diese Leerzeichenfolgen sollen einzelne Token erstellt werden.
            # Per Definition werden Token ab einer Folge von drei oder mehr Leerzeichen als Trenner erzeugt
            # re_split_tab = re.compile(r'(\s{3,})')
            re_split_tab = re.compile(r'(\s{3,})')
            str_list = re_split_tab.split(ps.replace('\n', ''))

            for s in str_list:
                if s == '':
                    continue
                strng = s.strip().replace('\n', '')

                if s.strip() == '':
                    type = 'empty'
                    strng = s
                elif style_dict['bold'] == True:
                    # Prüfen ob innerhalb des Paragraphen noch Operanden Definitionen beginnen.
                    # Dann handelt es sich um den Anfang eines neuen Qualifiers, ansonsten
                    # wird ein in der Vorzeile beginnnender Qualifier fortgesetzt.
                    # Kommt vor bei Anwendungsfällen mit vielen PIDs (z.B. 11116). Dort muss ein langer Qualifier
                    # (z.B. UITLMD) auf die nächste Zeile umgebrochen werden
                    re_operand = re.compile(r'X |M |S |K ')
                    # Zeile ab der aktuellen Position
                    rest_line = look_ahead[look_ahead.find(strng) + len(strng):]
                    if re_operand.search(rest_line) == None:
                        type = 'bold'
                    else:
                        type = 'bold_x'
                elif style_dict['color'] == 'gray':
                    type = 'gray'
                elif float(style_dict['margin-left'][:-2]) > 20:
                    type = 'indented'
                else:
                    if pos > 29:
                        type = 'plain_c'
                    else:
                        type = 'plain'
                    # Anhand der Position innerhalb einer Zeile wird unterschieden,
                    # ob es sich Texte für Bedingungen handelt oder nicht.
                    # Die Grenze wird bei Position 20 gezogen. Stichproben haben
                    # ergeben, dass keine X vor 30 liegen und die Anfänge der Qualifier Beschreiber
                    # höchstens bei 10. 20 liegt dazwischen
                    # if pos > 20:
                    if pos > 29:
                        type = 'plain_c'
                    else:
                        type = 'plain'

                    # # Es kann nicht mehr anhand der Position unterschieden werden,
                    # # ob es sich um Text für Bedingungen handelt oder nicht.
                    # # Stattdessen wird gegen die in Bedingungen möglichen Zeichen geprüft
                    # if all(c in u'0123456789[]P.XSMKUB∧∨⊻()usanol ' for c in strng.replace('\xa0', '')):
                    #     type = 'plain_c'
                    # else:
                    #     type = 'plain'


                yield DataClassToken(type, pos, strng, line_no)

                pos += len(s)
        yield DataClassToken('newline', 0, '', line_no)

def parse_conditions(token, tokens, conds):
    cond = ''
    while token.type in ('plain_c', 'indented'):
        cond = "".join((cond, token.strng))
        token = tokens.__next__()
        if token.type == 'empty':
            cond = "".join((cond, token.strng))
            conds.append((cond, token.strng, token.line))
            cond = ''
            token = skip_empty(token, tokens)
    if token.type != 'newline':
        raise De_syntax_error(f'Erwartet "newline", empfangen "{token}"')
    if cond != '':
        conds.append((cond, '', token.line))
    token = tokens.__next__()
    token = skip_empty(token, tokens)

    return token, conds

def map_conditons(conds, pids):
    new_conds = {}
    lines = {}
    for c in conds:
        if c[2] in lines:
            lines[c[2]] += 1
        else:
            lines[c[2]] = 1
    cnt_cnd = min(lines.values())
    if cnt_cnd == max(lines.values()) and cnt_cnd == len(pids):
        cnt_cnd = list(lines.values())[0]
        strngs = [s for s, _, _ in conds]
        new_conds = dict((list(pids.keys())[i], "".join(list(strngs[i::cnt_cnd]))) for i in range(cnt_cnd))
    old_line = 0
    raw = ''
    for c in conds:
        if c[2] != old_line:
            raw += '\n'
        raw += c[0]
    return {'cond':new_conds, 'raw':raw}

def skip_empty(token, tokens):
    while token.type == 'empty':
        token = tokens.__next__()
    return token

class De_syntax_error(Exception):
    pass

def parse_dataelement(dataelement, pids):
    """
        Es wird folgende Syntax verwendet:
        ALT----------------------
        Datenelement = gray empty* Condition [De_NextLine]*
        De_Nextline  = [gray empty*][Condition]
        Condition    = [plain|indented empty*]* newline

        Qualifier    = bold empty plain Condition [Qu_NextLine]*
        Qu_NextLine  = [indented] Condition

        Neu----------------------
        Datenelement = gray empty* Condition [De_NextLine]*
        De_Nextline  = [gray empty*][Condition]
        Condition    = [plain_c|indented empty*]* newline

        Qualifier    = bold_x empty plain Condition [Qu_NextLine]*
        Qu_NextLine  = [(bold empty* [plain] empty*) | (indented empty*)] Condition

        :param dataelement:
        :return:
        """
    # print('#')
    tokens = get_de_token(dataelement)
    details_list = []
    if "".join(dataelement.strings).find('19120') > 0:
        pass
    details = {}
    try:
        try:
            # while True:
            #     print(tokens.__next__())

            token = tokens.__next__()
            if token.type == 'gray':
                details['de'] = token.strng
                details['conds'] = []
                token = tokens.__next__()
                token = skip_empty(token, tokens)

                # while token.type == 'plain':
                token, details['conds'] = parse_conditions(token, tokens, details['conds'])

                while token.type in ('gray', 'plain_c'):
                    if token.type == 'gray':
                        details['de'] = " ".join((details['de'], token.strng))
                        token = tokens.__next__()
                        token = skip_empty(token, tokens)
                    token, details['conds'] = parse_conditions(token, tokens, details['conds'])

            # Qualifier
            elif token.type == 'bold_x':
                while token.type == 'bold_x':
                    details = {'qual': token.strng, 'descr':'', 'conds': []}
                    token = tokens.__next__()
                    if token.type != 'empty':
                        raise De_syntax_error(f'Erwartet "empty", empfangen "{token}"')
                    token = tokens.__next__()
                    if token.type != 'plain':
                        raise De_syntax_error(f'Erwartet "plain", empfangen "{token}"')
                    details['descr'] = token.strng
                    token = tokens.__next__()
                    while token.type == 'plain':
                        details['descr'] = ''.join((details['descr'],token.strng))
                        token = tokens.__next__()

                    token = skip_empty(token, tokens)
                    token, details['conds'] = parse_conditions(token, tokens, details['conds'])
                    # token = skip_empty(token, tokens)

                    while token.type in ('indented', 'plain_c', 'bold'):
                        match token.type:
                            case 'indented':
                                details['descr'] = " ".join((details['descr'], token.strng))
                                token = tokens.__next__()
                                token = skip_empty(token, tokens)

                            case 'bold':
                                details['qual'] = "".join((details['qual'], token.strng))
                                token = tokens.__next__()
                                token = skip_empty(token, tokens)
                                if token.type == 'plain':
                                    details['descr'] = " ".join((details['descr'], token.strng))
                                    token = tokens.__next__()
                                    token = skip_empty(token, tokens)

                        token, details['conds'] = parse_conditions(token, tokens, details['conds'])

                    details['conds'] = map_conditons(details['conds'], pids)
                    details_list.append(details)
                    details = {}

            raise De_syntax_error(f'Erwartet kein Token, empfangen "{token}"')

        except StopIteration:
            if details != {}:
                details['conds'] = map_conditons(details['conds'], pids)
                details_list.append(details)

    except De_syntax_error as err:
        logging.error(f'Syntaxfehler: {context} {err}, Pids: {pids}')
        details = {'error': str(err), 'raw': " ".join(dataelement.strings), 'conds': {'cond':{}}}
        details_list.append(details)

    return details_list

def get_segment_group(item, start_pos_ref, sgr_parent, pids, dataelement):
    sgr = DataClassSegment()
    sgr.name = item
    sgr.pids = segment_pids(pids, start_pos_ref, dataelement)
    sgr.cond = seg_conditions(dataelement, sgr.pids, start_pos_ref)
    if sgr_parent.name == sgr.name:
        if sgr.cond != {}:
            sgr_parent = sgr
    else:
        sgr_parent = sgr

    logging.debug('segmentgruppe:%s pids:%s cond:%s', sgr.name, sgr.pids, sgr.cond)
    logging.debug('segmentgruppe parent:%s pids:%s cond:%s', sgr_parent.name, sgr_parent.pids, sgr_parent.cond)
    return sgr, sgr_parent

def get_segment(item, start_pos_ref, pids, dataelement):
    seg = DataClassSegment()
    seg.name = item
    seg.pids = segment_pids(pids, start_pos_ref, dataelement)
    seg.cond = seg_conditions(dataelement, seg.pids, start_pos_ref)
    logging.debug('Segment:%s', seg)
    return seg

def get_dataelement(item, seg_name, sgr_parent, seg, dataelement, head_str):
    # details_list = parse_dataelement(dataelement, seg.pids)
    de_list = []
    for details in parse_dataelement(dataelement, seg.pids):
        if details['conds']['cond'] == {}:
            cur_pids = seg.pids
        else:
            cur_pids = details['conds']['cond']
        for pid in cur_pids.keys():
            de = DataClassDataElement(head=head_str, pid=pid)
            if 'error' in details:
                de.error = details['error']
                de.raw = details['raw']
                de_list.append(de)
                continue

            de.de = item
            if 'de' in details:
                de.de_descr = details['de']
            de.seg_name = seg_name
            de.sgr = sgr_parent.name
            if pid in sgr_parent.cond:
                de.sgr_cond = sgr_parent.cond[pid]
            de.seg = seg.name
            if pid in seg.cond:
                de.seg_cond = seg.cond[pid]
            if 'qual' in details:
                de.qual = details['qual']
                de.qual_descr = details['descr']
            if details['conds']['cond'] == {}:
                de.condition = '##nicht ermittelbar##'
            else:
                de.condition = details['conds']['cond'][pid]
            de.raw_condition = details['conds']['raw']
            de_list.append(de)

    return de_list

def get_condition(condtion):
    # ToDo Bedingungen funktionieren nicht, wenn Bedingungen über Seitenwechsel hinaus gehen
    conds = set()
    cnd = util.get_text(condtion)
    cnd = cnd.strip().replace('\n', '').replace('\t', '')
    re_split_cond = re.compile(r'(\[\d+\].*?)', flags=re.DOTALL)
    splits = re_split_cond.split(cnd)

    if len(splits) > 1:
        # Der erste Listeneintrag wird nicht verwendet. Deshalb startet der ZIP bei ein
        split_dict = dict(zip(splits[1::2], splits[2::2]))
        conds.update(split_dict.items())
    return conds

def get_use_case(head_str, tab_list, pids):
    logging.debug(f'Anwendungsfälle {head_str}')
    segments=[]
    conditions=set()

    sgr_parent = DataClassSegment()
    for table in tab_list:
        for row_no, row in enumerate(filter(filter_white_rows, table.find_all('tr'))):
            col_it = iter(row.find_all('td'))
            structure = col_it.__next__()
            dataelement = col_it.__next__()
            condition = col_it.__next__()

            match first_col_item(structure):
                case ('Segmentname', item):
                    seg_name = item
                    logging.debug('(expliziter)Segmentname:%s', seg_name)
                    continue

                case ('Segmentgruppe', item):
                    sgr, sgr_parent = get_segment_group(item, start_pos_ref, sgr_parent, pids, dataelement)

                case ('Segment',item):
                    if item in ('UNB', 'UNH'):
                        # Aus der ersten Zeile der Datenelementspalte werden die Startpositionen der
                        # Bedingungen ermittelt. Das UNH Segment enthält für alle PIDs des Anwendungsfalls
                        # Muss-Bedingungen. Die Startpositionen werden für Ermittlung der darauffolgenden
                        # Bedingungen von Segementgruppen, Segmenten, Datenelementen und Qualifiern genutzt.
                        start_pos_ref = list(zip(pids, [match.start() for match in re_ahb_status.finditer(get_line(dataelement).__next__())]))
                    seg = get_segment(item, start_pos_ref, pids, dataelement)

                case ('Dataelement',item):
                    de = get_dataelement(item, seg_name, sgr_parent, seg, dataelement, head_str)
                    #segments.extend(get_dataelement(item, seg_name, sgr_parent, seg, dataelement, head_str))
                    if seg.name == 'UNH':
                        if de[0].de  == '0065':
                            context.format = de[0].qual
                    segments.extend(de)

            conditions.update(get_condition(condition))
    for s in segments:
        s.format = context.format

    return segments, conditions

def extract_ahb(path):
    with open(path, 'r') as f:
        data = f.read()

    logging.debug('starte %s', path)

    edi_doc = BeautifulSoup(data, 'html.parser')
    table_index = get_table_index(edi_doc)
    new_file = path.parent / f'{path.stem}.xlsx'
    writer = pd.ExcelWriter(new_file)

    # Anwendungsfälle
    use_cases = []
    conds = set()
    for pids, tab_list in table_index['UseCases'].items():
        segments, conditions = get_use_case(tab_list[0], tab_list[1], pids)
        use_cases.extend(segments)
        conds.update(conditions)

    df = pd.DataFrame(use_cases)
    rn_dict = {'head':'Kapitel', \
               'pid':'PID',
               'seg_name':'Segmentname', \
               'sgr':'Segmentgruppe', \
               'sgr_cond':'Segmentgruppe Bedingung', \
               'seg':'Segment', \
               'seg_cond':'Segment Bedingung', \
               'de':'Datenelement', \
               'de_descr': 'Datenelement Beschreibung', \
               'qual': 'Qualifier', \
               'qual_descr':'Qualifier Beschreibung', \
               'condition':'Bedingung', \
               'raw_condition':'Bedingung (ohne PID Zuordnung)', \
               'error': 'Fehler', \
               'raw':'Fehlerdaten'}
    df.rename(columns=rn_dict, inplace=True)
    df.to_excel(writer, sheet_name='Anwendungsfälle', index=False)

    df = pd.DataFrame(conds)
    rn_dict = {0: 'Bedingung', \
               1: 'Beschreibung'}
    df.rename(columns=rn_dict, inplace=True)
    df.to_excel(writer, sheet_name='Bedingungen', index=False)

    #  Änderungshistorie
    util.change_history_to_excel(table_index['ChangeHist'], writer)

    # close Pandas Excel writer and output the Excel file
    writer.close()
    logging.info('%s', path)

def main():
    paths = []

    #paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20221001'))
    #paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230401'))
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20231001'))
    # paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230414'))
    for path in paths:
        for file in path.rglob('*.HTML'):
            context.path = file
            if '_AHB_' in file.stem:
                source = file.parent / file.name
                try:
                    extract_ahb(source)
                except BaseException as err:
                    logging.error(f'File {source}. Fehler {err}')

if __name__ == '__main__':
    logging.basicConfig(level=logging.ERROR)
    context = DataClassContext()
    main()

