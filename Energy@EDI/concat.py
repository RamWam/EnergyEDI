from pathlib import Path
import pandas as pd

def write_concat_file(sheet, concat_files, path, doc_type=''):
    # Tabellen zusammenführen
    result = pd.concat(concat_files)
    # Datum und File als Index setzen, damit sie in die ersten beiden Spalten geschrieben werden.
    # Gleichzeitig den alten, bei der Erstellung der ursprünglichen Excel Dateien übernommenen
    # Index löschen
    result = result.set_index(['datum', 'file'])
    if doc_type != '':
        doc_type = doc_type +  '_'
    new_file = path / 'Zusammenführung' / f'{doc_type}{sheet}.xlsx'
    result.to_excel(new_file, sheet_name=sheet, merge_cells=False)

def main():
    change_hists = []
    path = Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY')

    # Datenstruktur mit den Dokumententype MIG/AHB/EBD und den für diese in den
    # Excel Extrakten enthaltenen Sheets
    concat_dict = {'MIG': {'Segmentlayout': [],
                           'Nachrichtenstruktur': [],
                           'Segmentlayout_Qual': []},
                   'AHB': {'Anwendungsfälle': [],
                           'Bedingungen': []},
                   'EBD': {'EBD': [],
                           'CodeLists': []}}

    # Alle Excel Tabellen in den Extraktions Verzeichnissen
    for file in path.rglob('_Extractions/*.xlsx'):
        # Je noch Dokumentenart verschiedenen Sheet zusammenführen
        for doc_type, file_dict in concat_dict.items():
            if doc_type in file.stem:
                for sheet, concat_files in file_dict.items():
                    # Sheet lesen
                    df = pd.read_excel(file, sheet_name=sheet)
                    # Dateiname und Datum aus übergeordneten Ordner als Spalten hinzufügen
                    df['file'] = file.stem
                    df['datum'] = file.parent.parent.stem
                    concat_files.append(df)

        # Änderungshistorie lesen
        df = pd.read_excel(file, sheet_name='Änderungshistorie')
        # Dateiname und Datum aus übergeordneten Ordner als Spalten hinzufügen
        df['file'] = file.stem
        df['datum'] = file.parent.parent.stem
        change_hists.append(df)

    # Die Änderungshistorien liegen in diversen Dokumentenarten vor (AHB,MIG, EBD ...) vor. Sie werden hier in eine
    # Excel zusammengeführt.
    write_concat_file('Änderungshistorie', change_hists, path)

    for doc_type, file_dict in concat_dict.items():
        for sheet, concat_files in file_dict.items():
            write_concat_file(sheet, concat_files, path, doc_type=doc_type)

if __name__ == '__main__':
    main()