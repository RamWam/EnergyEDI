import subprocess
from pathlib import Path
import shutil
import time
import os

# work_dir = Path('/EDI@ENERGY/EDI_ENERGY/temp')
work_dir = Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/temp')

def run_word(source, target_dir):
    # in dem Working Directory wird jede PDF Datei nach temp.kopiert und dann
    # ein mit APPLE Script erzeugt APP aufgerufen, die ein Word-Macro aufruft.
    # Das Word Macro offnet die temp.pdf Datei und speichert sie unter temp.html
    # wieder ab. Danach wird temp.html

    DUMMY_TEXT = 'ist leer'
    temp_docx = work_dir / 'temp.docx'
    temp_html = work_dir / 'temp.html'

    shutil.copy(source, temp_docx)

    with open(temp_html, 'w') as f:
        data = DUMMY_TEXT
        f.write(data)

    subprocess.run(["open", "../docx_to_html.app"])

    # Subprocess wird beim Aufruf der APP beendet, wenn der Launcher fertig ist und nicht wenn
    # Word beendet ist.
    # Deshalb wird hier geprüft, ob Word die Datei geschrieben hat.
    # temp.html muss aber schon vor Aufruf von Word vorhanden sein, weil es
    # sonst ein Berechtigungs Problem gibt. Dieses tritt nur auf, wenn das
    # Word Macro über die Apple-Script APP gestartet wird.
    file_filled = False
    for _ in range(120):
        try:
            with open(temp_html, 'r') as f:
                data = f.read()
            if data == DUMMY_TEXT:
                time.sleep(1)
            else:
                file_filled = True
                break
        except UnicodeDecodeError:
            file_filled = True
            break

    if file_filled == False:
        raise Exception(f"PDF {source.name} konnte nicht in HTML konvertiert werden.")

    time.sleep(3)
    target_file = source.stem + '.HTML'
    target = target_dir / target_file

    os.makedirs(os.path.dirname(target), exist_ok=True)
    shutil.copy(temp_html, target)
    temp_html.unlink()
    temp_docx.unlink()

def main():
    paths = list()
    #paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20221001'))
    #paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20230401'))
    paths.append(Path('/Users/rainerwappler/PycharmProjects/pythonProject/EDI@ENERGY/EDI_ENERGY/20231001'))

    for path in paths:
        target = path / '_Extractions'
        file_list = path.rglob('*.docx')
        for file in file_list:
            source = file.parent / file.name
            run_word(source, target)

if __name__ == "__main__":
    main()