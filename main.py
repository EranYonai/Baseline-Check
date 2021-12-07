from difflib import SequenceMatcher

import openpyxl
from PyQt5 import QtWidgets, uic

import config
import sys

# Global variables
name_to_color = {'red': '#FF0000', 'yellow': '#FFA500', 'green': '#008000'}
Alph_Zero_Val = ord('A') - 1


class MainWindow(QtWidgets.QMainWindow):
    """
    Main dialog of the application
    """

    def __init__(self):
        """
        init - initialize on dialog call.
        """
        super(MainWindow, self).__init__()
        self.excel_map = explore_excel()
        uic.loadUi(config.FILE_PATHS['MAIN_UI'], self)  # Loads MainWindow UI using pyqt5 uic.
        self.check_button.clicked.connect(self.scan_baseline)  # Check button click event

    def scan_baseline(self):
        """
        Gets a string of a baseline description from GUI and colors the lines according to item's presence DB
        """
        try:
            res = ''
            bl_text = self.textEdit.toPlainText()
            sheet_name = column_name = item_val = None

            for line in bl_text.split('\n'):
                if not line or line.startswith('-----'):  # New Section - reset variables
                    sheet_name = column_name = item_val = None
                elif sheet_name is not None:  # Item Category (sheet) was already found
                    column_name, item_val = line.split(':')
                    item_val = item_val.strip()
                    if sheet_name == 'System':  # Columns in System WS
                        if column_name.startswith('Aquarium'):
                            column_name = 'Aquarium Number'
                    elif sheet_name == 'WS':  # Columns in Workstation WS
                        pass
                    elif sheet_name == 'Catheters':  # Columns in Catheters WS
                        column_name = 'Catalog Number'
                    elif sheet_name == 'Pacers':
                        if column_name.endswith('Serial Number'):
                            column_name = 'Pacer Serial Number'
                # Find which sheet to search in
                elif line.startswith('System #'):
                    sheet_name = 'System'
                    column_name = item_val = None
                elif line.startswith('Workstation #'):
                    sheet_name = 'WS'
                    column_name = item_val = None
                elif line.startswith('Catheters') or line.startswith('Extenders'):
                    sheet_name = 'Catheters'
                    column_name = item_val = None
                elif line.startswith('Pacer') or line.startswith('Printer') or line.startswith('EPU'):
                    sheet_name = 'Pacers'
                    column_name = item_val = None

                # if Sheet, column and item_value have all been detected- Search for similarities in the specified column and sheet
                if sheet_name and column_name and sheet_name in self.excel_map and column_name in self.excel_map[
                    sheet_name]:
                    diff = find_word_in_db(sheet_name, self.excel_map[sheet_name][column_name], item_val)
                    color = 'green' if diff == 1 else 'yellow' if diff >= 0.8 else 'red'
                    res += f'{color_str(color, line)}<br>'  # Append colored line
                else:
                    res += f'{line}<br>'  # append unchecked lines as they are

            self.textEdit.setHtml(res)
        except Exception as e:
            print("Problem in scanning baselines: " + str(e))


def color_str(color: str, word: str) -> str:
    """
    Adds color tags to the given color.

    :param color: str = 'green', 'red, 'yellow'.
    :param word: str of the word you want tagged.
    :return: string: 'color_tag + word + close_tag'
    :rtype: str
    """
    return f'<span style="color:{name_to_color[color]};">{word}</span>'


def find_word_in_db(sheet_name, column, word):
    diff = 0
    try:
        wb = openpyxl.load_workbook(config.FILE_PATHS['EQP_EXCEL'])  # load excel workbook
        ws = wb[sheet_name]  # loads specific sheet, always system until smarter
        for row in range(3, ws.max_row):  # iterates all rows (row = tuple)
            if ws[f'{column}{row}'].value is not None:  # Iterate all rows of specified column
                diff = max(diff, similar(word, str(ws[f'{column}{row}'].value)))
                if diff == 1:  # Stop if perfect match was found
                    return diff
        return diff
    except Exception as e:
        print("Problem finding word in db: " + str(e))


def similar(str1, str2):
    """
    Gets two string and returns their correlation in percentage (float 0-1).
    Using SequenceMatcher of difflib library.

    :param str1: String 1
    :param str2: String 2
    :return: float value
    """
    return SequenceMatcher(None, str1, str2).ratio()


def explore_excel():
    """
    Maps the excel file specified in config.py
    :return: dictionary mapping the excel {<sheet_title>: {<column_title>: <column_letter>, ...}, ...}
    """
    try:
        res = dict()
        wb = openpyxl.load_workbook(config.FILE_PATHS['EQP_EXCEL'])  # load excel workbook
        for sheet in wb:
            res[sheet.title] = dict()
            for cell in sheet[2]:  # Iterate cells in 2nd row
                res[sheet.title][cell.value] = chr(cell.column + Alph_Zero_Val)  # get column letter
        return res
    except Exception as e:
        print("Problem exploring excel: " + str(e))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    win.setFocus()
    sys.exit(app.exec_())
