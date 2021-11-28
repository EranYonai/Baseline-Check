from PyQt5 import QtWidgets, uic, QtGui, QtCore, Qt
import config, sys
import openpyxl
from difflib import SequenceMatcher

class MainWindow(QtWidgets.QMainWindow):
    """
    Main dialog of the application
    """

    def __init__(self):
        """
        init - initialize on dialog call.
        """
        super(MainWindow, self).__init__()
        uic.loadUi(config.FILE_PATHS['MAIN_UI'], self)
        self.check_button.clicked.connect(self.inspect_text)

    def inspect_text(self):
        try:
            word = self.textEdit.toPlainText()
            diff = find_word_in_db(word)
            if diff == 'match':
                self.textEdit.setText(color_word('green', word))
            elif diff == 'partial_match':
                self.textEdit.setText(color_word('yellow', word))
            else:
                self.textEdit.setText(color_word('red', word))
        except Exception as e:
            print(e)



def color_word(color: str, word: str) -> str:
    """
    :param color: str = 'green', 'red, 'yellow'.
    :param word: str of the word you want tagged.
    :return: string: color_tag + word + close_tag
    :rtype: str
    """
    RED_TAG = "<span style=\" color:#ff0000;\" >"
    GREEN_TAG = "<span style=\" color:#008000;\" >"
    YELLOW_TAG = "<span style=\" color:#FFFF00;\" >"
    CLOSE_TAG = "</span>"
    if color == 'red':
        return RED_TAG + word + CLOSE_TAG
    if color == 'green':
        return GREEN_TAG + word + CLOSE_TAG
    if color == 'yellow':
        return YELLOW_TAG + word + CLOSE_TAG

def find_word_in_db(word):
    try:
        wb = openpyxl.load_workbook(config.FILE_PATHS['EQP_EXCEL'])  # load excel workbook
        sheet_system = wb[config.EXCEL_SHEET['system']]  # loads specific sheet, always system until smarter
        for row in sheet_system.iter_rows():  # iterates all rows (row = tuple)
            for cell in row:  # iterates cells in row
                if cell.value is not None:  # cell.value to get str value.
                    diff = similar(word, str(cell.value))
                    if  diff == 1:
                        return 'match'
                    elif diff > 0.8:
                        return 'partial_match'
        return 'no_match'
    except Exception as e:
        print(e)

def similar(str1, str2):
    """
    Gets two string and returns their correlation in precentage (0-1).
    Using SequenceMatcher of difflib library.
    :param str1: String 1
    :param str2: String 2
    :return: float value
    """
    return SequenceMatcher(None, str1, str2).ratio()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    win.setFocus()
    sys.exit(app.exec_())
