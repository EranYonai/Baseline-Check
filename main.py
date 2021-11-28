from PyQt5 import QtWidgets, uic, QtGui, QtCore, Qt
import config, sys
import openpyxl


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
        self.check_button.clicked.connect(self.find_word_in_db)

    def find_word_in_db(self):
        word = self.textEdit.toPlainText()
        try:
            wb = openpyxl.load_workbook(config.FILE_PATHS['EQP_EXCEL'])  # load excel workbook
            sheet_system = wb[config.EXCEL_SHEET['system']]  # loads specific sheet, always system until smarter
            for row in sheet_system.iter_rows():  # iterates all rows (row = tuple)
                for cell in row:  # iterates cells in row
                    if cell.value is not None:  # cell.value to get str value.
                        if word == str(cell.value):  # should actually return int match.
                            colored_word = color_word(color='green', word=word)
                            self.textEdit.append(colored_word)

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


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    win.setFocus()
    sys.exit(app.exec_())
