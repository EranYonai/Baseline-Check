from PyQt5 import QtWidgets, uic, QtGui, QtCore, Qt
import config, sys

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

    def find_word_in_db(self, word):
        pass

    def color_word(self):
        pass
    

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    win.setFocus()
    sys.exit(app.exec_())