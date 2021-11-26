# Config file
import os


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = os.path.expanduser("~/PycharmProjects/OracleWeek")
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


FILE_PATHS = {'MAIN_UI': resource_path('dependencies/MainWindow.ui'),
              }