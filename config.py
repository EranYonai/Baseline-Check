# Config file
import os, sys


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# File paths is dumb right now
FILE_PATHS = {'MAIN_UI': resource_path('dependencies\\Baseline-check.ui'),
              'EQP_EXCEL': resource_path('bin\\Equipment Traceability.xlsx')
              }


EXCEL_SHEET = {'system': 'System',
               'ws': 'WS',
               'monitor': 'Monitor'}
