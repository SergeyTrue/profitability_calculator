import sys
import argparse
from unittest.mock import MagicMock

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    files = QFileDialog.getOpenFileNames(
        None,
        "Select one or more files to open",
        r"C:\Users\belose\Downloads",
        "Excels (*.xlsx)")
    files, _ = files

    try:
        past = argparse.ArgumentParser
        _ = MagicMock()
        argparse.ArgumentParser = lambda: _

        class Data:
            def __init__(self, **kwargs):
                self.__dict__.update(kwargs)

        _.parse_args.return_value = Data(**{'data_receipt_way': 'dialog',
                                            'sender': r'Sergey.Belousov@delaval.com',
                                            'subj': r'test',
                                            'file_list': files,
                                            'initial_mail': None})
        import mc.__main__
    finally:
        argparse.ArgumentParser = past
