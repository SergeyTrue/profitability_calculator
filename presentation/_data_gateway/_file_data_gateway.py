from pathlib import Path
from typing import Iterable

import openpyxl

from ..input_data import Input
from ._i_data_gateway import IDataGateway, NoExcelFiles, correct_files_suffix


class OutlookDataGateway(IDataGateway):

    def __init__(self, sender_path: Path, subj_path: Path, file_list_path: Path, initial_mail_path: Path):
        self._sender_path = sender_path
        self._subj_path = subj_path
        self._file_list_path = file_list_path
        self._initial_mail = initial_mail_path

    def process(self, valid_suffixes: Iterable):
        with self._file_list_path.open() as stream:
            file_list = list(stream.read().strip().split('|'))
            file_list = [Path(item) for item in file_list]
            file_list = (item for item in file_list if item.suffix.lower() in valid_suffixes)
            file_list = tuple(correct_files_suffix(file_list))
            excel_sheets = {}
            if len(file_list) is not 0:

                for excel in file_list:
                    wb = openpyxl.load_workbook(excel)
                    excel_sheets[excel] = wb.worksheets

            result = Input()
            result.initial_mail_path = self._initial_mail
            result.sender_path = self._sender_path
            result.subj_path = self._subj_path
            result.file_list = file_list
            result.excel_sheets = excel_sheets
            return result
