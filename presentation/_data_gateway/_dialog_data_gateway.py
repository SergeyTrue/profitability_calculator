from pathlib import Path
from typing import List
import openpyxl

from ..input_data import Input
from ._i_data_gateway import IDataGateway, NoExcelFiles, correct_files_suffix


class DialogDataGateway(IDataGateway):

    def __init__(self, sender: str, subj: str, file_list: List, initial_mail: Path):
        self._sender = sender
        self._subj = subj
        self._file_list = file_list
        self._input_mail = initial_mail

    def process(self, valid_suffixes: List, *args, **kwargs):
        file_list = [Path(item) for item in self._file_list]
        file_list = (item for item in file_list if item.suffix.lower() in valid_suffixes)
        file_list = tuple(correct_files_suffix(file_list))
        if file_list is not None:

            excel_sheets = {}

            for excel in file_list:
                wb = openpyxl.load_workbook(excel)
                excel_sheets[excel] = wb.worksheets

        result = Input()
        result.sender = self._sender
        result.subject = self._subj
        result.file_list = file_list
        result.excel_sheets = excel_sheets
        result.initial_mail_path = self._input_mail
        return result
