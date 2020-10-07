from pathlib import Path

import pandas as pd
import openpyxl

from . import IOutput
from mc.common import Role, config, NoExcelFilesResult


class InputFileModifierOutput(IOutput):
    def __init__(self):
        self.file_list_to_save = []

    def process(self, messages):
        boss_message = (item for item in messages if item.receiver.role == Role.BOSS)
        boss_message = next(boss_message, None)
        if isinstance(boss_message.data, NoExcelFilesResult):
            return
        if boss_message and boss_message.data.accepted_user_result is not None:
            for item in boss_message.data.accepted_user_result.children:
                if item.bad_articles is None:
                    self._delete_last_run_sheets(item.path)
                    self._add_result_sheets(item)
                    self.file_list_to_save.append(item.path)

    @staticmethod
    def _add_result_sheets(line_item):
        file = line_item.path

        with pd.ExcelWriter(str(file), engine='openpyxl', mode='a') as writer:
            df = line_item.discount_margin_relation_merged_df
            df.to_excel(writer, sheet_name=config.result_sheet_names[0], index=False)

            df = line_item.complete_price_cost
            df.to_excel(writer, sheet_name=config.result_sheet_names[1], index=False)

    @staticmethod
    def _delete_last_run_sheets(file: Path):
        excel = openpyxl.load_workbook(str(file))
        for sheet_name in excel.sheetnames:
            if sheet_name in config.result_sheet_names:
                del excel[sheet_name]
        excel.save(str(file))
