from pathlib import Path
from typing import Dict, Iterable, Optional, List

from openpyxl.worksheet.worksheet import Worksheet


ExcelFile = Path
Sheets = Iterable[Worksheet]


class Input:
    def __init__(self):
        self.initial_mail_path: Optional[Path] = None
        self.sender_path: Optional[Path] = None
        self.subj_path: Optional[Path] = None
        self.excel_sheets: Dict[ExcelFile, Sheets] = {}
        self.file_list_path: Optional[Path] = None
        self.file_list: Optional[List[Path]] = None
