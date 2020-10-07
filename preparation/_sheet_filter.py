from typing import Iterable, AnyStr

from openpyxl.worksheet.worksheet import Worksheet


SheetNamePart = AnyStr


def filter_sheets(sheets: Iterable[Worksheet], exclude_sheets_name_similar_to: Iterable[SheetNamePart]):
    not_empty_sheets = (sheet for sheet in sheets if sheet.max_column != 1 and sheet.max_row != 1)
    filtered_sheets = (sheet for sheet in not_empty_sheets
                       if not _is_string_contains_any_substring(sheet.title.lower(), exclude_sheets_name_similar_to))

    return tuple(filtered_sheets)


def _is_string_contains_any_substring(string, exclude_sheets_name_similar_to):
    checks = (substring in string for substring in exclude_sheets_name_similar_to)
    return any(checks)


