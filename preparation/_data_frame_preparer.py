from typing import Mapping, Iterable
from tabulate import tabulate

import pandas as pd
import numpy as np

from mc.presentation import Input
from ._sheet_filter import filter_sheets as _filter_sheets


class DataFramePreparer:
    @staticmethod
    def process(input_data: Input, sheet_names_to_be_excluded: Iterable):
        result = PreparationResult()

        for excel, sheets in input_data.excel_sheets.items():
            filtered_sheets = _filter_sheets(sheets, sheet_names_to_be_excluded)
            df = _to_data_frame(excel, filtered_sheets)
            df = _concatenate_data_frames(df)
            df.reset_index(drop=True, inplace=True)
            df, bad_999999_price = _correct_and_return_999999price(df)
            df = _remove_thousand_separators(df)
            if bad_999999_price is not None:
                result.warning[excel] = bad_999999_price
            errors = _find_errors(df)
            if errors:
                result.wrong[excel] = errors
                continue

            result.accepted[excel] = df

        result.display_not_accepted()
        return result


class PreparationResult:
    def __init__(self):
        self.accepted = {}
        self.wrong = {}
        self.warning = {}

    def display_not_accepted(self):
        for item in self.wrong, self.warning:
            for file, file_wrongs in item.items():
                print(file)
                if isinstance(file_wrongs, Mapping):
                    for k, v in file_wrongs.items():
                        print(k)
                        print(tabulate(v, headers='keys', tablefmt='psql'), end='\n\n')
                else:
                    print(tabulate(file_wrongs, headers='keys', tablefmt='psql'), end='\n\n')


_MIN_ARTICLE = 3000


def _to_data_frame(excel, sheets):
    dfs = []
    for sheet in sheets:
        #TODO у pd.read_excel в коде на прошлом компе было ещё skip_blank_lines=False
        df = pd.read_excel(excel, sheet_name=sheet.title, header=None, usecols='A:F')
        df = df.rename(columns={
            0: 'article', 1: 'Description', 2: 'Quantity', 3: 'Price_spec', 4: 'Price_spec x Qty',
            5: 'Equipment type'
        })

        if 'Price_spec' not in df.columns:
            df['Price_spec'] = np.NaN
        if 'Price_spec x Qty' not in df.columns:
            df["Price_spec x Qty"] = np.NaN
        df = df.astype({'article': str})
        df['article'] = df['article'].map(lambda x: x.strip())
        df = df[pd.to_numeric(df['article'], errors='coerce').notnull()]
        df = df.astype({'article': float})

        df = df[df.article > _MIN_ARTICLE]
        df['article'] = df['article'].values.astype('int64')
        df.drop(df.columns[5:], axis=1, inplace=True)
        df['Equipment type'] = sheet.title

        dfs.append(df)

    return dfs


def _concatenate_data_frames(excel_data_frames):
    concatenated_df = pd.DataFrame()
    for df in excel_data_frames:
        concatenated_df = pd.concat([concatenated_df, df], sort=False)
    return concatenated_df


def _correct_and_return_999999price(df):
    bad_999999_price_df = df.loc[df['Price_spec'].astype(str).str[:7] == '999,999']
    if bad_999999_price_df.empty:
        bad_999999_price_df = None
    df.loc[df['Price_spec'].astype(str).str[:7] == '999,999', ['Price_spec', 'Price_spec x Qty']] = np.nan
    return df, bad_999999_price_df


def _find_errors(df):
    errors = {}
    non_numeric_qty_df = df.loc[~df['Quantity'].apply(np.isreal)]
    non_numeric_price_df = df.loc[~df['Price_spec'].apply(np.isreal)]
    if not non_numeric_qty_df.empty:
        errors['wrong_quantity'] = non_numeric_qty_df
    if not non_numeric_price_df.empty:
        errors['wrong_price'] = non_numeric_price_df

    cond = df['Quantity'].isnull()
    rows = df.loc[cond, :]
    if not rows.empty:
        errors['empty_qty'] = rows
    return errors


def _dot_remover(x):  # TODO: съедает десятичную часть
    try:
        float(x)
    except ValueError:
        if ',' in x:
            x = x.replace(',', '')
            try:
                float(x)
            except ValueError:
                pass
    try:
        return float(x)
    except ValueError:
        return x


def _remove_thousand_separators(df):
    df['Price_spec'] = df['Price_spec'].apply(lambda x: _dot_remover(x))
    df['Price_spec x Qty'] = df['Price_spec x Qty'].apply(lambda x: _dot_remover(x))
    return df
