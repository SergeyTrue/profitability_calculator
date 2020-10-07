import numpy as np
import pandas as pd
from tabulate import tabulate

from .common import SuccessUserResult, ResultLineItem, config


_BAD_COST = [0.01, 0, np.nan]


class Processor:
    @staticmethod
    def process(excel_data_frames, sender_company, multiplicity_df, minimum_quantity_df, warning_999999_price):
        data = dict(excel_data_frames)
        price_cost_df = _create_price_cost_df(sender_company)
        data = _merge_data_frames(data, price_cost_df)
        bad_cost_ratio_dict = _calculate_bad_cost_ratio(data)
        bad_cost_dict = _make_bad_cost(data, sender_company)
        data = _take_missing_prices_from_spec(data)
        data = _fill_empty_cost_with_spec_price(data)
        data = _fill_bad_cost_with_spec_price(data)
        data = _add_columns_price_cost_x_qty(data)

        result = SuccessUserResult()

        for excel, df in data.items():
            bad_articles = _get_bad_articles(df)
            if bad_articles is not None:
                child = ResultLineItem(
                    path=excel,
                    discount_margin_relation_merged_df=None,
                    discount_margin_relation_grouped_df=None,
                    price_by_type_of_equipment=None,
                    equal_margin_df=None,
                    bad_cost=None,
                    full_emulator=None,
                    complete_price_cost=None,
                    bad_articles=bad_articles,
                    minimum_quantity_warning=
                    _create_minimum_quantity_warning(df, minimum_quantity_df),
                    multiplicity_warning=_create_multiplicity_warning(df, multiplicity_df),
                    bad_cost_ratio=bad_cost_ratio_dict[excel],
                    warning_999999_price=warning_999999_price.get(excel)
                )
            else:
                child = ResultLineItem(
                    path=excel,
                    discount_margin_relation_merged_df=_make_merged_discount_margin_dfs(df),
                    discount_margin_relation_grouped_df=_make_grouped_discount_margin_dfs(df),
                    price_by_type_of_equipment=_make_price_by_type_of_equipment_df(df),
                    equal_margin_df=_make_equal_margin_df(df),
                    bad_cost=bad_cost_dict.get(excel),
                    full_emulator=_make_full_emulator_df(df, sender_company),
                    complete_price_cost=df,
                    bad_articles=None,
                    minimum_quantity_warning=
                    _create_minimum_quantity_warning(df, minimum_quantity_df),
                    multiplicity_warning=_create_multiplicity_warning(df, multiplicity_df),
                    bad_cost_ratio=bad_cost_ratio_dict[excel],
                    warning_999999_price=warning_999999_price.get(excel)
                )
            result.children.append(child)

        return result


def _create_price_cost_df(company):

    price_cost_df = pd.read_excel(config.PRICE_COST_FILE, sheet_name='PL', usecols='A:G')
    price_cost_df.rename(columns={
        config.PRICE_COST_COLUMN[company][0]: 'Price_list',
        config.PRICE_COST_COLUMN[company][1]: 'Cost'
    }, inplace=True)
    price_cost_df = price_cost_df[['article', 'Name', 'Price_list', 'Cost']]
    return price_cost_df


def _create_minimum_quantity_warning(df, moq_df):
    qty_comparison_df = df.merge(moq_df, on='article', how='left')
    qty_comparison_df = qty_comparison_df.loc[:, ['article', 'Name_x', 'Quantity_x', 'Quantity_y', 'Equipment type']]
    qty_comparison_df.rename(columns={'Name_x': 'Name', 'Quantity_x': 'Spec_quantity', 'Quantity_y': 'Min_quantity'
                                      }, inplace=True)
    qty_comparison_df.dropna(axis='index', how='any', subset=['Min_quantity'])
    qty_comparison_df = qty_comparison_df.loc[qty_comparison_df['Spec_quantity'] < qty_comparison_df['Min_quantity']]

    return qty_comparison_df

# def create_multiplicity_warning(df, multiplicity_df):
#
#     qty_comparison_df = df.merge(multiplicity_df, on='article', how='left')
#     qty_comparison_df = qty_comparison_df.loc[:, ['article', 'Name_x', 'Quantity_x', 'Quantity_y', 'Equipment type']]
#     qty_comparison_df.rename(columns={'Name_x': 'Name', 'Quantity_x': 'Spec_quantity', 'Quantity_y': 'Min_quantity'},
#                              inplace=True)
#     qty_comparison_df.dropna(axis='index', how='any', subset=['Min_quantity'])
#     qty_comparison_df = qty_comparison_df.loc[qty_comparison_df['Spec_quantity'] % qty_comparison_df['Min_quantity']!=0]
#
#     return qty_comparison_df


def _create_multiplicity_warning(df, multiplicity_df):

    df = df.merge(multiplicity_df, on='article', how='left')
    df = df.loc[:, ['article', 'Name_x', 'Quantity_x', 'Quantity_y', 'Equipment type']]
    df.rename(columns={'Name_x': 'Name', 'Quantity_x': 'Spec_quantity', 'Quantity_y': 'Mult_quantity'}, inplace=True)
    df['remainder'] = df['Spec_quantity'].mod(df['Mult_quantity'])
    df = df.dropna(subset=['Mult_quantity'])
    df = df.loc[df['remainder'] != 0]

    return df


def _merge_data_frames(excel_data_frames, price_cost_df, ):

    for excel, df in excel_data_frames.items():
        df = df.merge(price_cost_df,  on='article', how='left')
        excel_data_frames[excel] = df

    return excel_data_frames


def _calculate_bad_cost_ratio(excel_data_frames):
    bad_cost_ratio_dict = {}
    for _, df in excel_data_frames.items():
        df_bad_cost = df.loc[df['Cost'].isin(_BAD_COST)]
        bad_cost_sum = sum(df_bad_cost.Price_spec)
        total_sum = sum(df.Price_spec)
        ratio = bad_cost_sum / total_sum
        bad_cost_ratio_dict[_] = ratio
    return bad_cost_ratio_dict


def _take_missing_prices_from_spec(excel_data_frames):
    for _, df in excel_data_frames.items():
        df.Price_list.fillna(df.Price_spec, inplace=True)

    return excel_data_frames


def _fill_empty_cost_with_spec_price(excel_data_frames):
    for excel, df in excel_data_frames.items():
        df['Cost'].fillna(df['Price_spec'], inplace=True)
        excel_data_frames[excel] = df
    return excel_data_frames


def _make_bad_cost(excel_data_frames, company):
    bad_costs = {}
    for excel, df in excel_data_frames.items():

        df_bad_cost = df.loc[df['Cost'].isin(_BAD_COST)]
        emulator_df = pd.DataFrame(columns=['company', 'article', '1', 'article2', '01'])
        emulator_df['article'] = df_bad_cost['article'].copy()
        emulator_df['article2'] = df_bad_cost['article'].copy()
        emulator_df['company'] = company
        emulator_df['1'] = 1
        emulator_df['01'] = '01'
        emulator_df.drop_duplicates(inplace=True)
        print(tabulate(emulator_df.head(), headers='keys', tablefmt='psql'))
        bad_costs[excel] = emulator_df
    return bad_costs


def _get_bad_articles(df):
    df_bad_articles = df[(df['Price_spec'].isnull() & df['Cost'].isnull() & df['Price_list'].isnull())]
    if not df_bad_articles.empty:
        return df_bad_articles
    else:
        return None


def _fill_bad_cost_with_spec_price(excel_data_frames):

    for excel, df in excel_data_frames.items():
        df['Cost'] = np.where(df['Cost'].isin(_BAD_COST), df['Price_spec'], df['Cost'])
        excel_data_frames[excel] = df
    return excel_data_frames


def _add_columns_price_cost_x_qty(excel_data_frames):

    for excel, df in excel_data_frames.items():
        df['Price_list_x_Qty'] = df.Quantity * df['Price_list']
        df['Cost_x_Qty'] = df.Quantity * df['Cost']
    return excel_data_frames


def _create_discount_margin_df(price, cost):

    discount_margin_df = pd.DataFrame(columns=['Price', 'Discount', 'Sale', 'Margin', 'Cost'], index=range(50))
    discount_margin_df.Price = price
    discount_margin_df.Cost = cost
    discount_margin_df.Discount = np.arange(0, 1, 0.02)
    discount_margin_df.Sale = discount_margin_df.Price * (1 - discount_margin_df.Discount)
    discount_margin_df.Margin = (discount_margin_df.Sale - discount_margin_df.Cost) / discount_margin_df.Sale
    discount_margin_df = discount_margin_df[discount_margin_df.Margin.between(0, 0.80)]
    format_mapping = {'Price': '€ {:,.2f}', 'Discount': '{:.0%}', 'Sale': '€ {:,.2f}', 'Margin': '{:.0%}',
                      'Cost': '€ {:,.2f}'}
    for key, value in format_mapping.items():
        discount_margin_df[key] = discount_margin_df[key].apply(value.format)

    for column in ['Price', 'Sale', 'Cost']:
        discount_margin_df[column] = discount_margin_df[column].astype(str).str.replace(',', ' ').str.replace('.', ',')

    return discount_margin_df


def _make_merged_discount_margin_dfs(df):

    total_price = sum(df.Price_list_x_Qty)
    total_cost = sum(df.Cost_x_Qty)
    discount_margin_df = _create_discount_margin_df(total_price, total_cost)
    #print(tabulate(discount_margin_df.head(), headers='keys', tablefmt='psql'))

    return discount_margin_df


def _make_grouped_discount_margin_dfs(df):
    equipment_dfs = {}

    g = df.groupby(['Equipment type'])
    if len(list(g.groups.keys())) == 1:
        pass
    else:
        for name, group in df.groupby(["Equipment type"]):
            total_price = sum(group.Price_list_x_Qty)
            total_cost = sum(group.Cost_x_Qty)
            group_discount_margin_df = _create_discount_margin_df(total_price, total_cost)
            equipment_dfs[name] = group_discount_margin_df

    return equipment_dfs


def _make_price_by_type_of_equipment_df(df):

    df_body = df.groupby('Equipment type')['Price_list_x_Qty'].sum()
    df_last_line = pd.Series(df['Price_list_x_Qty'].sum(), index=['Итого:'])
    df_complete = pd.concat([df_body, df_last_line], axis=0)
    df_complete = df_complete.to_frame()
    df_complete.reset_index(level=0, inplace=True)
    df_complete.columns = ['Тип оборудования', 'Стоимость']
    format_mapping = {'Стоимость': '€ {:,.2f}'}
    for key, value in format_mapping.items():
        df_complete[key] = df_complete[key].apply(value.format)
    df_complete['Стоимость'] = df_complete['Стоимость'].astype(str).str.replace(',', ' ').str.replace('.', ',')

    return df_complete


def _make_equal_margin_df(df):

    dfg = df.groupby('Equipment type')
    if dfg.ngroups > 1:
        equal_cost_df = pd.DataFrame(df.groupby('Equipment type')['Cost_x_Qty'].sum())
        equal_margin_df = pd.DataFrame(df.groupby('Equipment type')['Price_list_x_Qty'].sum())
        equal_margin_df = equal_margin_df.join(equal_cost_df)
        equal_margin_df.reset_index(level=0, inplace=True)
        equal_margin_df.round(2)
        equal_margin_df.columns = ['Тип оборудования', 'Стоимость по прайсу', 'Себестоимость']
        equal_margin_df.loc["Итого"] = equal_margin_df.sum(numeric_only=True)
        equal_margin_df.insert(loc=2, column='Скидка',
                               value=[r'=1-indirect("R[0]C[1]",0)/indirect("R[0]C[-1]",0)'
                                    for i in range(equal_margin_df.shape[0])])
        equal_margin_df.insert(loc=3, column='Цена продажи',
                                value=[r'=indirect("R[0]C[2]",0)/(1-indirect("R[0]C[1]",0))'
                                    for i in range(equal_margin_df.shape[0])])
        equal_margin_df.insert(loc=4, column='Маржа', value=['=INDIRECT("R2C5",0)'
                                    for i in range(equal_margin_df.shape[0])])
        height = str(len(equal_margin_df) - 1)
        equal_margin_df.iloc[-1, 0] = "Итого:"
        equal_margin_df.iloc[-1, [1, 3, 5]] = '=SUM(INDIRECT("R[-' + height + ']C:R[-1]C",0))'
        #print(tabulate(equal_margin_df, headers='keys', tablefmt='psql'))
        return equal_margin_df
    else:
        return None


def _make_full_emulator_df(df, company):

    full_emulator_df = pd.DataFrame(columns=['company', 'article', 'Quantity', 'article2', '01'])
    full_emulator_df['article'] = df['article'].copy()
    full_emulator_df['article2'] = df['article'].copy()
    full_emulator_df['company'] = company
    full_emulator_df['Quantity'] = df['Quantity']
    full_emulator_df.Quantity = full_emulator_df.Quantity.astype('int32')
    full_emulator_df['01'] = '01'
    #print(tabulate(full_emulator_df, headers='keys', tablefmt='psql'))
    return full_emulator_df

