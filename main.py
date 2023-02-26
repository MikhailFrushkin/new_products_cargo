from pathlib import Path

import numpy as np
import pandas as pd
from loguru import logger


def read_file(self, file_name_goods, file_name_min, file_name_vsl):
    # colum_goods = ['Ссылка', 'Номер', 'Перенос под график', 'Груз', 'Номер SO',
    #                'Код номенклатуры', 'Наименование номенклатуры', 'Номер партии',
    #                'Кол-во в графике', 'Отгруженное количество', 'Полученное количество',
    #                'Доступное количество', 'Доступное количество с графиком', 'Объем',
    #                'Вес нетто', 'Вес брутто', 'Количество паллет']
    # colum_min = ['SG', 'NG', 'Признак Новинки', 'name', 'Номенклатура',
    #              'Значения в базе', 'Новые значения']
    # colum_vsl = ['БЮ', 'Склад', 'Местоположение', 'Код \nноменклатуры',
    #              'Краткое наименование', 'Описание товара', 'Reason code', 'ТГ', 'НГ',
    #              'Поставщик', 'Наименование', 'Физические \nзапасы', 'Продано',
    #              'Зарезерви\nровано', 'Доступно']

    df_goods = pd.read_excel(file_name_goods,
                             usecols=['Код номенклатуры', 'Наименование номенклатуры', 'Отгруженное количество',  # noqa
                                      'Объем', 'Вес брутто'])  # noqa

    df_min = pd.read_excel(file_name_min, usecols=['SG', 'Номенклатура'])  # noqa

    df_vsl = pd.read_excel(file_name_vsl,
                           usecols=['Местоположение', 'Код \nноменклатуры', 'Доступно'])  # noqa
    df_vls = df_vsl[(df_vsl['Местоположение'] == 'V-Sales_825') & (df_vsl['Доступно'] > 0)]
    result_df = pd.merge(df_goods, df_min,
                         left_on=['Код номенклатуры'],
                         right_on=['Номенклатура'], how='left')
    result_df = pd.merge(result_df, df_vls,
                         left_on=['Код номенклатуры'],
                         right_on=['Код \nноменклатуры'], how='left')
    result_df.drop(['Номенклатура', 'Местоположение', 'Код \nноменклатуры'], axis=1, inplace=True)
    tg = list(result_df['SG'].unique())
    colum_dict = {'SG': 'first',
                  'Отгруженное количество': sum,
                  'Объем': sum,
                  'Вес брутто': sum,
                  }
    df_main = result_df.groupby(['SG'], as_index=False).agg(colum_dict)

    with pd.ExcelWriter('Приход кроссдок 19.02.xlsx', engine='xlsxwriter') as writer:
        df_sort = df_main.sort_values(by='Объем', ascending=False)
        df_colors = df_sort.style.set_properties(
            **{
                "text-align": "left",
                "font-weight": "bold",
                "font-size": "14px",
                "border": "1px solid black"
            })
        df_colors.to_excel(writer, sheet_name='Общее', index=False, na_rep='')
        worksheet = writer.sheets['Общее']
        set_column_main(df_main, worksheet)

        for group in tg:
            df_temp = result_df[((result_df['SG'] == group) & (result_df['Доступно'].isnull()))]
            if len(df_temp) > 0:
                name = group.split(' ')[0]
                df_temp.to_excel(writer, sheet_name=f'Новинки ТГ.{name}', header=True, index=False, na_rep='')
                worksheet = writer.sheets[f'Новинки ТГ.{name}']
                set_column(df_temp, worksheet)


def set_column(df, worksheet):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 80)
    worksheet.set_column('C:G', 30)


def set_column_main(df, worksheet):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:D', 30)


class Test:
    def __init__(self):
        self.current_dir = Path.cwd()


if __name__ == '__main__':
    test = Test()
    read_file(test, r'C:\Users\receipt of goods\следующий кросс.xlsx', r'C:\Users\receipt of goods\min.xlsx',
              r'C:\Users\receipt of goods\vsl.xlsx')
