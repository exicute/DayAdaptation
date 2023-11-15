import pandas as pd
import numpy as np
import os
import requests
from datetime import datetime
import dateutil.relativedelta
import math
import pygsheets



class ExcelReview():

    def __init__(self, table):
        self.table = table
        self.df = pd.DataFrame()


    def get_table(self):
        report_rows_count = requests.post("http://dc0-prod-bi-external-01.esoft.local:10022/report/api/v1/requestReportRowsCount", json=self.table)

        iterations_count = math.ceil(report_rows_count.json() / self.table['rows'])

        result = pd.DataFrame()
        for i in range(iterations_count):
            start = i * self.table['rows']
            self.table['start'] = start
            report_data_iteration = requests.post("http://dc0-prod-bi-external-01.esoft.local:10022/report/api/v1/requestRawReportData",
                                                json=self.table)
            temp_json_table = list(report_data_iteration.json().values())
            result = pd.concat([result, pd.DataFrame.from_dict(temp_json_table)])

        return result


    def load_full_table(self, xlsx_key, xlsx_sheet, begin_cell):
        gc = pygsheets.authorize(service_file=r'C:\Users\ws-tmn-an-15\Desktop\Харайкин М.А\Python документы\python-automation-script-jupyter-notebook-266007-21fda3e2971a.json')
        sh = gc.open_by_key(xlsx_key)
        wks = sh.worksheet_by_title(xlsx_sheet)
        wks.clear(start=begin_cell, end=None)
        wks.set_dataframe(self.df, begin_cell, copy_index=False, extend=True, fit=False, escape_formulae=True)


    def load_by_day(self, xlsx_key, xlsx_sheet, merge_column, other_columns):
        gc = pygsheets.authorize(service_file=r'C:\Users\ws-tmn-an-15\Desktop\Харайкин М.А\Python документы\python-automation-script-jupyter-notebook-266007-21fda3e2971a.json')
        sh = gc.open_by_key(xlsx_key)
        wks = sh.worksheet_by_title(xlsx_sheet)
        names_country = wks.get_as_df(start='a1')

        def make_sort_table(xlsx_table, merge_column, other_columns):
            loaded_table = xlsx_table.merge(self.df, how='outer', on=merge_column)
            loaded_table = loaded_table.fillna(0)

            other_columns_df = pd.DataFrame()
            for col in other_columns:
                other_columns_df[col] = loaded_table[f'{col}_x']
                other_columns_df.loc[other_columns_df[col]==0, col] = loaded_table.loc[other_columns_df.loc[other_columns_df[col]==0].index.values,
                                                                                   f'{col}_y']

                loaded_table = loaded_table.drop(columns={f'{col}_y'})
                loaded_table = loaded_table.rename(columns={f'{col}_x':f'{col}'})
                loaded_table[col] = other_columns_df[col]

            #используя суффиксы _x, _y объединить другие столбцы в один и вставить в начало таблицы
            return loaded_table

        loaded_table = make_sort_table(names_country, merge_column, other_columns)
        wks.clear(start='a1', end=None)
        wks.set_dataframe(loaded_table, start='a1', copy_index=False, extend=True, fit=False, escape_formulae=True)


    def load_month_table(self):
        pass

        

if __name__ == "__main__":
    pass