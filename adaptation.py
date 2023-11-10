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
        wks.clear(start='a9', end=None)
        wks.set_dataframe(self.table, begin_cell, copy_index=False, copy_head=False, extend=True, fit=False, escape_formulae=True)


    def load_by_day(self, columns):
        pass


if __name__ == "__main__":
    pass