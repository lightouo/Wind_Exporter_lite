from WindPy import w
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
try:
    from typing import Literal
except ImportError:
    from typing_extensions import Literal

class Wind_Exporter:
    """
    Wind_Exporter
    =====
    code: list
        股票代码, 例如['000001.SZ', '000002.SZ']
    indicator: str
        指标, 例如'close'
    method: str (default: 'wsd')
        wsd: 单日期 
        wss: 日期截面数据
    Date_List: list (default: None)
        日期列表, 例如['2020-01-01', '2020-01-02']
    StartDate: str (default: None)
        起始日期, 例如'2020-01-01'
    EndDate: str (default: None)
        结束日期, 例如'2020-01-02'
    options: str (default: None)
        选项, 遵循Wind API的options参数
    """
    def __init__(self, code, indicator:str, method:Literal['wsd', 'wss']='wsd', Date_List:list=None, StartDate=None, EndDate=None, options=None):
        self.method = method
        self.code = code
        self.EndDate = EndDate if EndDate else datetime.today().strftime('%Y-%m-%d')
        self.Date_List = Date_List
        self.StartDate = EndDate if StartDate is None else StartDate
        self.indicator = indicator
        self.options = options
        self.data = []

    def __len__(self):
        return len(self.data)

    @staticmethod
    def check_connection(func):
        if not w.isconnected():
            w.start()
        return func

    @staticmethod
    def create_date_col(df, date):
        df['date'] = date
        column = np.roll(np.arange(len(df.columns)), 1)
        column[0], column[1] =column[1], column[0]
        return df.iloc[:, column]

    @check_connection
    def get_data(self, output=Literal['df', 'excel', None], round_=None):
        if self.method == 'wsd':
            self.get_data_wsd()
        elif self.method == 'wss':
            self.get_data_wss()
        else:
            raise ValueError('method must be wsd or wss')


        if round_ is not None:
            for i in range(len(self.data)):
                self.data[i] = self.data[i].round(round_)
        if output == 'df':
            return self.data
        elif output == 'excel':
            pass
        else:
            return self

    def get_data_wsd(self):
        if self.Date_List is None:
            if self.StartDate == self.EndDate:
                multi_data = []
                for i in self.code:
                    data_ = w.wsd(i, self.indicator, self.StartDate, self.EndDate, options=self.options, usedf=True)
                    multi_data.append(data_[1])
                df = self.create_date_col(pd.concat(multi_data, axis=0), self.EndDate) 
                self.data.append(df)
            else:
                multi_data = []
                for i in self.code:
                    data_ = w.wsd(i, self.indicator, self.StartDate, self.EndDate, options=self.options, usedf=True)
                    multi_data.append(data_[1])
                self.data.append(multi_data)
        else:
            for date in self.Date_List:
                multi_data = []
                for i in self.code:
                    data_ = w.wsd(i, self.indicator, date, date, options=self.options, usedf=True)
                    multi_data.append(data_[1])
                df = self.create_date_col(pd.concat(multi_data, axis=0), date)
                self.data.append(df)
        return self

    def get_data_wss(self):
        pass


    def excel_export(self, path=None, sheet_name:list=None, column_name:list=None):
        """
        导出为Excel
        =====
        path: str (default: None)
            导出路径, 例如'./data.xlsx', 默认为None, 会在output目录下生成{日期/开始日期_结束日期}.xlsx
        sheet_name: list (default: None)
            sheet名称, 例如['sheet1', 'sheet2'], 默认为None, 会使用工作簿数量作为sheet名称
        column_name: list (default: None)
            列名称, 例如['col1', 'col2'], 默认为None, 会使用Wind API返回的列名称
        """
        if path is None:
            if self.Date_List is None:
                if self.StartDate == self.EndDate:
                    path = f'./output/{self.EndDate}.xlsx'
                else:
                    path = f'./output/{self.StartDate}_{self.EndDate}.xlsx'
            else:
                path = f'./output/{self.Date_List[0]}.xlsx'
        if sheet_name is None:
            sheet_name = np.arange(len(self.data)) + 1
        if column_name is not None:
            for i in range(len(self.data)):
                self.data[i].columns = column_name
        with pd.ExcelWriter(path) as writer:
            for i in range(len(self.data)):
                self.data[i].to_excel(writer, sheet_name=f'{sheet_name[i]}')
        return self
    def df_export(self):
        return self.data