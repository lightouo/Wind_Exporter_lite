import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from WindPy import w
try:
    from typing import Literal
except ImportError:
    from typing_extensions import Literal
from dateutil.relativedelta import relativedelta


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

    def __init__(self, code=None, indicator: str = None, method: Literal['wsd', 'wss'] = 'wsd', 
                    Date_List: list = None, StartDate: Literal['before1m', 'before1y', None] = None, EndDate=None, options=None):
        self.method = method
        self.code = code if isinstance(code, list) else [
            i for i in code.split(',')]
        self.EndDate = EndDate if EndDate else datetime.today().strftime('%Y-%m-%d')
        self.Date_List = Date_List
        self.StartDate = EndDate if StartDate is None else StartDate
        self.set_date()
        self.indicator = indicator
        self.options = options
        self.data = []

    def __len__(self):
        return len(self.data)

    def __repr__(self):
        return f'Wind_Exporter({self.code}, {self.indicator}, {self.method}, {self.Date_List}, {self.StartDate}, {self.EndDate}, {self.options}), With data length {len(self)}'

    def check_connection(func):
        if not w.isconnected():
            w.start()
        return func

    @staticmethod
    def create_date_col(df, date):
        df['date'] = date
        column = np.roll(np.arange(len(df.columns)), 1)
        column[0], column[1] = column[1], column[0]
        return df.iloc[:, column]

    def set_date(self):
        date_before_1m = (datetime.strptime(self.EndDate, '%Y-%m-%d') -
                          relativedelta(months=1) + relativedelta(days=1)).strftime('%Y%m%d')
        date_before_1y = (datetime.strptime(self.EndDate, '%Y-%m-%d') -
                          relativedelta(years=1) + relativedelta(days=1)).strftime('%Y%m%d')
        if self.StartDate == 'before1m':
            self.StartDate = date_before_1m
        elif self.StartDate == 'before1y':
            self.StartDate = date_before_1y
        else:
            pass

    @check_connection
    def get_data(self, output: Literal['df', 'excel'] = None, round_=None):
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
                    data_ = w.wsd(i, self.indicator, self.StartDate,
                                  self.EndDate, options=self.options, usedf=True)
                    multi_data.append(data_[1])
                self.data.append(pd.concat(multi_data, axis=0))
            else:
                for i in self.code:
                    data_ = w.wsd(i, self.indicator, self.StartDate,
                                  self.EndDate, options=self.options, usedf=True)
                    self.data.append(data_[1])
        else:
            for date in self.Date_List:
                multi_data = []
                for i in self.code:
                    data_ = w.wsd(i, self.indicator, date, date,
                                  options=self.options, usedf=True)
                    multi_data.append(data_[1])
                self.data.append(pd.concat(multi_data, axis=0))
        return self

# w.wss("002450.OF","NAV_adj_return","startDate=20221117;endDate=20221118")
# w.wss("009878.OF", "sec_name","startDate=20221117;endDate=20221118")

    @check_connection
    def combine_wss(self, group_data: list):
        for i in group_data:
            data_ = []
            for j in i:
                data_.append(j[1])
            data_ = pd.concat(data_, axis=1)
            self.data.append(data_)
        return self

    def get_data_wss(self):
        self.StartDate = self.StartDate.replace('-', '')
        self.EndDate = self.EndDate.replace('-', '')
        args = [self.options, self.StartDate, self.EndDate]
        if self.options is None:
            args = args[1:]
        options = f'startDate={args[0]};endDate={args[1]};tradeDate={args[1]}' if len(
            args) == 2 else f'{args[0]};startDate={args[1]};endDate={args[2]};tradeDate={args[2]}'
        data_ = w.wss(self.code, self.indicator, options=options, usedf=True)
        self.data.append(data_[1])
        return self

    def add_data(self, we_obj, method: Literal['concat', 'append'] = 'concat', round_=None, axis=1):
        if method == 'concat':
            we_obj.get_data(round_=round_)
            if len(we_obj.data) == 1:
                self.data[-1] = pd.concat([self.data[-1], we_obj.data[0]], axis=axis)
            elif len(we_obj.data) == len(self.data):
                for i in range(len(we_obj.data)):
                    self.data[i] = pd.concat([self.data[i], we_obj.data[i]], axis=axis)
            else:
                raise ValueError('data length must be equal or 1')
        elif method == 'append':
            we_obj.get_data(round_=round_)
            for i in we_obj.data:
                self.data.append(i)
        else:
            raise ValueError('method must be concat or append')
        return self

    def excel_export(self, path=None, sheet_name: list = None, column_name: list = None):
        """
        导出为Excel
        =====
        path: str (default: None)
            导出路径, 例如'./data.xlsx', 默认为None, 会在output目录下生成{日期/开始日期_结束日期}.xlsx
        sheet_name: list (default: None)
            sheet名称, 例如['sheet1', 'sheet2'], 默认为None, 会使用工作簿数量作为sheet名称
        column_name: list (default: None)
            列名称, 二维数组, 例如[['col1', 'col2'], ['col1', 'col2']], 默认为None, 会使用Wind默认列名称
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
            sheet_name = [f'sheet{i}' for i in range(len(self.data))]
        if column_name is not None:
            try:
                _ = column_name[0][0]
            except:
                raise ValueError('column_name 应当是二维列表')
            for i, data in enumerate(self.data):
                try:
                    data.columns = column_name[i]
                except:
                    raise ValueError('列名数与数据不匹配')
        with pd.ExcelWriter(path, date_format='YYYY-MM-DD',datetime_format='YYYY-MM-DD', engine='openpyxl') as writer:
            for i in enumerate(self.data):
                i[1].to_excel(writer, sheet_name=sheet_name[i[0]])
                sheet = writer.sheets[sheet_name[i[0]]]
                for j in range(len(i[1].columns)+ 1):
                    sheet.column_dimensions[chr(64+j+1)].width = 26
        return self
