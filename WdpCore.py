from typing import Literal
from WindPy import w
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

class Wind_Exporter:
    def __init__(self, code:str|list, indicator:str, method:Literal['wsd','wss']='wsd', Date_List:list=None, StartDate=None, EndDate=None, options=None):
        self.method = method
        self.code = code
        self.EndDate = EndDate if EndDate else datetime.today().strftime('%Y-%m-%d')
        self.Date_List = Date_List
        self.StartDate = EndDate if StartDate is None else StartDate
        self.indicator = indicator
        self.options = options
        self.data = []

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
    def get_data(self, output:Literal[None,'df','excel']=None, round=None):
        if self.method == 'wsd':
            self.get_data_wsd()
        elif self.method == 'wss':
            self.get_data_wss()
        else:
            raise ValueError('method must be wsd or wss')


        if round is not None:
            for i in range(len(self.data)):
                self.data[i] = self.data[i].round(round)

        
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