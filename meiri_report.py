import pandas as pd
import datetime as dt
import numpy as np
import xlwings as xw
def meiri_report(date=dt.datetime.now()):
    zhai_data = pd.read_excel(f'./output/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='每日',index_col=0)
    zhai_data.iloc[:,[1]] = zhai_data.iloc[:,[1]].astype(np.datetime64)
    etf_data = pd.read_excel(f'./output/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='ETF',index_col=0)
    cun_data = pd.read_excel(f'./output/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='同业',index_col=0)

    app1 = xw.App(visible=False,add_book=False)
    app1.display_alerts = False
    book = app1.books.open('./template/每日播报.xlsm')
    book.sheets[0]["A2"].options(header=False).value = zhai_data
    book.sheets[1]["A2"].options(header=False).value = etf_data
    book.sheets[1]["K2"].value = cun_data["近1周回报"][0]
    book.sheets[3].activate()
    book.macro('Automatic.存单').run()
    book.macro('Automatic.ETF').run()
    book.macro('Automatic.Copy').run()
    book.save(f'./output/每日播报{date.strftime("%Y-%m-%d")}.xlsm')
    app1.quit()