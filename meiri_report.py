import pandas as pd
import datetime as dt
import numpy as np
import xlwings as xw
from WdpCore import Wind_Exporter
import time
import os
from pathlib import Path

def get_picture(app,sheet_,range_:str,picture_name:str,path:str):
    app.books.add()
    chart_ = app.books[-1].sheets[0].charts.add(left=0,top=0,width=sheet_[range_].width,height=sheet_[range_].height)
    sheet_[range_].api.CopyPicture(Appearance=1,Format=-4147)
    time.sleep(0.1)
    chart_.api[1].Paste()
    path = Path(path).resolve().parent
    chart_.api[1].Export(f"{os.path.join(path, 'picture')}\\{picture_name}.png")
    app.books[-1].close()
    print(f"-----{picture_name}图片生成done-----")

def meiri_report(date=dt.datetime.now(), path='./output/'):
    zhai_data = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='每日',index_col=0)
    zhai_data.iloc[:,[1]] = zhai_data.iloc[:,[1]].astype(np.datetime64)
    etf_data = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='ETF',index_col=0)
    cun_data = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='同业',index_col=0)
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
    book.macro('Automatic.Until_date').run()

    for i in range(3,13):
        book.sheets[i]["B1"].value = f"优选基金&售后跟踪({date.strftime('%Y-%m-%d')})"

    book.save(f'{path}/每日播报{date.strftime("%Y-%m-%d")}.xlsx')

    pic_data = [["B1:I21","广发证券"],["B1:I15","中金财富"],["B1:I17","粤开证券"],["B1:I14","国海"],
                ["B1:I16","万联"],["B1:I22","招商证券"],["B1:I17","国信证券"],["B1:I14","国盛证券"],
                ["B1:I12","长城证券"],["B1:I13","安信证券"]]
    try:
        for n,i in enumerate(pic_data):
            get_picture(app1,book.sheets[n+3],i[0],f"{date.strftime('%Y-%m-%d')}{i[1]}",path)
    except:
        print("-----图片生成失败-----", "请检查是否有未关闭的Excel文件(包括进程)")
        quit_app()
    quit_app()  


def tongcun_rank_report(date=dt.datetime.now(), path='./output/'):
    cun_data = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='存单')
    cun_data_ = cun_data.sort_values(by='区间收益率',ascending=False).iloc[:,[0,1,3,4]]
    cun_data_['Rank'] = np.arange(1,len(cun_data_)+1)
    cun_data_ = cun_data_.iloc[:,[4,0,1,2,3]]
    app1 = xw.App(visible=False,add_book=False)
    app1.display_alerts = False
    book = app1.books.open('./template/全市场同存排名.xlsx')
    book.sheets[0]["A3"].options(index=False,header=False).value = cun_data_
    i= int(book.sheets[0]["B34"].value)
    book.sheets[0][f"A{i}:E{i}"].api.Interior.Color = 65535
    book.sheets[0]["A1"].value = f"全市场同存排名({date.strftime('%Y')}年{date.strftime('%m')}月{date.strftime('%d')}日)"

    book.save(f'{path}/全市场同存排名{date.strftime("%Y-%m-%d")}.xlsx')
    time.sleep(3)
    try:
        get_picture(app1,book.sheets[0],"A1:E33",f"{date.strftime('%Y-%m-%d')}全市场同存排名",path)
    except:
        print("-----图片生成失败-----", "请检查是否有未关闭的Excel文件(包括进程)")
        quit_app()
    quit_app()  

def tongcun_report(date=dt.datetime.now(),path='./output/'):
    a = Wind_Exporter(code="015645.OF", indicator="nav_date,nav,NAV_adj_return1,return_1w,return_1m",options="annualized=1",method="wsd",StartDate="ED-21TD",EndDate=date.strftime("%Y-%m-%d"))
    a.get_data()
    b = a.data[0]
    b.NAV_ADJ_RETURN1 = (b.NAV_ADJ_RETURN1/100).round(4)
    b.RETURN_1W = (b.RETURN_1W/100).round(4)
    day30 = b.RETURN_1M[-1]/100
    b.iloc[:,[0,1,2]]
    b.RETURN_1W
    app1 = xw.App(visible=False,add_book=False)
    app1.display_alerts = False
    book = app1.books.open('./template/平安同存收益率.xlsx')
    sheet_ = book.sheets[0]
    sheet_["A4"].options(index=False,header=False).value = b.iloc[:,[0,1,2]]
    sheet_["F4"].options(index=False,header=False).value = b.RETURN_1W
    sheet_["F2"].value = day30
    book.save(f'{path}/平安同存收益率{date.strftime("%Y-%m-%d")}.xlsx')
    time.sleep(3)
    try:
        get_picture(app1,book.sheets[0],"A1:F29",f"{date.strftime('%Y-%m-%d')}平安同存收益率",path)
    except:
        print("-----图片生成失败-----", "请检查是否有未关闭的Excel文件(包括进程)")
        quit_app()
    quit_app()

def quit_app():
    for i in xw.apps:
        i.quit()