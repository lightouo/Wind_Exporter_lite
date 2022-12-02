import xlwings as xw
import datetime as dt
from WdpCore import Wind_Exporter

def bond_report(date = dt.datetime.today()):
    code_for_zhai = ['005754.OF', '005756.OF', '007935.OF', '007936.OF']
    info = [
    """购买信息：
    A类(005754)申购费率0.10%-0.30%，赎回费率0.00%-1.50%（>30天赎回费率=0）""",
    """购买信息：
    E类(005756)销售服务费率0.25%（每年），赎回费率0.00%-1.50%（＞30天赎回费率=0）""",
    """购买信息：
    A类(007935)申购费率：M＜100万，费率0.8%；100万≤M＜300万，费率0.50%；300万≤M＜500万，费率0.30%；M≥500万，每笔1000元；赎回费率：0.00%-1.50%（≥30天赎回费率=0）""",
    """购买信息：
    C类(007936)销售服务费率0.50%（每年）；赎回费率：0.00%-1.50%（≥30天赎回费率=0）"""]


    # 1. 获取数据
    a = Wind_Exporter(code=code_for_zhai, indicator="sec_name,nav_date,nav,NAV_adj_return1,return_1y",options="annualized=1",method="wsd",StartDate="ED-21TD",EndDate=date.strftime("%Y-%m-%d"))
    a.get_data()
    data = []
    for i,code in enumerate(code_for_zhai):
        b = a.data[i].sort_values(by='NAV_DATE', ascending=False)
        name = b['SEC_NAME'].iloc[0]
        b.NAV_ADJ_RETURN1 = (b.NAV_ADJ_RETURN1/100).round(4)
        b.RETURN_1Y = (b.RETURN_1Y/100).round(4)
        data.append([b,name,code])

    print("数据获取完毕")

    # 2. 写入Excel
    invisible_app = xw.App(visible=False)
    book = invisible_app.books.open('./template/zhai2_tem.xlsx')
    for i in range(len(data)//2):
        i_ = i*2
        sheet_ = book.sheets[0]
        sheet_.name = data[i_][1]
        sheet_["A4"].options(index=False, header=False).value = data[i_][0].sort_values(by='NAV_DATE', ascending=False).loc[:,["NAV_DATE","NAV","NAV_ADJ_RETURN1"]]
        sheet_["F4"].options(index=False, header=False).value = data[i_][0].sort_values(by='NAV_DATE', ascending=False).loc[:,["RETURN_1Y"]]
        sheet_["A1"].value = f"{data[i_][1]}（{data[i_][2].replace('.OF','')}）收益情况播报"
        sheet_["A26"].value = info[i_]

        sheet_["A31"].options(index=False, header=False).value = data[i_+1][0].sort_values(by='NAV_DATE', ascending=False).loc[:,["NAV_DATE","NAV","NAV_ADJ_RETURN1"]]
        sheet_["F31"].options(index=False, header=False).value = data[i_+1][0].sort_values(by='NAV_DATE', ascending=False).loc[:,["RETURN_1Y"]]
        sheet_["A28"].value = f"{data[i_+1][1]}（{data[i_+1][2].replace('.OF','')}）收益情况播报"
        sheet_["A53"].value = info[i_+1]
        sheet_.copy()
    sheet_.delete()
    book.save(f"./output/债券基金收益情况播报_{date.strftime('%Y-%m-%d')}.xlsx")
    invisible_app.quit()
    print("数据写入完毕")