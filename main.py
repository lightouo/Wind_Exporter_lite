from wdpcore import Wind_Exporter
from datetime import datetime, timedelta
from chinese_calendar import is_holiday
from report_exp import report_export
code_for_zhai = ['005754.OF', '005756.OF', '008911.OF', '008913.OF', '007935.OF', '007936.OF', '008696.OF', '004827.OF', '006851.OF']
code_for_ETF = ['516760.OF', '516820.OF', '515700.OF', '561600.OF']
code_cundan_str = """015645.OF,015644.OF,015826.OF,015862.OF,014437.OF,015823.OF,015648.OF,014427.OF,015822.OF,015647.OF,015875.OF,
                        014426.OF,015864.OF,015861.OF,015646.OF,015643.OF,014430.OF,015956.OF,015825.OF,014428.OF,015827.OF,015863.OF,014429.OF,015944.OF,015955.OF,016082.OF,016063.OF,016083.OF"""
code_for_meiri_str = """002450.OF,004827.OF,015645.OF,008694.OF,005754.OF,700003.OF,000739.OF,007935.OF,009661.OF,009878.OF,
                        010126.OF,014460.OF,013767.OF,013687.OF,004390.OF,012475.OF,007893.OF,011828.OF,885001.WI"""
def export_data(choice_data):
    a = Wind_Exporter(code=code_for_zhai, indicator="sec_name,nav_date,nav,NAV_adj_return1,return_1m,return_3m,return_1y",options="annualized=1",method="wsd",EndDate=choice_data)
    b = Wind_Exporter(code=code_for_ETF, indicator="sec_name,nav_date,nav,NAV_adj_return1,return_1w,return_1m",options="annualized=1",method="wsd",EndDate=choice_data)
    c = Wind_Exporter(code="015645.OF", indicator="sec_name,nav_date,nav,NAV_adj_return1,return_1w,return_1m",options="annualized=1",method="wsd",EndDate=choice_data)
    d = Wind_Exporter(code=code_cundan_str, indicator="sec_name,nav_date,return,risk_annualintervalyield,issue_date,fund_setupdate",options="annualized=0",method="wss",StartDate="2022-07-01",EndDate=choice_data)
    e = Wind_Exporter(code=code_for_meiri_str, indicator="sec_name,nav_date,nav,NAV_adj_return1,NAV_adj_return,return_ytd",options="annualized=0",method="wss",StartDate="before1m",EndDate=choice_data)
    e_ = Wind_Exporter(code=code_for_meiri_str, indicator="NAV_adj_return",method="wss",StartDate="before1y",EndDate=choice_data)
    a.get_data(round_=4).add_data(b, method='append',round_=4).add_data(c, method='append',round_=4).add_data(d, method='append',round_=4).add_data(e, method='append',round_=4).add_data(e_, method='concat',round_=4)
    for i in a.data:
        date_col = ['NAV_DATE','ISSUE_DATE','FUND_SETUPDATE']
        for j in date_col:
            try:
                if j in i.columns:
                    i[j] = i[j].astype(str)
            except:
                pass

    a.excel_export(sheet_name=['债', 'ETF', '同业', '存单', '每日'], column_name=[['证券简称', '基金净值日期','单位净值', '当期复权单位净值增长率', '近1月回报', '近3月回报', '近1年回报'],
                    ['证券简称', '基金净值日期','单位净值', '当期复权单位净值增长率', '近1周回报', '近1月回报'],['证券简称', '基金净值日期','单位净值', '当期复权单位净值增长率', '近1周回报', '近1月回报'],
                    ['证券简称', '基金净值日期','区间回报', '区间收益率', '发行日期', '基金成立日'], ['证券简称', '基金净值日期','单位净值', '当期复权单位净值增长率', '复权单位净值增长率(截止日1月前)', '今年以来回报', '复权单位净值增长率(截止日1年前)']],
                    path='./output/{}.xlsx'.format(choice_data))



if __name__ == '__main__':
    choice_data = input('请输入日期(格式为YYYY-MM-DD),或者相对于今天的日期偏移值,(如-1代表昨天),不输入默认为今天:')
    if choice_data == '':
        choice_data = datetime.now()
    if choice_data != '':
        try:
            choice_data = int(choice_data)
            choice_data = datetime.now() + timedelta(days=choice_data)
        except:
            pass
        if isinstance(choice_data, str):
            try:
                choice_data = datetime.strptime(choice_data, '%Y-%m-%d')
            except:
                print('输入日期格式错误')
                exit()
    if is_holiday(choice_data):
        print('今天是假期，不需要获取数据')
        pass
    else:
        choice_data_ = choice_data.strftime('%Y-%m-%d')
        print(f'正在获取{choice_data_}的数据...')
        export_data(choice_data_)
        print('数据获取完成')
        print('正在生成报告...')
        report_export(choice_data)
        print('报告生成完成')