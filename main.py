import datetime
import pandas as pd
from dateutil.relativedelta import relativedelta
from chinese_calendar import is_holiday
from WindPy import w
w.start()
code_for_zhai = ['005754.OF', '005756.OF', '008911.OF', '008913.OF', '007935.OF', '007936.OF', '008696.OF', '004827.OF', '006851.OF']
code_for_ETF = ['516760.OF', '516820.OF', '515700.OF', '561600.OF']
code_cundan_str = """015645.OF,015644.OF,015826.OF,015862.OF,014437.OF,015823.OF,015648.OF,014427.OF,015822.OF,015647.OF,015875.OF,014426.OF,015864.OF,015861.OF,015646.OF,015643.OF,014430.OF,015956.OF,015825.OF,014428.OF,015827.OF,015863.OF,014429.OF,015944.OF,015955.OF,016082.OF,016063.OF,016083.OF"""

def export_data_wsd(date=datetime.datetime.now().strftime('%Y-%m-%d')):
    writer = pd.ExcelWriter('./{}.xlsx'.format(date), date_format='YYYY-MM-DD')
    data_zhai = []
    data_tongye =[]
    data_etf = []
    for code in code_for_zhai:
        data_zhai.append(w.wsd(code, "sec_name,nav,NAV_adj_return1,return_1m,return_3m,return_1y", date, date, "annualized=1").Data)
    excel_export(data_zhai, writer, '债', date)
    for code in code_for_ETF:
        data_etf.append(w.wsd(code, "sec_name,nav,NAV_adj_return1,return_1w,return_1m", date, date, "annualized=1").Data)
    excel_export(data_etf, writer, 'ETF', date)
    data_tongye.append(w.wsd("015645.OF", "sec_name,nav,NAV_adj_return1,return_1w,return_1m", date, date, "annualized=1").Data)
    excel_export(data_tongye, writer, '同业', date)
    export_data_wss(endDate=date, writer=writer)

def export_data_wss(writer, startDate="2022-07-01" ,endDate=datetime.datetime.now().strftime('%Y-%m-%d')):
    start_date_ = startDate.replace('-', '') 
    end_date_ = endDate.replace('-', '')
    date_before_1m = (datetime.datetime.strptime(endDate, '%Y-%m-%d') - relativedelta(months=1) + relativedelta(days=1)).strftime('%Y%m%d')
    date_before_1y = (datetime.datetime.strptime(endDate, '%Y-%m-%d') - relativedelta(years=1) + relativedelta(days=1)).strftime('%Y%m%d')
    data_cundan = w.wss(code_cundan_str, "sec_name,return,risk_annualintervalyield,issue_date,fund_setupdate",f"annualized=0;startDate={start_date_};endDate={end_date_}")
    data_cundan_2 = w.wss(code_cundan_str, "return_std",f"annualized=1;tradeDate=20220920")
    sec_name = data_cundan.Data[0]
    return_area = data_cundan.Data[1]
    annualintervalyield = data_cundan.Data[2]
    issue_date = data_cundan.Data[3]
    fund_setupdate = data_cundan.Data[4]
    return_from_now = data_cundan_2.Data[0]
    data_cundan = pd.DataFrame({'证券简称':sec_name,'基金净值日期':endDate, '区间回报':return_area, '区间收益率':annualintervalyield, '发行日期':issue_date, '基金成立日':fund_setupdate, '成立以来回报':return_from_now}, index=[code_cundan_str.split(',')])
    data_cundan.iloc[:, 4:6] = data_cundan.iloc[:, 4:6].applymap(lambda x: x.strftime('%Y-%m-%d'))
    data_cundan.round(4).to_excel(writer, sheet_name='存单')

    code_for_meiri_str = """002450.OF,004827.OF,015645.OF,008694.OF,005754.OF,700003.OF,000739.OF,007935.OF,009661.OF,009878.OF,010126.OF,014460.OF,013767.OF,013687.OF,004390.OF,012475.OF,007893.OF,885001.WI"""
    data_meiri = w.wss(code_for_meiri_str, "sec_name,nav,NAV_adj_return1,NAV_adj_return,return_ytd",f"tradeDate={end_date_};startDate={date_before_1m};endDate={end_date_};annualized=0")
    data_meiri_2 = w.wss(code_for_meiri_str,"NAV_adj_return",f"startDate={date_before_1y};endDate={end_date_}")
    sec_name = data_meiri.Data[0]
    nav = data_meiri.Data[1]
    Nav_adj_return1 = data_meiri.Data[2]
    Nav_adj_return = data_meiri.Data[3]
    return_ytd = data_meiri.Data[4]
    nav_adj_return1_2 = data_meiri_2.Data[0]
    data_meiri = pd.DataFrame({'证券简称': sec_name,'基金净值日期':endDate, '单位净值': nav,'当期复权单位净值增长率':Nav_adj_return1, '复权单位净值增长率(截止日1月前)':Nav_adj_return, '今年以来回报':return_ytd, '复权单位净值增长率(截止日1年前)': nav_adj_return1_2}, index=[code_for_meiri_str.split(',')])
    data_meiri.round(4).to_excel(writer, sheet_name='每日')
    writer.close()

def excel_export(data, writer, sheet_name='Sheet1', date=datetime.datetime.now().strftime('%Y-%m-%d')):
    df = pd.DataFrame(data)
    df = df.applymap(lambda x: x[0]).round(4)
    if sheet_name == '债':
        df.columns = ['证券简称', '单位净值', '当期复权单位净值增长率', '近1月回报', '近3月回报', '近1年回报']
        df.index = code_for_zhai
        df['基金净值日期'] = date
        df.loc[:,['证券简称', '基金净值日期','单位净值', '当期复权单位净值增长率', '近1月回报', '近3月回报', '近1年回报']].to_excel(writer, sheet_name=sheet_name)
    elif sheet_name == 'ETF':
        df.columns = ['证券简称', '单位净值', '当期复权单位净值增长率', '近1周回报', '近1月回报']
        df.index = code_for_ETF
        df['基金净值日期'] = date
        df.loc[:,['证券简称', '基金净值日期','单位净值', '当期复权单位净值增长率', '近1周回报', '近1月回报']].to_excel(writer, sheet_name=sheet_name)
    elif sheet_name == '同业':
        df.columns = ['证券简称', '单位净值', '当期复权单位净值增长率', '近1周回报', '近1月回报']
        df.index = ['015645.OF']
        df['基金净值日期'] = date
        df.loc[:,['证券简称', '基金净值日期','单位净值', '当期复权单位净值增长率', '近1周回报', '近1月回报']].to_excel(writer, sheet_name=sheet_name)

def up_or_down(num):
    if num < 0:
        return '下跌📉'
    else:
        return '上涨📈'

def report_export(date=datetime.datetime.now()):
    data_ = w.wsd("000001.SH,881001.WI", "pct_chg", "2022-11-16", "2022-11-16", "").Data[0]
    data_ = [round(i, 2) for i in data_]
    # data_ = [1.2, -1.3]
    data_1 = pd.read_excel(f'./{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='每日',index_col=0).round(2)
    data_2 = pd.read_excel(f'./{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='同业',index_col=0).round(2)
    text_block_1 = f"""📍【市场指数表现】📍
上证指数{up_or_down(data_[0])}：{abs(data_[0])}%
万得全A{up_or_down(data_[1])}：{abs(data_[1])}%
    """
    text_block_2 = f"""🌼平安睿享文娱-黄维（002450）
本日{up_or_down(data_1['当期复权单位净值增长率']['002450.OF'])}：{abs(data_1['当期复权单位净值增长率']['002450.OF'])}%
近一月{up_or_down(data_1['复权单位净值增长率(截止日1月前)']['002450.OF'])}：{abs(data_1['复权单位净值增长率(截止日1月前)']['002450.OF'])}%
    """
    text_block_3 = f"""🌼平安策略先锋-神爱前（700003）
今日{up_or_down(data_1['当期复权单位净值增长率']['700003.OF'])}：{abs(data_1['当期复权单位净值增长率']['700003.OF'])}%
近一年{up_or_down(data_1['复权单位净值增长率(截止日1年前)']['700003.OF'])}：{abs(data_1['复权单位净值增长率(截止日1年前)']['700003.OF'])}%
    """
    text_block_4 = f"""🌼平安转型创新-神爱前（004390）
今日{up_or_down(data_1['当期复权单位净值增长率']['004390.OF'])}：{abs(data_1['当期复权单位净值增长率']['004390.OF'])}%
近一年{up_or_down(data_1['复权单位净值增长率(截止日1年前)']['004390.OF'])}：{abs(data_1['复权单位净值增长率(截止日1年前)']['004390.OF'])}%
    """
    text_block_5 = f"""🌼平安惠澜纯债A（007935）
今日{up_or_down(data_1['当期复权单位净值增长率']['007935.OF'])}：{abs(data_1['当期复权单位净值增长率']['007935.OF'])}%
近一年{up_or_down(data_1['复权单位净值增长率(截止日1年前)']['007935.OF'])}：{abs(data_1['复权单位净值增长率(截止日1年前)']['007935.OF'])}%
    """
    text_block_6 = f"""🌼平安低碳经济-何杰（009878）
今日{up_or_down(data_1['当期复权单位净值增长率']['009878.OF'])}：{abs(data_1['当期复权单位净值增长率']['009878.OF'])}%
近一月{up_or_down(data_1['复权单位净值增长率(截止日1月前)']['009878.OF'])}：{abs(data_1['复权单位净值增长率(截止日1月前)']['009878.OF'])}%
    """
    text_block_7 = f"""🌼平安同业存单（015645）
本日{up_or_down(data_2['当期复权单位净值增长率'][0])}：{abs(data_2['当期复权单位净值增长率'][0])}%
近七天年化{up_or_down(data_2['近1周回报'][0])}：{abs(data_2['近1周回报'][0])}%
"""
    text_block_8 = f"""🌼平安品质优选-神爱前（014460）
本日{up_or_down(data_1['当期复权单位净值增长率']['014460.OF'])}：{abs(data_1['当期复权单位净值增长率']['014460.OF'])}%
近一月{up_or_down(data_1['复权单位净值增长率(截止日1月前)']['014460.OF'])}：{abs(data_1['复权单位净值增长率(截止日1月前)']['014460.OF'])}%
"""
    data = f"""🏅 净值播报{date.strftime('%m.%d')}

{text_block_1}
{text_block_2}
{text_block_3}
{text_block_4}
{text_block_5}
{text_block_6}
{text_block_7}
{text_block_8}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—广发证券

{text_block_1}
📍【持营池产品】📍
权益可坚持定投
{text_block_2}
{text_block_7}
📍【纯债推荐】📍
{text_block_5}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—中金财富

{text_block_1}
📍【持营池产品】📍
{text_block_8}
📍【货币+】
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—粤开证券

{text_block_1}
📍【持营池产品】📍
{text_block_3}
📍【货币+】
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—国海证券

{text_block_1}
📍【持营池产品】📍
{text_block_4}
📍【货币+】
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—万联证券

{text_block_1}
📍【货币+】
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—招商证券

{text_block_1}
📍【核心公募产品】📍
{text_block_4}
{text_block_7}
📍【持营池产品】📍
{text_block_5}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—国信证券

{text_block_1}
📍【持营池产品】📍
{text_block_6}
📍【货币+】
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—国盛证券

{text_block_1}
📍【持营池产品】📍
{text_block_4}
📍【货币+】
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—长城证券

{text_block_1}
📍【货币+】
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]



🏅 净值播报{date.strftime('%m.%d')}—安信证券
{text_block_1}
📍【持营精选池】📍
货币+
{text_block_7}
平安基金与您携手同行
祝晚安[月亮]
    """
    with open(f'净报{date.strftime("%Y-%m-%d")}.txt', 'w', encoding='utf-8') as f:
        f.write(data)
if __name__ == '__main__':
    choice_data = input('请输入日期(格式为YYYY-MM-DD),不输入默认为今天:')
    if choice_data == '':
        choice_data = datetime.datetime.now().strftime('%Y-%m-%d')
    if choice_data != '':
        try:
            choice_data = datetime.datetime.strptime(choice_data, '%Y-%m-%d')
        except:
            print('输入日期格式错误')
            exit()

    if is_holiday(choice_data):
        print('今天是假期，不需要获取数据')
        pass
    else:
        print(f'正在获取{choice_data}的数据...')
        export_data_wsd(choice_data.strftime('%Y-%m-%d'))
        print('数据获取完成')
        print('正在生成报告...')
        report_export(choice_data)
        print('报告生成完成')