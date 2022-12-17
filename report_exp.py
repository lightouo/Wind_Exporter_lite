from WindPy import w
import pandas as pd
import datetime
from decimal import Decimal,ROUND_HALF_UP



def round_half_up(number, ndigits):
    if isinstance(number, list):
        return [Decimal(str(num)).quantize(Decimal('0.' + '0' * ndigits), rounding=ROUND_HALF_UP) for num in number]
    if isinstance(number, float):
        return Decimal(str(number)).quantize(Decimal('0.' + '0' * ndigits), rounding=ROUND_HALF_UP)


def up_or_down(num):
    if num < 0:
        return '下跌📉'
    else:
        return '上涨📈'

def report_export(date=datetime.datetime.now(), path='./output'):
    data_ = w.wsd("000001.SH,881001.WI", "pct_chg", date.strftime("%Y-%m-%d"), date.strftime("%Y-%m-%d"), "").Data[0]
    if None in data_:
        print('Warning: 报告存在缺失数据，请检查现在数据是否已经公布!')
    data_ = [round(i, 2) if i is not None else 0 for i in data_]
    # data_ = [1.2, -1.3]
    data_1 = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='每日',index_col=0).round(2)
    data_2 = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='同业',index_col=0).round(2)
    data_3 = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='存单',index_col=0).round(2)
    data_4 = pd.read_excel(f'{path}/{date.strftime("%Y-%m-%d")}.xlsx',sheet_name='债',index_col=0).round(2)
    date_only5 = datetime.datetime.strptime(data_1["基金净值日期"]['015510.OF'], "%Y-%m-%d")
# 共享文字内容
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
    text_block_9 = f"""🌼平安价值领航-何杰（015510）
本日{up_or_down(data_1['当期复权单位净值增长率']['015510.OF'])}：{abs(data_1['当期复权单位净值增长率']['015510.OF'])}%（数据截至{date_only5.strftime("%m")}月{date_only5.strftime("%d")}日）
近一月{up_or_down(data_1['复权单位净值增长率(截止日1月前)']['015510.OF'])}：{abs(data_1['复权单位净值增长率(截止日1月前)']['015510.OF'])}%（数据截至{date_only5.strftime("%m")}月{date_only5.strftime("%d")}日）
"""

# 额外内容
    cun_data = [data_2['近1周回报']['015645.OF'], data_3['区间收益率']['015645.OF']]
    cun_data = round_half_up(cun_data, 2)
    zhai_data = [data_4['近1年回报'][f'{i}'] for i in ['005754.OF', '005756.OF']]
    zhai_data = round_half_up(zhai_data, 2)
    zhai_data_2 = [data_4['近1年回报'][f'{i}'] for i in ['007935.OF', '007936.OF']]
    zhai_data_2 = round_half_up(zhai_data_2, 2)


# 主体内容框架
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
📍【首发产品】📍
{text_block_9}
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


‼ 粤开重点持营产品：【平安同业存单】‼

💨市场风大雨大，不想冒险怎么办？
✌新型闲钱投资利器来了！ 【平安同业存单指数基金】
🌟基金代码：015645

货币替代好工具，产品定位货币增强30~50bp，因货币基金采用“摊余成本法”，按万份收益计算，看上去没有波动；而同存基金采用“市值法”，所以每日会有净值波动，但是风险级别属于除了货币外最低的产品了，即使有短期波动也无需担心。7月以来年化收益率{cun_data[1]}%，近7日年化收益率{cun_data[0]}%，而同期货币基金收益率进入“1”时代。 ‼

⛑适合想短期防御风险的投资者
🎈超短久期品种，波动较小，持有体验好
🧐低风险、安全性高 
💌不收取认购费用、赎回费
🌟7天最短持有期到期后可随时赎回
[爱心]现有额度充足，买入即确认，产品不配售

数据来源：wind，截至{date.strftime('%Y')}年{date.strftime('%m')}月{date.strftime('%d')}日。


‼ 安信重点持营产品：【平安同业存单】‼

💨市场风大雨大，不想冒险怎么办？
✌新型闲钱投资利器来了！ 【平安同业存单指数基金】
🌟基金代码：015645

货币替代好工具，产品定位货币增强30~50bp，因货币基金采用“摊余成本法”，按万份收益计算，看上去没有波动；而同存基金采用“市值法”，所以每日会有净值波动，但是风险级别属于除了货币外最低的产品了，即使有短期波动也无需担心。7月以来年化收益率{cun_data[1]}%，近7日年化收益率{cun_data[0]}%，而同期货币基金收益率进入“1”时代。 ‼

⛑适合想短期防御风险的投资者
🎈超短久期品种，波动较小，持有体验好
🧐低风险、安全性高 
💌不收取认购费用、赎回费
🌟7天最短持有期到期后可随时赎回
[爱心]现有额度充足，买入即确认，产品不配售

数据来源：wind，截至{date.strftime('%Y')}年{date.strftime('%m')}月{date.strftime('%d')}日。


‼ 国海重点持营产品：【平安同业存单】‼

💨市场风大雨大，不想冒险怎么办？
✌新型闲钱投资利器来了！ 【平安同业存单指数基金】
🌟基金代码：015645

货币替代好工具，产品定位货币增强30~50bp，因货币基金采用“摊余成本法”，按万份收益计算，看上去没有波动；而同存基金采用“市值法”，所以每日会有净值波动，但是风险级别属于除了货币外最低的产品了，即使有短期波动也无需担心。7月以来年化收益率{cun_data[1]}%，近7日年化收益率{cun_data[0]}%，而同期货币基金收益率进入“1”时代。 ‼

⛑适合想短期防御风险的投资者
🎈超短久期品种，波动较小，持有体验好
🧐低风险、安全性高 
💌不收取认购费用、赎回费
🌟7天最短持有期到期后可随时赎回
[爱心]现有额度充足，买入即确认，产品不配售

数据来源：wind，截至{date.strftime('%Y')}年{date.strftime('%m')}月{date.strftime('%d')}日。



‼ 国盛重点持营产品：【平安同业存单】‼

💨市场风大雨大，不想冒险怎么办？
✌新型闲钱投资利器来了！ 【平安同业存单指数基金】
🌟基金代码：015645

货币替代好工具，产品定位货币增强30~50bp，因货币基金采用“摊余成本法”，按万份收益计算，看上去没有波动；而同存基金采用“市值法”，所以每日会有净值波动，但是风险级别属于除了货币外最低的产品了，即使有短期波动也无需担心。7月以来年化收益率{cun_data[1]}%，近7日年化收益率{cun_data[0]}%，而同期货币基金收益率进入“1”时代。 ‼

⛑适合想短期防御风险的投资者
🎈超短久期品种，波动较小，持有体验好
🧐低风险、安全性高 
💌不收取认购费用、赎回费
🌟7天最短持有期到期后可随时赎回
[爱心]现有额度充足，买入即确认，产品不配售

数据来源：wind，截至{date.strftime('%Y')}年{date.strftime('%m')}月{date.strftime('%d')}日。



🟠平安惠澜纯债🟠

❗❗

📈A类(007935)近1年年化收益率{zhai_data_2[0]}%；C类（007936）近1年年化收益率{zhai_data_2[1]}%。


近30日及近1年收益情况请见上表，供各位老师参考[抱拳]

（数据来源：wind，截至{date.strftime('%Y')}年{date.strftime('%m')}月{date.strftime('%d')}日）




🟠平安短债🟠

❗单日限额提升到20万每日❗

📈A类(005754)近1年年化收益率{zhai_data[0]}%；E类（005756）近1年年化收益率{zhai_data[1]}%。


近30日收益情况请见上表，供各位老师参考[抱拳]

（数据来源：wind，截至{date.strftime('%Y')}年{date.strftime('%m')}月{date.strftime('%d')}日）





    """
    with open(f'{path}/净报{date.strftime("%Y-%m-%d")}.txt', 'w', encoding='utf-8') as f:
        f.write(data)