# -*- coding: utf-8 -*-

import urllib
import urllib2
import os
import win32com.client
from bs4 import BeautifulSoup, SoupStrainer
import traceback
import datetime


city={
u'阿克苏':('AKU','KUQA',),  # <- 库车
u'库车':('KCA',),  # -> 阿克苏
u'阿勒泰':('AAT','FYN',),  # <- 富蕴
u'安康':('AKA',),
u'安庆':('AQG',),
u'鞍山':('AOG',),
u'保山':('BSD',),
u'包头':('BAV',),
u'北海':('BHY',),
u'北京':('PEK','NAY',),
u'蚌埠':('BFU',),
u'长春':('CGQ',),
u'常德':('CGD',),
u'长沙':('CSX',),
u'长治':('CIH',),
u'常州':('CZX',),
u'朝阳':('CHG',),
u'成都':('CTU',),
u'赤峰':('CIF',),
u'重庆':('CKG','WXN',), # <- 万县
u'达县':('DAX',),
u'达州':('DAX',),  # <- 达县
u'大连':('DLC',),
u'大理':('DLU',),
u'丹东':('DDG',),
u'大同':('DAT',),
u'东营':('DOY',),
u'敦煌':('DNH',),
u'恩施':('ENH',),
u'阜阳':('FUG',),
u'富蕴':('FYN',),  # ->阿勒泰
u'福州':('FOC',),
u'赣州':('KOW',),
u'格尔木':('GOQ',),
u'广汉':('GHN',),
u'广州':('CAN',),
u'桂林':('KWL',),
u'贵阳':('KWE',),
u'哈尔滨':('HRB',),
u'海口':('HAK',),
u'海拉尔':('HLD',),
u'哈密':('HMI',),
u'杭州':('HGH',),
u'汉中':('HZG',),
u'合肥':('HFE',),
u'黑河':('HEK',),
u'香港':('HKG',),
u'衡阳':('HNY',),
u'和田':('HTN',),
u'黄山':('TXN',),
u'黄岩':('HYN',),
u'台州':('HYN',),  # 黄岩 -> 台州
u'呼和浩特':('HET',),
u'吉安':('KNC',),
u'佳木斯':('JMU',),
u'嘉峪关':('JGN',),
u'吉林':('JIL',),
u'济南':('TNA',),
u'济宁':('JNG',),
u'景德镇':('JDZ',),
u'景洪':('JHG',),
u'晋江':('JJN',),
u'泉州':('JJN',), # 晋江 -> 泉州
u'锦州':('JNZ',),
u'酒泉':('CHW',),
u'九江':('JIU',),
u'九寨黄龙':('JZH',),
u'阿坝':('JZH',), # 九寨 -> 阿坝
u'克拉玛依':('KRY',),
u'喀什':('KHG',),
u'库尔勒':('KRL',),
u'且末':('IQM',),
u'巴音郭楞蒙古自治州':('KRL','IQM',),  # 库尔勒, 且末
u'昆明':('KMG',),
u'兰州':('LHW',),
u'拉萨':('LXA',),
u'连云港':('LYG',),
u'丽江':('LJG',),
u'临沂':('LYI',),
u'柳州':('LZH',),
u'洛阳':('LYA',),
u'泸州':('LZO',),
u'澳门':('MFM',),
u'芒市':('LUM',),
u'德宏':('LUM',), # <- 芒市
u'满洲里':('NZH',),
u'呼伦贝尔':('NZH',), # <- 满洲里
u'梅县':('MXZ',),
u'梅州':('MXZ',), # <- 梅县
u'绵阳':('MIG',),
u'牡丹江':('MDG',),
u'南昌':('KHN',),
u'南充':('NAO',),
u'南京':('NKG',),
u'南宁':('NNG',),
u'南通':('NTG',),
u'南阳':('NNY',),
u'宁波':('NGB',),
u'青岛':('TAO',),
u'庆阳':('IQN',),
u'秦皇岛':('SHP',),
u'齐齐哈尔':('NDG',),
u'泉州':('JJN',),
u'衢州':('JUZ',),
u'三亚':('SYX',),
u'上海':('SHA','PVG',),
u'汕头':('SWA',),
u'沙市':('SHS',),
u'荆州':('SHS',), # <- 沙市
u'深圳':('SZX',),
u'沈阳':('SHE',),
u'石家庄':('SJW',),
u'思茅':('SYM',),
u'苏州':('SZV',),
u'塔城':('TCG',),
u'太原':('TYN',),
u'天津':('TSN',),
u'通化':('TNH',),
u'通辽':('TGO',),
u'铜仁':('TEN',),
u'乌鲁木齐':('URC',),
u'万县':('WXN',), # -> 重庆
u'潍坊':('WEF',),
u'威海':('WEH',),
u'温州':('WNZ',),
u'武汉':('WUH','WJD',),
u'乌兰浩特':('HLH',),
u'武夷山':('WUS',),
u'南平':('WUS',), # <- 武夷山
u'无锡':('WUX',),
u'梧州':('WUZ',),
u'厦门':('XMN',),
u'西安':('XIY',),
u'襄樊':('XFN',),
u'西昌':('XIC',),
u'锡林浩特':('XIL',),
u'西宁':('XNN',),
u'徐州':('XUZ',),
u'延安':('ENY',),
u'延吉':('YNJ',),
u'烟台':('YNT',),
u'盐城':('YNZ',),
u'宜宾':('YBP',),
u'宜昌':('YIH',),
u'银川':('INC',),
u'伊宁':('YIN',),
u'伊犁哈萨克':('YIN',), # <- 伊宁
u'义乌':('YIW',),
u'金华':('YIW',), # <- 义乌
u'永州':('LLF',),
u'榆林':('UYN',),
u'昭通':('ZAT',),
u'张家界':('DYG',),
u'湛江':('ZHA',),
u'芷江':('HJJ',),
u'怀化':('HJJ',), # <- 芷江
u'中甸':('DIG',),
u'迪庆':('DIG',), # <- 中甸
u'郑州':('CGO',),
u'舟山':('HSN',),
u'珠海':('ZUH',),
u'遵义' :('ZYI',),
}

air_lines ={
'CZ' : u'中国南方航空公司', 
'MU' : u'中国东方航空公司', 
'CA' : u'中国国际航空公司', 
'HU' : u'海南航空公司', 
'MF' : u'厦门航空公司', 
'FM' : u'上海航空公司', 
'ZH' : u'深圳航空公司', 
'SC' : u'山东航空公司', 
'3U' : u'四川航空公司', 
'EU' : u'鹰联航空有限公司', 
'BK' : u'奥凯航空公司', 
'KN' : u'中国联合航空公司',
'8C' : u'东星航空公司',
'HO' : u'吉祥航空公司',
'G5' : u'华夏航空公司',
'8L' : u'祥鹏航空公司'
}

air_name = {
u'AOG': u'鞍山腾鳌机场',
}

def air_pair(fa, ta):
    result = []
    for fi in fa:
        for ti in ta:
            apair = {'from':fi, 'to': ti}
            result.append(apair)
    return result


def get_sn(para):

    url = 'http://webflight.linkosky.com/WEB/Flight/WaitingSearch.aspx?'
    headers = {
            'Referer': 'http://www.caac.gov.cn/S1/GNCX/',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:38.0) Gecko/20100101 Firefox/38.0',
            'Cookie': 'ASP.NET_SessionId=jj410qerwzch10jua5ilk02s',
            'Host': 'webflight.linkosky.com', 
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            }
    url = url + urllib.urlencode(para)
    req = urllib2.Request(url, headers=headers)
    print 'url\n', req.get_full_url()

    response = urllib2.urlopen(req)
    resp_data = response.read()
    # print resp_data.decode('utf8').encode('gbk')
    pos = resp_data.find('Sn=')
    if pos == -1:
        print 'cannot find Sn!!!!!!!'
    sn = resp_data[pos+3: pos+35]
    print 'Get Sn=%s'%(sn)
    return sn


def load_html(f, t):

    para = {
            'AL': 'ALL',
            'BD': '',
            'BT': '7',
            'DR': 'true',
            'DT': '7',
            'JT': '1',
            'dst2': 'CAN',
            'dstDesp': 'GUANGZHOU广州'.decode('utf8').encode('gbk'),
            'image.x': '36',
            'image.y': '13'
            }
    now = datetime.datetime.now()
    dd = now + datetime.timedelta(days=3)
    para['DD'] = dd.strftime('%Y-%m-%d')
    para['OC'] = f
    para['DC'] = t
    para['Sn'] = get_sn(para)
    
    url = 'http://webflight.linkosky.com/WEB/Flight/FlightSearchResultDefault.aspx?'
    url = url + urllib.urlencode(para)
    print 'url: ', url
    data = urllib.urlopen(url).read()
    print 'get data'
    return data


def parse_air(menu, content, f_c, t_c):
    result = []
    menu_texts = [text for text in menu.stripped_strings]
    con_texts = [text for text in content.stripped_strings]
    
    if len(menu_texts) == 2:
        air_line_sn = menu_texts[0]
        air_line_comp = air_lines.get(air_line_sn[0:2], u'未指明航空公司')
        result.append(air_line_comp) # 航空公司
        result.append(air_line_sn) # 航班号
        result.append((menu_texts[1].split(u'：'))[1]) # 机型
    else:
        result.append(menu_texts[0]) # 航空公司
        result.append(menu_texts[1]) # 航班号
        result.append((menu_texts[2].split(u'：'))[1]) # 机型

    f_add = 0 if u':' in con_texts[0] else 1
    t_add = 0 if u':' in con_texts[f_add + 2] else 1

    result.append((con_texts[f_add + 0].lstrip(u'（').rstrip(u'）'))) # 出发时间
    result.append((con_texts[f_add + t_add + 2].lstrip(u'（').rstrip(u'）'))) # 到达时间
    
    if f_add == 1:
        f_name = con_texts[0]
        if f_c not in air_name.keys():
            air_name[f_c] = f_name
    else:
        f_name = air_name.get(f_c, u'未指明机场')
    result.append(f_name) # 出发机场
    
    if t_add == 1:
        t_name = con_texts[f_add + 2]
        if t_c not in air_name.keys():
            air_name[t_c] = t_name
    else:
        t_name = air_name.get(t_c, u'未指明机场')
    result.append(t_name) # 到达机场

    JT = (con_texts[f_add + 1].split(u'：'))[1]  # 是否经停
    if JT == '0':
        result.append(u'否')
    else:
        result.append(u'是')
    set_price = []
    price_begin = 3 + f_add + t_add
    text_len = (len(con_texts)-price_begin)//3*3+price_begin
    for i in range(price_begin, text_len, 3):
        set_price.append('%s%s'%(con_texts[i+1], con_texts[i]))
    result.append(u'；'.join(set_price))  # 舱位价格

    return result

def get_air(fa, ta):
    
    result = []
    air_line_sn = []
    pair_list = air_pair(fa, ta)
    for pair in pair_list:
        data = load_html(pair['from'], pair['to'])
        flight = SoupStrainer(id='FlightListFlight0')
        soup = BeautifulSoup(data, "html.parser", parse_only = flight, from_encoding='gb18030')
        if len(soup.contents) == 0:
            if u'没有满足条件的航班'.encode('utf8') in data:
                print u'from:%s,to:%s. 没有满足条件的航班'%(pair['from'], pair['to'])
            else:
                print u'from:%s,to:%s. 解析HTML文档出现错误'%(pair['from'], pair['to'])
            continue
        divs = soup.contents[0].find_all('div', recursive=False)
        div_len = len(divs)
        if div_len == 2:
            print soup.text
        elif div_len % 2 != 0 or div_len < 3:
            print 'div_len is %d, something wrong!'%(div_len)
            continue
        else:
            for i in range(2, div_len, 2):
                menu = divs[i]
                content = divs[i+1]
                air_line = parse_air(menu, content, pair['from'], pair['to'])
                if air_line[1] not in air_line_sn:
                    result.append(air_line)
                    air_line_sn.append(air_line[1])

    return result


if __name__ == '__main__':

    wb_name = u'交通需求表.xlsx'
    try:
        rpath = os.path.split(os.path.realpath(__file__))[0]
    except NameError:  # We are the main py2exe script, not a module
        import sys
        rpath = os.path.split(os.path.realpath(sys.argv[0]))[0]
    wb = None
    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True
        wb = excel.Workbooks.Open((os.path.join(rpath.decode('gbk'), wb_name)).encode('gbk'))
        ws = wb.Worksheets(1)
        row = 1
        while True:
            if ws.Cells(row, 1).Value == u'出发城市':
                break
            row += 1
        row += 1 # 第一行真实数据
        while True:
            if ws.Cells(row, 1).Value is None:
                print 'assumed the empty Cell(%d,A) is the end...'%row                
                break
            
            rxc3 = ws.Cells(row, 3).Value
            if rxc3 is not None and rxc3 == u'没有满足条件的航班':
                row = row + 1
                continue

            rxc12 = ws.Cells(row, 12).Value
            if rxc12 is not None:
                step = int(rxc12)  # 已经计算过并且没有错误                
                row = row + step   # 下一个
                continue
    
            c_from = ws.Cells(row,1).Value
            c_to = ws.Cells(row,2).Value
            print u'==============%s==>>%s==============='%(c_from, c_to)
    
            error = False
            fa = city.get(c_from, None)
            ta = city.get(c_to, None)
            if fa is None:
                print u'出发城市没有机场'
                if ws.Cells(row, 8).Value is None:
                    ws.Cells(row, 8).Value = u'出发城市没有机场'
            if ta is None:
                print u'到达城市没有机场'
                if ws.Cells(row, 9).Value is None:
                    ws.Cells(row, 9).Value = u'到达城市没有机场'
            if fa is None or ta is None:
                row = row + 1 # 下一个               
                continue
    
            air_data = get_air(fa, ta)
            air_num = len(air_data)
            if air_num == 0:
                ws.Cells(row, 3).Value = u'没有满足条件的航班'
                row = row + 1 # 下一个
                wb.Save() # 保存一下
                continue

            for i in range(1, air_num):
                rangeObj = ws.Range('C%d:K%d'%(row+i, row+i))
                rangeObj.EntireRow.Insert()
            if air_num > 1:
                ws.Range('A%d:B%d'%(row, row)).Copy()
                ws.Paste(ws.Range('A%d:B%d'%(row+1, row+air_num-1)))
            for i in range(0, air_num):
                rangeObj = ws.Range('C%d:K%d'%(row+i, row+i))
                rangeObj.Value = air_data[i]
                ws.Cells(row+i, 12).Value = air_num
            if row % 10 == 0:
                wb.Save() # 每做完一个保存一下
            row = row + air_num # 下一个
    except:
        if wb is not None:
            wb.Save()
        print "Unexpected error:", traceback.print_exc() # 出错信息

    raw_input('input any key to close.')
