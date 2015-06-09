# -*- coding: utf-8 -*-

import json
import urllib
import os
import win32com.client
import re


def geo_conv(orig, dest):

    para = {
            'ak': '5abc6f583d7ac0d217dd536768af0a10',
            'from': '5',
            'to': 6,
            }
    para['coords'] = '%s,%s;%s,%s'%(orig['x'], orig['y'], dest['x'], dest['y'])
    
    url = 'http://api.map.baidu.com/geoconv/v1/?'
    url = url + urllib.urlencode(para)
    data = urllib.urlopen(url).read()
    response = json.loads(data)
    if response['status'] != 0:
        print 'geo_conv url: ', url
        print "response", json.dumps(response, ensure_ascii=False, indent=2)

    orig['x'] = response['result'][0]['x']
    orig['y'] = response['result'][0]['y']
    dest['x'] = response['result'][1]['x']
    dest['y'] = response['result'][1]['y']


def map_content(orig, dest, sy):

    url = 'http://map.baidu.com/?'

    para = {
            'newmap': '1',
            'reqflag': 'pcmap',
            'biz': '1',
            'from': 'webmap',
            'da_par': 'baidu',
            'pcevaname': 'pc3',
            'qt': 'nav',
            'drag': '0',
            'reqtp': '1',
            'version': '4',
            #'mrs': '1',
            #'route_traffic': '1',
            'extinfo': '63',
            'tn': 'B_NORMAL_MAP',
            'nn': '0',
            'ie': 'utf-8',
            'l': '14',
            }

    # para['t'] = '1433835506095',  # time 注释无影响
    # para['c'] = '218' # city code 注释掉这个之后，没有打车费
    # para['sc'] = '218'  # start city code  注释无影响
    # para['ec'] = '218'  # end city code  注释无影响
    # para['b'] = '(12728551.55,3541379.45;12744279.55,3550051.45)'
    # para['sq'] = ('%s'%('华中科技大学')).decode('utf8').encode('gbk') 注释无影响
    # para['eq'] = ('%s'%('中南政法大学')).decode('utf8').encode('gbk') 注释无影响
    para['sn'] = ('1$$$$%s,%s$$中国$$$$$$'%(orig['x'], orig['y'])).decode('utf8').encode('gbk')
    para['en'] = ('1$$$$%s,%s$$中国$$$$$$'%(dest['x'], dest['y'])).decode('utf8').encode('gbk')
    para['sy'] = sy

    url = url + urllib.urlencode(para)
    print 'url\n', url
    data = urllib.urlopen(url).read()
    response = json.loads(data)
    # print "response\n", json.dumps(response, ensure_ascii=False, indent=2)
    # fb = open('b.txt', 'w')
    # fb.write(json.dumps(response, ensure_ascii=False, indent=2).encode('gbk'))
    # fb.close()
    return response


def set_result(content, ws, row, offset):

    route = content['content']['routes'][0]
    steps = content['content']['steps']

    if route['tab'] != '1_1':
        raise Exception("the first route's tab is not1_1!")

    r_s =  route['legs'][0]['duration']
    r_h = r_s/3600
    r_m = (r_s%3600 + 59)/60
    ws.Cells(row, 5+offset*4).Value = u'%s小时%s分'%(r_h, r_m)

    distance = float(route['legs'][0]['distance']) / 1000
    ws.Cells(row, 6+offset*4).Value = u'%.1f公里'%distance
    ws.Cells(row, 7+offset*4).Value = u'%d个'%route['light_num']
    ws.Cells(row, 8+offset*4).Value = route['main_roads']

    step_index = []
    step_inst = []
    for s in route['legs'][0]['stepis']:
        for i in range(s['s'], s['s']+s['n']):
            step_index.append(i)
    for i in step_index:
        step_inst.append(re.sub(r'</?\w+[^>]*>','', steps[i]['instructions']))
    ws.Cells(row, 17+offset).Value = u'；'.join(step_inst)


if __name__ == '__main__':

    wb_name = u'交通需求表.xlsx'
    ws_name = u'公路需求说明'

    rpath = os.path.split(os.path.realpath(__file__))[0]
    
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(os.path.join(rpath, wb_name))
    ws = wb.Worksheets(ws_name)
    row = 1

    while True:
        if ws.Cells(row, 1).Value == u'经度':
            break
        row += 1

    orig = {}
    dest = {}

    while True:
        row += 1
        if ws.Cells(row, 1).Value is None:
            break

        orig['x'] = ws.Cells(row,1) # 经度
        orig['y'] = ws.Cells(row,2) # 纬度
        dest['x'] = ws.Cells(row,3)
        dest['y'] = ws.Cells(row,4)
        geo_conv(orig, dest)

        content = map_content(orig, dest, sy=0) # 推荐路线
        set_result(content, ws, row, offset=0)

        content = map_content(orig, dest, sy=1) # 最短路线
        set_result(content, ws, row, offset=1)

        content = map_content(orig, dest, sy=2) # 不走高速
        set_result(content, ws, row, offset=2)

