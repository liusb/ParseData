# -*- coding: utf-8 -*-

import json
import urllib
import os
import win32com.client

def Direction(orig, dest, tactics, f = 0):

    para = {
            'mode': 'driving',
            'output': 'json',
            'coord_type': 'gcj02',
            'ak': '5abc6f583d7ac0d217dd536768af0a10',
            }

    para['origin'] = '%s|%s,%s'%(orig['name'], orig['lat'], orig['lng'])
    para['destination'] ='%s|%s,%s'%(dest['name'], dest['lat'], dest['lng'])
    para['origin_region'] = orig['city']
    para['destination_region'] = dest['city']
    para['tactics'] = tactics    
    
    url = 'http://api.map.baidu.com/direction/v1?'
    url = url + urllib.urlencode(para)
    print 'url\n', url
    data = urllib.urlopen(url).read()
    response = json.loads(data)
    if response['status'] != 0:
        print "response_data", json.dumps(response, ensure_ascii=False, indent=2)
    
    if f == 1:
        fb = open('a.txt', 'w')
        fb.write(json.dumps(response, ensure_ascii=False, indent=4).encode('gbk'))
        fb.close()

    return response


def Geocoding(location):
    para = {
            'output': 'json',
            'ak': '5abc6f583d7ac0d217dd536768af0a10',
            'coordtype': 'gcj02ll'
            }

    para['location'] = '%s,%s'%(location['lat'], location['lng']) #纬度,经度
    
    url = 'http://api.map.baidu.com/geocoder/v2/?'
    url = url + urllib.urlencode(para)
    print 'url\n', url
    data = urllib.urlopen(url).read()
    response = json.loads(data)
    if response['status'] != 0:
        print "error_response", json.dumps(response,
                ensure_ascii=False, indent=2)

    return response


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

    origin = {}
    dest = {}

    while True:
        row += 1
        if ws.Cells(row, 1).Value is None:
            break

        origin['lng'] = ws.Cells(row,1) # 经度
        origin['lat'] = ws.Cells(row,2) # 纬度
        dest['lng'] = ws.Cells(row,3)
        dest['lat'] = ws.Cells(row,4)
        origGeo = Geocoding(origin)
        destGeo = Geocoding(dest)
        origin['city'] = origGeo['result']['addressComponent']['city'].encode('gb2312')
        origin['name'] = origGeo['result']['formatted_address'].encode('gb2312')
        dest['city'] = destGeo['result']['addressComponent']['city'].encode('gb2312')
        dest['name'] = destGeo['result']['formatted_address'].encode('gb2312')
        
        tactics = 11 # 最少时间(推荐)
        direct = Direction(origin, dest, tactics, f=1)
        result = direct['result']['routes'][0]
        r_s =  int(result['duration'])
        r_h = r_s/3600
        r_m = (r_s%3600)/60
        ws.Cells(row,5).Value = u'%s小时%s分'%(r_h, r_m)
        distance = float(result['distance']) / 1000
        ws.Cells(row,6).Value = u'%.1f公里'%distance
        ws.Cells(row,17).Value = 17
        
        tactics = 12 # 最短路程
        direct = Direction(origin, dest, tactics)
        result = direct['result']['routes'][0]
        r_s =  int(result['duration'])
        r_h = r_s/3600
        r_m = (r_s%3600)/60
        ws.Cells(row,9).Value = u'%s小时%s分'%(r_h, r_m)
        distance = float(result['distance']) / 1000
        ws.Cells(row,10).Value = u'%.1f公里'%distance
        ws.Cells(row,18).Value = 18
        
        tactics = 10 # 不走高速
        direct = Direction(origin, dest, tactics)
        result = direct['result']['routes'][0]
        r_s =  int(result['duration'])
        r_h = r_s/3600
        r_m = (r_s%3600)/60
        ws.Cells(row,13).Value = u'%s小时%s分'%(r_h, r_m)
        distance = float(result['distance']) / 1000
        ws.Cells(row,14).Value = u'%.1f公里'%distance
        ws.Cells(row,19).Value = 19

        break
