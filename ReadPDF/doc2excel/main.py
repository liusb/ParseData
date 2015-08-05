# -*- coding: utf-8 -*-

import wx
import sqlite3
from threading import Thread
import os
import win32com.client
import pythoncom
import sys
import traceback
import copy
import re


def except_hook(etype, value, tb):
    message = u'异常信息:\n'
    message += ''.join(traceback.format_exception(etype, value, tb))
    wx.LogMessage(message)
    print message
    wx.MessageBox(u'程序出现异常', caption=u'错误信息', style=wx.ICON_ERROR)


class SqlData():
    def __init__(self, db_name=None):
        if db_name is None:
            db_name = 'config.db'
        self.conn = sqlite3.connect(db_name)
        self.conn.text_factory = str
        self.cur = self.conn.cursor()

    def __del__(self):
        self.cur.close()
        self.conn.close()

    def get_config(self):
        self.cur.execute('select name, value from config')
        rows = self.cur.fetchall()
        config = {}
        for row in rows:
            config[row[0]] = row[1]
        return config

    def save_config(self, config):
        try:
            for key, value in config.items():
                temp = ("update config set value = '%s' where name = '%s'" % (value, key)).encode('gbk')
                self.cur.execute(temp)
            self.conn.commit()
        except sqlite3.Error:
            self.conn.rollback()
            return False
        return True

    def get_keys(self):
        self.cur.execute('select id, key from keys order by id')
        rows = self.cur.fetchall()
        keys = []
        for row in rows:
            key = {
                'id': row[0],
                'key': row[1],
            }
            keys.append(key)
        return keys

    def save_keys(self, config):
        try:
            keys = config['keys'].split('|')
            for i in range(1, len(keys)+1):
                temp = ("insert or replace into keys(id, key) values(%d, '%s')" % (i, keys[i-1])).encode('gbk')
                self.cur.execute(temp)
            self.cur.execute('delete from keys where id > %d' % len(keys))
            del config['keys']
            for key, value in config.items():
                temp = ("update config set value = '%s' where name = '%s'" % (value, key)).encode('gbk')
                self.cur.execute(temp)
            self.conn.commit()
        except sqlite3.Error:
            self.conn.rollback()
            return False
        return True


class Word:

    def __init__(self, visible, keys, config):
        self.word = win32com.client.Dispatch('Word.Application')
        self.word.Visible = visible
        self.word.ScreenUpdating = visible
        self.word.DisplayAlerts = False
        self.doc = None
        self.keys = keys
        self.filter_row = config['filter_row'].decode('gbk').split('|')
        self.filter_col = config['filter_col'].decode('gbk').split('|')

    def __del__(self):
        try:
            if self.word.Documents.Count == 0:
                self.word.Quit()
        except:
            pass

    def open(self, name):
        self.doc = self.word.Documents.Open(name, ReadOnly=True)

    def close(self):
        if self.doc is not None:
            self.doc.Close()
            self.doc = None

    def doc_pages(self):
        last_range = self.doc.Range(self.doc.Content.End-1, self.doc.Content.End)
        return last_range.Information(3)  # 页码

    def find_next_paragraph(self, range_obj):
        count = 1
        while True:
            next_range = range_obj.Next(Unit=4, Count=count)
            # if len(next_one.Text.replace(' ', '').replace('\r', '').replace('\a', '')
            #     .replace('\t', '').replace('\x0c', '')) > 0:
            if next_range.ComputeStatistics(Statistic=6) > 0:
                return next_range
            count += 1

    def find_previous_paragraph(self, range_obj):
        count = 1
        while True:
            previous_range = range_obj.Previous(Unit=4, Count=count)
            # if len(previous_range.Text.replace(' ', '').replace('\r', '')
            #     .replace('\a', '').replace('\t', '').replace('\x0c', '')) > 0:
            if previous_range.ComputeStatistics(Statistic=6) > 0:
                return previous_range
            count += 1

    def row_yes_parse(self, table):
        result = []
        col_ct = table.Columns.Count
        row_ct = table.Rows.Count
        row_i = 1
        while row_i <= row_ct:
            row_content = []
            col_i = 1
            while col_i <= col_ct:
                try:
                    text = table.Cell(Row=row_i, Column=col_i).Range.Text
                    text = text.replace(' ', '').replace('\r', '').replace('\a', '').replace('\t', '')
                    if len(text) > 0:
                        row_content.append(text)
                    col_i += 1
                except:
                    col_i += 1
            if len(row_content) > 0:
                result.append(row_content)
            row_i += 1
        return result

    def col_yes_parse(self, table):
        result = []
        col_i = 1
        col_ct = table.Columns.Count
        row_ct = table.Rows.Count
        while col_i <= col_ct:
            col_content = []
            row_i = 1
            while row_i <= row_ct:
                try:
                    text = table.Cell(Row=row_i, Column=col_i).Range.Text
                    text = text.replace(' ', '').replace('\r', '').replace('\a', '').replace('\t', '')
                    if len(text) > 0:
                        col_content.append(text)
                    row_i += 1
                except:
                    row_i += 1
            if len(col_content) > 0:
                result.append(col_content)
            col_i += 1
        r_result = []
        table_row = 0
        for col in result:
            if table_row < len(col):
                table_row = len(col)
        table_col = col_ct
        for i in range(0, table_row):
            row_content = []
            for j in range(0, table_col):
                try:
                    row_content.append(result[j][i])
                except IndexError:
                    break
            if len(row_content) > 0:
                r_result.append(row_content)
        return r_result

    def text_split(self, text):
        result = []
        temp = text.split()
        for item in temp:
            result.extend(item.split('\r'))
        return result

    def uniform_parse(self, table):
        result = []
        col_ct = table.Columns.Count
        row_ct = table.Rows.Count
        if col_ct == 3 and row_ct >= 2:
            row_content = []
            for i in range(1, col_ct+1):
                text = table.Cell(1, i).Range.Text.replace('\r', '').replace('\a', '')
                row_content.append(text)
            result.append(row_content)
            row2_1 = self.text_split(table.Cell(2, 1).Range.Text.replace('\a', ''))
            row2_2 = self.text_split(table.Cell(2, 2).Range.Text.replace('\a', ''))
            row2_3 = self.text_split(table.Cell(2, 3).Range.Text.replace('\a', ''))
            if len(row2_1) == len(row2_2) and len(row2_2) == len(row2_3):
                for i in range(0, len(row2_1)):
                    result.append([row2_1[i], row2_2[i], row2_3[i]])
            elif len(row2_1)-1 == len(row2_2) and len(row2_2) == len(row2_3):
                max_index = 0
                min_index = 0
                for i in range(1, len(row2_1)):
                    if len(row2_1[i]) > len(row2_1[max_index]):
                        max_index = i
                    elif len(row2_1[i]) <= len(row2_1[min_index]):
                        min_index = i
                if max_index == min_index - 1:
                    row2_1[max_index] = row2_1[max_index] + row2_1[min_index]
                    row2_1.pop(min_index)
                for i in range(0, len(row2_3)):
                    result.append([row2_1[i], row2_2[i], row2_3[i]])
            else:
                row_content = list()
                for i in range(1, col_ct+1):
                    text = table.Cell(2, i).Range.Text.replace(' ', '').replace('\r', '')\
                        .replace('\a', '').replace('\t', '')
                    row_content.append(text)
                result.append(row_content)
            for j in range(3, row_ct+1):
                row_content = list()
                for i in range(1, col_ct+1):
                    text = table.Cell(j, i).Range.Text.replace(' ', '').replace('\r', '')\
                        .replace('\a', '').replace('\t', '')
                    row_content.append(text)
                result.append(row_content)
            return result
        else:
            return self.row_yes_parse(table)

    def parse_table(self, table, may_col=0):
        col_ct = table.Columns.Count
        row_ct = table.Rows.Count
        if col_ct == 2:
            wx.LogMessage(u'WORD表格行数:%d, 列数:%d，按列继续处理' % (row_ct, col_ct))
            result = self.col_yes_parse(table)
        elif table.Uniform is True:
            wx.LogMessage(u'WORD表格是个规范表格 行数:%d, 列数:%d，按行继续处理' % (row_ct, col_ct))
            result = self.uniform_parse(table)
        elif may_col != 0 and may_col == col_ct:
            wx.LogMessage(u'WORD表格行数:%d, 列数:%d，按列继续处理' % (row_ct, col_ct))
            result = self.col_yes_parse(table)
        elif row_ct <= 7:
            wx.LogMessage(u'WORD表格行数:%d, 列数:%d，按行继续处理' % (row_ct, col_ct))
            result = self.row_yes_parse(table)
        else:
            wx.LogMessage(u'WORD表格行数:%d, 列数:%d，按列继续处理' % (row_ct, col_ct))
            result = self.col_yes_parse(table)
        return result

    def is_other_table(self, find_table):
        for item in find_table[0]:
            for key in self.filter_row:
                if key in item:
                    return True
        for item in find_table:
            for key in self.filter_col:
                if key in item[0]:
                    return True
        return False

    def parse(self):
        result = {}
        find_any_key = 0
        remarks = []
        for key in self.keys:
            wx.LogMessage(u'查找关键【%s】字' % key['key'].decode('gbk'))
            cur_pos = self.doc.Content.Start  # 每次从头开始
            while True:
                self.doc.Activate()
                selection = self.word.Selection
                selection.SetRange(cur_pos, cur_pos)  # 设置光标
                find = selection.Find
                find.IgnoreSpace = True
                find.Forward = True
                if find.Execute(FindText=key['key']) is not True:
                    wx.LogMessage(u'查找到文件尾，查找下一个关键字')
                    break
                find_any_key += 1
                find_s = selection.Start
                find_e = selection.End
                assert find_s > cur_pos
                cur_pos = find_e-2
                key_range = self.doc.Range(find_s, find_e)
                key_range.Expand(Unit=4)
                key_range_s = key_range.Start
                key_range_e = key_range.End
                if key_range_e - key_range_s > 80:
                    key_range_text = key_range.Text
                    if u'元' not in key_range_text and u'占' not in key_range_text:
                        wx.LogMessage(u'关键字所在段落太长，排除关键字')
                        continue
                key_range_page = key_range.Information(3)
                wx.LogMessage(u'在%d页查找到关键字' % key_range_page)
                if key_range.Information(12) is True:  # wdWithInTable = 12
                    wx.LogMessage(u'关键字在表格中，排除关键字')
                    continue
                if (u'人民币' not in key_range.Text and len(key_range.Text.split('\t')) > 2) or\
                        (u'合计销售' in key_range.Text and len(key_range.Text.split('\t')) == 2):
                    wx.LogMessage(u'关键字在文字表格中，排除关键字')
                    continue

                # 向后读到一个表为止
                find_try_next = True  # 是否还要继续处理
                find_go_next = False  # 直接开始查找下一个
                tow_table_flag = False  # 用来标记是否为分页表格，在查找表格后一段的时候需要使用
                next_table_range = None
                while True:
                    table_range = key_range.Next(Unit=15)  # 15 => table
                    if table_range is None:   # 查找到文档尾
                        table_s = self.doc.Content.End  # 设置表头为文件尾部，方便分析文字
                        break
                    table_s = table_range.Start  # 不能移到后面去，分析文字的时候需要用
                    table_e = table_range.End
                    if key['id'] == 5:  # 对于可能是总表的情况看看是否要做特殊处理
                        text_row_range = self.find_next_paragraph(key_range)
                        text_row_range_text = text_row_range.Text.replace('\r', '').replace(' ', '')
                        row_content = text_row_range_text.split('\t')
                        if len(row_content) == 2 and u'合计销售金额' in row_content[0] and u',' in row_content[1]:
                            text_row_range = self.find_next_paragraph(text_row_range)
                            text_row_range_text = text_row_range.Text.replace('\r', '').replace(' ', '')
                            row_content = text_row_range_text.split('\t')
                            if len(row_content) == 2 and u'占年度' in row_content[0] and u'%' in row_content[1]:
                                break
                    table_t = table_range.Tables(1)
                    if table_s - find_e > 400:
                        wx.LogMessage(u'关键字和表格间大于400字，排除表格')
                        break
                    table_page = table_range.Information(3)  # 页码
                    table_index = (table_page, table_s)
                    if table_index in result.keys():
                        table_item = result[table_index]
                        if table_item['key']['find_s'] < find_s:  # 关键字离表格更近
                            table_item['key']['id'] = key['id']
                            table_item['key']['key'] = key['key']
                            table_item['key']['find_s'] = find_s
                        find_go_next = True
                        break
                    find_table = self.parse_table(table_t)
                    next_table_range = table_range.Next(Unit=15)
                    next_table_page = next_table_range.Information(3)  # 页码
                    if table_page + 1 == next_table_page:
                        next_table_s = next_table_range.Start
                        table_between_ct = self.doc.Range(table_e, next_table_s).ComputeStatistics(Statistic=6)
                        if table_between_ct == 0:
                            tow_table_flag = True
                            wx.LogMessage(u'表格分页，分别在%d，%d两页' % (table_page, next_table_page))
                            next_table_t = next_table_range.Tables(1)
                            may_col = len(find_table[-1]) if len(find_table) > 0 else 0
                            find_next_table = self.parse_table(next_table_t, may_col)
                            if len(find_next_table) > 1 and len(find_table) > 0 \
                                    and find_table[0][0] == find_next_table[0][0]:
                                find_next_table.pop(0)
                            elif len(find_next_table) > 0 and len(find_next_table[0]) == 3 and len(find_table) > 0 \
                                    and len(find_next_table[0][1]) == 0 and len(find_next_table[0][2]) == 0:
                                find_table[-1][0] += find_next_table[0][0]
                                find_next_table.pop(0)
                            find_table.extend(find_next_table)
                    if len(find_table) == 0:  # 表不合规则
                        break
                    if len(find_table) > 10:
                        wx.LogMessage(u'表格行数超过10行，排除')
                        break
                    if self.is_other_table(find_table) is True:
                        wx.LogMessage(u'取到其他表格，排除')
                        break

                    # 对表格做特殊处理
                    if len(find_table) >= 2 and len(find_table[-1]) == 3 and len(find_table[-2]) == 4:
                        find_table[-1].insert(1, u'-')
                    if len(find_table[-1]) == 1 and find_table[-1][0] == u'（%）':
                        find_table.pop()
                    if len(find_table) >= 2 and len(find_table[0]) == 1 and len(find_table[1]) == 2:
                        if u'比例' in find_table[0][0] and u'客户名称' in find_table[1]:
                            find_table[1].append(find_table[0][0])
                            find_table.pop(0)
                    if len(find_table) == 3 and len(find_table[0]) == 3 \
                            and len(find_table[1]) == 2 and len(find_table[2]) == 3:
                        if u',' in find_table[1][0]:
                            find_table[1].insert(0, u'-')

                    if len(find_table) == 8 and len(find_table[1]) == 3 and len(find_table[1][0]) == 0 \
                            and u'人民币' in find_table[1][1] and u'%' == find_table[1][2]:
                        find_table[0][1] = find_table[0][1] + u'(' + find_table[1][1] + u')'
                        find_table[0][2] = find_table[0][2] + u'(' + find_table[1][2] + u')'
                        find_table.pop(1)

                    if len(find_table) == 1 and u'客户名称' in find_table[0]:
                        wx.LogWarning(u'找到只有标题的空表，排除')
                        find_go_next = True
                        break
                    elif len(find_table) == 2 and u'客户名称' in find_table[0] and u'/' == find_table[1][0]:
                        wx.LogWarning(u'找到只有标题的空表，排除')
                        find_go_next = True
                        break
                    else:
                        find_try_next = False  # 找到了表格不用再尝试
                        break
                # 尝试处理表格完毕
                if find_go_next is True:
                    continue
                # 尝试处理文字表格
                while True:
                    if find_try_next is not True:
                        break
                    if table_s - key_range_e < 20:
                        # 表格和关键字间的字太少
                        find_go_next = True
                        break
                    # 做一个尝试，有没有可能是个表格
                    after_key_200_e = key_range_e + 200 if table_s - key_range_e > 200 else table_s
                    after_key_200 = self.doc.Range(key_range_e, after_key_200_e).Text\
                        .replace('\r', '').replace(' ', '').replace('\t', '').replace('\a', '')
                    remarks_add_flag = False
                    if u'客户名称' in after_key_200:
                        remarks.append(u'[%s]在%d页，复查' % (key['key'].decode('gbk'), key_range_page))
                        remarks_add_flag = True
                    # 下面开始解析文字
                    find_table = []
                    text_row_range = self.find_next_paragraph(key_range)
                    text_row_range_text = text_row_range.Text.replace('\r', '').replace(' ', '')
                    if u'√适用' in text_row_range_text:
                        text_row_range = self.find_next_paragraph(text_row_range)
                        text_row_range_text = text_row_range.Text.replace('\r', '').replace(' ', '')
                    row_content = text_row_range_text.split('\t')
                    table_col = len(row_content)
                    if table_col < 2:
                        wx.LogMessage(u'关键字后的文本不能正确解析')
                        break
                    if u'期' == row_content[0] and u'间' == row_content[1]:
                        row_content.pop(0)
                        row_content[0] = u'期间'
                        table_col = len(row_content)
                    find_table.append(row_content)
                    table_page = text_row_range.Information(3)  # 页码
                    table_s = text_row_range.Start
                    table_e = text_row_range.End
                    table_index = (table_page, table_s)
                    while True:
                        text_row_range = self.find_next_paragraph(text_row_range)
                        row_content = text_row_range.Text.replace('\r', '').replace(' ', '')\
                            .replace('\t\t', '\t').split('\t')
                        if len(row_content) == 1 and u'%' in row_content[0] and len(row_content[0]) < 6:
                            continue
                        if len(row_content) % table_col == 0:
                            for i in range(0, len(row_content), table_col):
                                find_table.append(row_content[i:i+table_col])
                            table_e = text_row_range.End
                        elif len(row_content) != table_col:
                            break
                    table_range = self.doc.Range(table_s, table_e)
                    if len(find_table) < 2:
                        break
                    # 对 合计 分开的做特殊处理
                    if u'合' in find_table[-2][-1]:
                        find_table[-2][-1] = find_table[-2][-1].replace(u'合', '')
                        find_table[-1][0] = u'合计'
                    if self.is_other_table(find_table) is True:
                        wx.LogMessage(u'取到其他表格，排除')
                        break
                    # 成功走到这里说明解析文字成功了，删除备注
                    if remarks_add_flag is True:
                        remarks.pop(-1)
                    # 找到了 不要再进行其他解析
                    find_try_next = False

                if find_go_next is True:
                    continue
                while True:
                    if find_try_next is not True:
                        break
                    # 第一个表格是以文字的形式描述的，对此做特殊处理
                    table_range = self.find_next_paragraph(key_range)
                    after_key_text = table_range.Text.replace(' ', '').replace('\r', '')\
                        .replace('\a', '').replace(u'％', u'%')
                    if u'元' in after_key_text and u'%' not in after_key_text:  # 占比在下面一段的情况
                        after_table_range = self.find_next_paragraph(table_range)
                        after_table_range_text = after_table_range.Text.replace(' ', '').replace('\r', '')\
                            .replace('\a', '').replace(u'％', u'%')
                        percent_flag_pos = after_table_range_text.find(u'%')
                        if percent_flag_pos != -1:
                            after_key_text += after_table_range_text[0:percent_flag_pos+1]
                    else:
                        key_range_text = key_range.Text.replace(' ', '').replace('\r', '')\
                            .replace('\a', '').replace(u'％', u'%')
                        if u'%' in key_range_text and u'占' in key_range_text:
                            table_range = key_range
                            table_page = table_range.Information(3)  # 页码
                            table_index = (table_page, table_range.Start)
                            if table_index in result.keys():
                                find_go_next = True
                                break
                            after_key_text = key_range_text  # 关键字本身为描述字段
                    # if len(after_key_text) > 100:
                    #    wx.LogMessage(u'关键字后一段文字超过100，排除')
                    #    break
                    if u'占' not in after_key_text:
                        break
                    if u'客户' not in after_key_text and u'主要产品经销商' not in after_key_text:
                        break
                    if u'营业收入' not in after_key_text and u'销售总额' not in after_key_text \
                            and u'销售收入' not in after_key_text and u'销售额' not in after_key_text \
                            and u'销售金额' not in after_key_text:
                        break
                    money_text_pos = after_key_text.find(u'元', 1)  # 从下标1开始查找,后面直接判断前面那个数字
                    percent_text_pos = after_key_text.find(u'%', money_text_pos+1)
                    if money_text_pos == -1 and percent_text_pos == -1:
                        break
                    if money_text_pos != -1:
                        if u'万' == after_key_text[money_text_pos-1]:
                            money_base = 10000
                            money_text_pos -= 1
                        elif u'亿' == after_key_text[money_text_pos-1]:
                            money_base = 100000000
                            money_text_pos -= 1
                        else:
                            money_base = 1
                        number_str = list()
                        for i in range(money_text_pos-1, 0, -1):
                            if (u'0' <= after_key_text[i] <= u'9') or u'.' == after_key_text[i]:
                                number_str.insert(0, after_key_text[i])
                            elif u',' == after_key_text[i]:
                                continue
                            else:
                                break
                        money_number = float(''.join(number_str)) * money_base
                        money_number_str = '{:.2f}'.format(money_number)
                    else:
                        money_number_str = ''
                    if percent_text_pos != -1:
                        percent_str = ['%', ]
                        for i in range(percent_text_pos-1, percent_text_pos-6, -1):
                            if (u'0' <= after_key_text[i] <= u'9') or u'.' == after_key_text[i]:
                                percent_str.insert(0, after_key_text[i])
                        percent_text = ''.join(percent_str)
                    else:
                        percent_text = ''
                    find_table = [[after_key_text, u'前五名客户合计销售金额', money_number_str],
                                  [u'纯文字描述', u'前五名客户合计销售金额占年度销售总额比例', percent_text]]
                    table_page = table_range.Information(3)  # 页码
                    table_index = (table_page, table_range.Start)
                    find_try_next = False  # 找到了不要再尝试了

                if find_go_next is True:
                    continue
                if find_try_next is True:
                    wx.LogMessage(u'未能成功解析，继续查找')
                    continue

                table_col = 0
                for row in find_table:
                    if table_col < len(row):
                        table_col = len(row)
                table_row = len(find_table)
                wx.LogMessage(u'在第%d页找到符合的表格,行:%d 列:%d' % (table_page, table_row, table_col))
                before1 = self.find_previous_paragraph(table_range)
                before2 = self.find_previous_paragraph(before1)
                if tow_table_flag is True:
                    after1 = self.find_next_paragraph(next_table_range)
                else:
                    after1 = self.find_next_paragraph(table_range)
                after2 = self.find_next_paragraph(after1)
                table_item = dict()
                table_item['key'] = {}
                table_item['key']['id'] = key['id']
                table_item['key']['key'] = key['key']
                table_item['key']['find_s'] = find_s
                table_item['table'] = {}
                table_item['table']['page'] = table_page
                table_item['table']['before2'] = before2.Text.replace(' ', '').replace('\r', '').replace('\a', '')
                table_item['table']['before1'] = before1.Text.replace(' ', '').replace('\r', '').replace('\a', '')
                table_item['table']['after1'] = after1.Text.replace(' ', '').replace('\r', '').replace('\a', '')
                table_item['table']['after2'] = after2.Text.replace(' ', '').replace('\r', '').replace('\a', '')
                if table_row == 8 and len(find_table[0]) == 4 and len(find_table[1]) == 3 \
                        and u'前' in find_table[0][0] and u'%' in find_table[0][2] and u'客户' in find_table[1][0]:
                    # 对于表1表2在一起的表格做特殊处理
                    find_table_1 = [[find_table[0][0], find_table[0][1]], [find_table[0][2], find_table[0][3]]]
                    table_item['table']['row'] = 2
                    table_item['table']['col'] = 2
                    table_item['table']['content'] = find_table_1
                    result[table_index] = table_item
                    find_table_2 = []
                    for i in range(1, 8):
                        find_table_2.append(find_table[i])
                    table_index_2 = (table_index[0], table_index[1]+2)
                    table_item_2 = copy.deepcopy(table_item)
                    table_item_2['table']['row'] = 7
                    table_item_2['table']['col'] = 3
                    table_item_2['table']['content'] = find_table_2
                    result[table_index_2] = table_item_2
                else:   # 正常情况下
                    table_item['table']['row'] = table_row
                    table_item['table']['col'] = table_col
                    table_item['table']['content'] = find_table
                    result[table_index] = table_item

        return result, find_any_key, remarks


class Excel():
    def __init__(self):
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel.Visible = True
        self.wkbk = None
        self.wksht = None
        self.first_row = 2
        self.current_begin = 1
        self.current_end = 1
        self.config = SqlData().get_config()
        self.word = None

    def init_word(self, visible):
        keys = SqlData().get_keys()
        self.word = Word(visible, keys, self.config)

    def open(self):
        self.wkbk = self.excel.Workbooks.Open(self.config['excel_file'])
        self.wksht = self.wkbk.Worksheets(self.config['sheet'])

    def save(self):
        self.wkbk.Save()

    def clear(self):
        begin_row = self.first_row
        end_row = self.first_row
        wx.LogMessage(u'查找Excel的末尾')
        while True:
            if self.wksht.Cells(end_row, 1).Value is None:
                break  # 最后一行为空结束运行
            end_row += 1
        if end_row > begin_row:
            end_row -= 1
            wx.LogMessage(u'Excel末尾行为%d' % end_row)
            self.wksht.Range('A%d:A%d' % (begin_row, end_row)).EntireRow.Delete()
            self.save()
            wx.LogMessage(u'成功清除Excel中的数据')
        else:
            wx.LogMessage(u'Excel中没有数据')

    def init(self):
        begin_row = self.first_row
        count = 1
        config = SqlData().get_config()
        dirs = os.listdir(config['pdf_dir'])
        for item in dirs:
            if item.endswith('.pdf'):
                wx.LogMessage(item)
                end_row = begin_row + 6
                self.wksht.Range('A%d:A%d' % (begin_row, end_row)).Value = count
                for row in range(begin_row, end_row+1):
                    link = self.wksht.Hyperlinks.Add(Anchor=self.wksht.Range("B%d" % row),
                                                     Address=os.path.join(config['pdf_dir'], item),
                                                     TextToDisplay=item)
                    if link.TextToDisplay != item:
                        link.TextToDisplay = item
                    doc_name = item[:-3]+'doc'
                    link = self.wksht.Hyperlinks.Add(Anchor=self.wksht.Range("C%d" % row),
                                                     Address=os.path.join(config['word_dir'], doc_name),
                                                     TextToDisplay=doc_name)
                    if link.TextToDisplay != doc_name:
                        link.TextToDisplay = doc_name
                self.wksht.Range('E%d:E%d' % (begin_row, end_row)).Value = item[0:4]
                self.wksht.Range('F%d:F%d' % (begin_row, end_row)).Value = item[5:11]
                begin_row = end_row + 1
                count += 1
        wx.LogMessage(u'总共找到%d个文件' % (count-1))
        self.save()
        wx.LogMessage(u'成功将找到的文件初始化到Excel中')

    def next_doc(self):
        self.current_begin = self.current_end + 1
        first = self.wksht.Cells(self.current_begin, 1).Value
        if first is None:
            wx.LogMessage(u'未找到更多的记录，处理结束')
            return False
        row = self.current_begin
        while True:
            row += 1
            if self.wksht.Cells(row, 1).Value != first:
                break
        self.current_end = row - 1
        return True

    def is_processed(self):
        if self.wksht.Cells(self.current_begin, 4).Value is not None\
                and self.wksht.Cells(self.current_begin, 4).Value == u'是':
            return True
        return False

    def adjust_row(self, row_ct):
        begin_row = self.current_begin
        end_row = self.current_end
        if end_row - begin_row + 1 == row_ct:  # 正好
            wx.LogMessage(u'调整表格：行数正好符合')
        elif end_row - begin_row + 1 > row_ct:  # 表格有多
            self.wksht.Range('A%d:F%d' % (begin_row+row_ct, end_row)).EntireRow.Delete()
            self.current_end = begin_row + row_ct - 1
            wx.LogMessage(u'调整表格：行数有多，删除%d行' % (end_row-self.current_end))
        else:  # 还有空余行，但不够用
            self.current_end = begin_row + row_ct - 1
            self.wksht.Range('A%d:F%d' % (end_row+1, self.current_end)).EntireRow.Insert()
            self.wksht.Range('A%d:F%d' % (begin_row, begin_row)).Copy()
            self.wksht.Paste(self.wksht.Range('A%d:F%d' % (end_row+1, self.current_end)))
            wx.LogMessage(u'调整表格：行数不够，添加%d行' % (self.current_end-end_row))

    def process_doc(self):
        doc_name = self.wksht.Cells(self.current_begin, 3).Hyperlinks(1).TextToDisplay
        doc_id = self.wksht.Cells(self.current_begin, 1).Value
        wx.LogMessage(u'====== 开始处理%s，编号：%d ======' % (doc_name, doc_id))
        self.word.open(os.path.join(self.config['word_dir'].decode('gbk'), doc_name))
        result, find_any_key, remarks = self.word.parse()
        wx.LogMessage(u'关键字查找完毕')
        doc_pages = self.word.doc_pages()
        self.write_data(result, doc_pages)
        if len(result) == 0:
            self.wksht.Range('D%d:D%d' % (self.current_begin, self.current_end)).Value = u'否'
            self.wksht.Cells(self.current_begin, 22).Value = u'需手工整理'
        for remark in remarks:
            self.write_remark(remark)
        self.write_find_time(find_any_key)
        wx.LogMessage(u'处理%s完毕' % doc_name)
        self.wkbk.Save()
        self.word.close()

    def write_remark(self, remark, col=22):
        write_row = self.current_begin-1
        for row in range(self.current_begin, self.current_end+1):
            cell_value = self.wksht.Cells(row, col).Value
            if cell_value is None:
                write_row = row
                break
        if write_row == self.current_begin-1:
            self.write_remark(remark, col+1)
        else:
            self.wksht.Cells(write_row, col).Value = remark

    def write_find_time(self, find_time, col=23):
        remark = u'总共找到%d次关键字' % find_time
        write_row = self.current_begin-1
        for row in range(self.current_begin, self.current_end+1):
            cell_value = self.wksht.Cells(row, col).Value
            if cell_value is None or (isinstance(cell_value, str) and u'次关键字' in cell_value):
                write_row = row
                break
        if write_row == self.current_begin-1:
            self.write_remark(remark, col+1)
        else:
            self.wksht.Cells(write_row, col).Value = remark

    def write_data(self, tables, doc_pages):
        row_ct = 0
        for table in tables.values():
            row_ct += table['table']['row']
        if row_ct == 0:
            wx.LogMessage(u'未找到数据')
            return
        wx.LogMessage(u'共需写入%d行数据' % row_ct)

        sorted_key = sorted(tables.keys())
        first_table = tables[sorted_key[0]]
        if first_table['table']['row'] == 2 and u'前' in first_table['table']['after1'] \
                and u'√不适用' in first_table['table']['after2']:
            first_table_col = first_table['table']['col']
            if first_table_col < 5:
                self.wksht.Cells(self.current_begin, 22).Value = u'第一个明细表不适用'
            else:
                self.wksht.Cells(self.current_begin, 17+first_table_col).Value = u'第一个明细表不适用'
        self.adjust_row(row_ct)
        count = 1
        end_row = self.current_begin - 1
        for key in sorted_key:
            table = tables[key]
            table_row = table['table']['row']
            table_col = table['table']['col']
            begin_row = end_row + 1
            end_row = end_row + table_row
            self.wksht.Range('G%d:G%d' % (begin_row, end_row)).Value = table['key']['id']
            self.wksht.Range('H%d:H%d' % (begin_row, end_row)).Value = table['key']['key']
            self.wksht.Range('J%d:J%d' % (begin_row, end_row)).Value = table['table']['row']
            self.wksht.Range('K%d:K%d' % (begin_row, end_row)).Value = table['table']['before2']
            self.wksht.Range('L%d:L%d' % (begin_row, end_row)).Value = table['table']['before1']
            self.wksht.Range('M%d:M%d' % (begin_row, end_row)).Value = table['table']['after1']
            self.wksht.Range('N%d:N%d' % (begin_row, end_row)).Value = table['table']['after2']
            self.wksht.Range('O%d:O%d' % (begin_row, end_row)).Value = u'表%d' % count
            self.wksht.Range('P%d:P%d' % (begin_row, end_row)).Value = table['table']['page']

            if table_row == 2 and table_col == 3 and u'纯文字描述' == table['table']['content'][1][0]:
                table['table']['content'][1][0] = ''
                self.wksht.Cells(begin_row, 22).Value = u'纯文字描述'

            begin_col = 17 if table_col > 4 else (21 - table_col)
            row_i = begin_row
            for row in table['table']['content']:
                col_i = begin_col
                for item in row:
                    self.wksht.Cells(row_i, col_i).Value = item
                    col_i += 1
                row_i += 1
            if count == 2 and first_table['table']['row'] <= 2:  # 表1为总表 表2是否在文档的20%后
                table_page = table['table']['page']
                if table_page > doc_pages * 0.2:
                    self.wksht.Cells(begin_row, 22).Value = u'表2在word的20%之后，需复查'
            count += 1
        self.wksht.Range('D%d:D%d' % (self.current_begin, self.current_end)).Value = u'是'
        self.wksht.Range('I%d:I%d' % (self.current_begin, self.current_end)).Value = len(tables)


class WorkerThread(Thread):

    def __init__(self, parent):
        Thread.__init__(self)
        self.running = False
        self.parent = parent

    def run(self):
        wx.LogMessage(u'工作线程已经启动')
        try:
            pythoncom.CoInitializeEx(0)
            self.running = True
            excel = Excel()
            excel.open()
            excel.init_word(self.parent.word_btn.GetValue())
            while self.running:
                if excel.next_doc() is False:
                    wx.LogMessage(u'已经到达Excel文件尾')
                    break
                if excel.is_processed() is True:
                    continue
                excel.process_doc()
            wx.LogMessage(u'工作线程已经结束')
        except:
            etype, value, tb = sys.exc_info()
            message = u'工作线程出现异常:\n'
            print message, traceback.print_exc()
            message += ''.join(traceback.format_exception(etype, value, tb))
            wx.LogMessage(message)
            wx.MessageBox(u'程序出现异常', caption=u'错误信息', style=wx.ICON_ERROR)
        finally:
            wx.CallAfter(self.parent.start_btn.Enable)
            wx.CallAfter(self.parent.stop_btn.Disable)

    def stop(self):
        self.running = False


class RunPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.start_btn = wx.Button(self, label=u'开始提取')
        self.stop_btn = wx.Button(self, label=u'暂停提取')
        self.clear_btn = wx.Button(self, label=u'清除日志')
        self.excel_init_btn = wx.Button(self, label=u'初始化EXCEL')
        self.excel_clear_btn = wx.Button(self, label=u'清空EXCEl')
        self.word_btn = wx.ToggleButton(self, label=u'显示Word')

        self.Bind(wx.EVT_TOGGLEBUTTON, self.word_click, self.word_btn)
        self.Bind(wx.EVT_BUTTON, self.start_click, self.start_btn)
        self.Bind(wx.EVT_BUTTON, self.stop_click, self.stop_btn)
        self.Bind(wx.EVT_BUTTON, self.clear_click, self.clear_btn)
        self.Bind(wx.EVT_BUTTON, self.excel_init_click, self.excel_init_btn)
        self.Bind(wx.EVT_BUTTON, self.excel_clear_click, self.excel_clear_btn)
        self.stop_btn.Disable()

        self.text_ctrl = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        self.text_ctrl.SetEditable(False)
        self.SetBackgroundColour('White')

        top_sizer = wx.BoxSizer(wx.VERTICAL)
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(self.word_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.start_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.stop_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.clear_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.excel_init_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.excel_clear_btn, 0, wx.ALL, 5)
        top_sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)
        top_sizer.Add(self.text_ctrl, 1, wx.EXPAND | wx.ALL, 10)
        self.SetSizerAndFit(top_sizer)

        self.thread = None

    def write(self, s):
        self.text_ctrl.AppendText(s)

    def clear_click(self, event):
        self.text_ctrl.Clear()

    def word_click(self, event):
        if self.word_btn.GetValue():
            wx.LogMessage(u'在运行时显示Word')
        else:
            wx.LogMessage(u'在运行时不会显示Word')
        
    def start_click(self, event):
        self.stop_btn.Enable()
        self.start_btn.Disable()
        self.thread = WorkerThread(self)
        self.thread.start()

    def stop_click(self, event):
        self.thread.stop()
        wx.LogMessage(u'程序已收到暂停请求，处理完当前文件会暂停工作！请稍等！！！')

    def excel_init_click(self, event):
        wx.LogMessage(u'初始化Excel开始，请稍等！')
        self.excel_init_btn.Disable()
        try:
            excel = Excel()
            excel.open()
            excel.init()
        finally:
            self.excel_init_btn.Enable()
        wx.LogMessage(u'初始化Excel结束！')

    def excel_clear_click(self, event):
        dlg = wx.MessageDialog(None, u"是否确定清空Excel?", u"确认信息", wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_NO:
            return
        wx.LogMessage(u'开始准备清空Excel，请稍等！')
        self.excel_clear_btn.Disable()
        try:
            excel = Excel()
            excel.open()
            excel.clear()
        finally:
            self.excel_clear_btn.Enable()
        wx.LogMessage(u'清空Excel完毕！')


class ConfigPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        config = SqlData().get_config()

        self.excel_label = wx.StaticText(self, label=u'EXCEL文件')
        self.excel_file = wx.TextCtrl(self, value=config['excel_file'])
        self.excel_btn = wx.Button(self, label=u'选择文件')
        self.Bind(wx.EVT_BUTTON, self.excel_click, self.excel_btn)

        self.pdf_label = wx.StaticText(self, label=u'PDF目录')
        self.pdf_dir = wx.TextCtrl(self, value=config['pdf_dir'])
        self.pdf_btn = wx.Button(self, label=u'选择目录')
        self.Bind(wx.EVT_BUTTON, self.pdf_click, self.pdf_btn)

        self.word_label = wx.StaticText(self, label=u'WORD目录')
        self.word_dir = wx.TextCtrl(self, value=config['word_dir'])
        self.word_btn = wx.Button(self, label=u'选择目录')
        self.Bind(wx.EVT_BUTTON, self.word_click, self.word_btn)

        self.sheet_label = wx.StaticText(self, label=u'当前Sheet名字')
        self.sheet = wx.TextCtrl(self, value=config['sheet'])

        self.save_config_btn = wx.Button(self, label=u'保存设置')
        self.Bind(wx.EVT_BUTTON, self.save_config_click, self.save_config_btn)

        self.top_sizer = wx.BoxSizer(wx.VERTICAL)
        self.path_sizer = wx.GridBagSizer(hgap=10, vgap=10)
        self.path_sizer.Add(self.excel_label, pos=(0, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.excel_file, pos=(0, 1), span=(1, 29), flag=wx.EXPAND)
        self.path_sizer.Add(self.excel_btn, pos=(0, 30), span=(1, 1), flag=wx.EXPAND)
        self.path_sizer.Add(self.pdf_label, pos=(1, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.pdf_dir, pos=(1, 1), span=(1, 29), flag=wx.EXPAND)
        self.path_sizer.Add(self.pdf_btn, pos=(1, 30), span=(1, 1), flag=wx.EXPAND)
        self.path_sizer.Add(self.word_label, pos=(2, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.word_dir, pos=(2, 1), span=(1, 29), flag=wx.EXPAND)
        self.path_sizer.Add(self.word_btn, pos=(2, 30), span=(1, 1), flag=wx.EXPAND)
        self.path_sizer.Add(self.sheet_label, pos=(3, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.sheet, pos=(3, 1), span=(1, 29), flag=wx.EXPAND)
        self.path_sizer.Add(self.save_config_btn, pos=(3, 30), span=(1, 1), flag=wx.EXPAND)
        self.path_static_sizer = wx.StaticBoxSizer(wx.StaticBox(self, -1, label=u'路径配置'), wx.VERTICAL)
        self.path_static_sizer.Add(self.path_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        self.top_sizer.Add(self.path_static_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)

        self.filter_row_label = wx.StaticText(self, label=u'表格头行过滤字')
        self.filter_row = wx.TextCtrl(self, style=wx.TE_MULTILINE, size=(200, 60), value=config['filter_row'])
        self.filter_col_label = wx.StaticText(self, label=u'表格头列过滤字')
        self.filter_col = wx.TextCtrl(self, style=wx.TE_MULTILINE, size=(200, 60), value=config['filter_col'])
        self.key_label = wx.StaticText(self, label=u'表格关键字')
        keys = SqlData().get_keys()
        keys_text = []
        for key in keys:
            keys_text.append(key['key'])
        self.keys_value = wx.TextCtrl(self, style=wx.TE_MULTILINE, size=(200, 150), value='|'.join(keys_text))
        self.save_keys_btn = wx.Button(self, label=u'保存关键字')
        self.Bind(wx.EVT_BUTTON, self.save_keys_click, self.save_keys_btn)

        self.key_sizer = wx.GridBagSizer(hgap=10, vgap=10)
        self.key_sizer.Add(self.filter_row_label, pos=(0, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.key_sizer.Add(self.filter_row, pos=(0, 1), span=(1, 30), flag=wx.EXPAND)
        self.key_sizer.Add(self.filter_col_label, pos=(1, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.key_sizer.Add(self.filter_col, pos=(1, 1), span=(1, 30), flag=wx.EXPAND)
        self.key_sizer.Add(self.key_label, pos=(2, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.key_sizer.Add(self.keys_value, pos=(2, 1), span=(1, 29), flag=wx.EXPAND)
        self.key_sizer.Add(self.save_keys_btn, pos=(2, 30), span=(1, 1))
        self.key_static_sizer = wx.StaticBoxSizer(wx.StaticBox(self, -1, label=u'关键字配置'), wx.VERTICAL)
        self.key_static_sizer.Add(self.key_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        self.top_sizer.Add(self.key_static_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)

        self.SetSizerAndFit(self.top_sizer)

    def save_config_click(self, event):
        config = {
            'excel_file': self.excel_file.GetValue().strip(),
            'pdf_dir': self.pdf_dir.GetValue().strip(),
            'word_dir': self.word_dir.GetValue().strip(),
            'sheet': self.sheet.GetValue().strip()
        }
        if SqlData().save_config(config):
            wx.MessageDialog(self, u'保持配置成功', u'消息', wx.OK_DEFAULT).ShowModal()
        else:
            wx.MessageDialog(self, u'保持配置失败', u'消息', wx.OK_DEFAULT).ShowModal()

    def save_keys_click(self, event):
        config = {
            'filter_row': self.filter_row.GetValue().strip(),
            'filter_col': self.filter_col.GetValue().strip(),
            'keys': self.keys_value.GetValue().strip()
        }
        if SqlData().save_keys(config):
            wx.MessageDialog(self, u'保持关键字成功', u'消息', wx.OK_DEFAULT).ShowModal()
        else:
            wx.MessageDialog(self, u'保持关键字失败', u'消息', wx.OK_DEFAULT).ShowModal()

    def excel_click(self, event):
        file_dialog = wx.FileDialog(self, u'选择Excel文件', '', '',
                                    u'excel文件(*.xls)|*.xls', wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if file_dialog.ShowModal() == wx.ID_CANCEL:
            return
        self.excel_file.SetValue(file_dialog.GetPath())

    def pdf_click(self, event):
        dir_dialog = wx.DirDialog(self, u'选择PDF目录', '',
                                  wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if dir_dialog.ShowModal() == wx.ID_CANCEL:
            return
        self.pdf_dir.SetValue(dir_dialog.GetPath())

    def word_click(self, event):
        dir_dialog = wx.DirDialog(self, u'选择WORD目录', '',
                                  wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if dir_dialog.ShowModal() == wx.ID_CANCEL:
            return
        self.word_dir.SetValue(dir_dialog.GetPath())

 
class KeysPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        wx.StaticText(self, label='Keys')


class MainWindow(wx.Frame):
    def __init__(self, title):
        wx.Frame.__init__(self, parent=None,
                          title=title, size=(800, 600),
                          style = wx.DEFAULT_FRAME_STYLE & ~wx.MAXIMIZE_BOX)
        self.Center()
        sys.excepthook = except_hook
        if sys.platform == 'win32':
            exe_name = sys.executable
            icon = wx.Icon(exe_name, wx.BITMAP_TYPE_ICO)
            self.SetIcon(icon)


if __name__ == '__main__':
    app = wx.App()
    main = MainWindow(title=u'提取客户信息')
    nb = wx.Notebook(main, style=wx.NB_FIXEDWIDTH)
    
    run_panel = RunPanel(nb)
    wx.Log.SetActiveTarget(wx.LogTextCtrl(run_panel.text_ctrl))
    nb.AddPage(run_panel, u'运行')
    config_panel = ConfigPanel(nb)
    nb.AddPage(config_panel, u'配置')
    # keys_panel = KeysPanel(nb)
    # nb.AddPage(keys_panel, u'关键字')

    main.Show()
    app.MainLoop()
