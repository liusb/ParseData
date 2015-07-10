# -*- coding: utf-8 -*-

import wx
import sqlite3
from threading import Thread
import os
import win32com.client
import pythoncom
import sys
import traceback
import re


def except_hook(etype, value, tb):
    message = 'Uncaught exception:\n'
    message += ''.join(traceback.format_exception(etype, value, tb))
    wx.LogMessage(message)


class SqlData():
    def __init__(self, db_name=None):
        if db_name is None:
            db_name = 'config.db'
        self.conn = sqlite3.connect(db_name)
        self.conn.text_factory = str

    def __del__(self):
        self.conn.close()

    def get_config(self):
        cur = self.conn.cursor()
        cur.execute('select name, value from config')
        rows = cur.fetchall()
        config = {}
        for row in rows:
            config[row[0]] = row[1]
        cur.close()
        return config

    def save_config(self, config):
        sql = """update config set value = '%s' where name = '%s' """
        cur = self.conn.cursor()
        try:
            for key, value in config.items():
                temp = (sql % (value, key)).encode('gbk')
                cur.execute(temp)
            self.conn.commit()
        except sqlite3.Error:
            self.conn.rollback()
            return False
        finally:
            cur.close()
        return True

    def get_keys(self):
        cur = self.conn.cursor()
        cur.execute('select id, key, row, col from keys order by id ')
        rows = cur.fetchall()
        keys = []
        for row in rows:
            key = {}
            key['id'] = row[0]
            key['key'] = row[1]
            key['row'] = row[2]
            key['col'] = row[3]
            keys.append(key)
        return keys


class Word:

    def __init__(self):
        self.word = win32com.client.Dispatch('Word.Application')
        self.word.Visible = True
        self.word.ScreenUpdating = True
        self.word.DisplayAlerts = False
        self.doc = None

    def __del__(self):
        if self.word.Documents.Count == 0:
            self.word.Quit()

    def open(self, name):
        self.doc = self.word.Documents.Open(name)

    def close(self):
        self.doc.Close()
        self.doc = None

    def find_next_paragraph(self, range_obj):
        count = 1
        next_one = None
        while True:
            next_one = range_obj.Next(Unit=4, Count=count)
            if len(next_one.Text) > 2:
                break
            count += 1
        return next_one

    def find_previous_paragraph(self, range_obj):
        count = 1
        previous_range = None
        while True:
            previous_range = range_obj.Previous(Unit=4, Count=count)
            if len(previous_range.Text) > 2:
                break
            count += 1
        return previous_range

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
                    text = text[:-2].replace(' ', '').lstrip('\r')
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
                    text = text[:-2].replace(' ', '').lstrip('\r')
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
            # print 21, row2_1
            # print 22, row2_2
            # print 23, row2_3
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

    def parse_table(self, table):
        col_ct = table.Columns.Count
        row_ct = table.Rows.Count
        if col_ct == 2:
            wx.LogMessage(u'WORD表格行数:%d, 列数:%d，按列继续处理' % (row_ct, col_ct))
            result = self.col_yes_parse(table)
        elif table.Uniform is True:
            wx.LogMessage(u'WORD表格是个规范表格 行数:%d, 列数:%d，按行继续处理' % (row_ct, col_ct))
            result = self.uniform_parse(table)
        elif row_ct <= 7:
            wx.LogMessage(u'WORD表格行数:%d, 列数:%d，按行继续处理' % (row_ct, col_ct))
            result = self.row_yes_parse(table)
        else:
            wx.LogMessage(u'WORD表格行数:%d, 列数:%d，按列继续处理' % (row_ct, col_ct))
            result = self.col_yes_parse(table)
        return result

    def parse(self, keys):
        result = {}
        for key in keys:
            wx.LogMessage(u'查找关键【%s】字' % key['key'].decode('gbk'))
            cur_pos = self.doc.Content.Start
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
                selection.ExtendMode = True
                selection.Extend()
                selection.ExtendMode = False
                find_s = selection.Start
                find_e = selection.End
                if find_s < cur_pos:
                    wx.LogWarning(u'==重复跳转到文件头，查找下一个关键字==')
                    break
                if find_e - find_s > 100:
                    cur_pos = find_e - 2
                    wx.LogMessage(u'关键字所在段落太长，排除关键字')
                    continue
                key_range = self.doc.Range(find_s, find_e)
                wx.LogMessage(u'在%d页查找到关键字' % key_range.Information(3))
                if key_range.Information(12) is True:  # wdWithInTable = 12
                    cur_pos = find_e-2
                    wx.LogMessage(u'关键字在表格中，排除关键字')
                    continue

                # 向后读到一个表为止
                find_table = []
                while True:
                    table_range = key_range.Next(Unit=15)  # 15 => table
                    if table_range is None:
                        break
                    table_s = table_range.Start
                    table_e = table_range.End
                    table_t = table_range.Tables(1)
                    if table_s - find_e > 500:
                        cur_pos = table_s - 500
                        wx.LogMessage(u'关键字距离表格太远，排除表格')
                        continue
                    cur_pos = table_e
                    find_table = self.parse_table(table_t)
                    table_page = table_range.Information(3)  # 页码

                    tow_table_flag = False
                    next_table_range = table_range.Next(Unit=15)
                    next_table_page = next_table_range.Information(3)  # 页码
                    if table_page + 1 == next_table_page:
                        next_table_s = next_table_range.Start
                        next_table_e = next_table_range.End
                        table_between_ct = self.doc.Range(table_e, next_table_s).ComputeStatistics(Statistic=6)
                        if table_between_ct == 0:
                            tow_table_flag = True
                            cur_pos = next_table_e
                            wx.LogMessage(u'表格分页，分别在%d，%d两页' % (table_page, next_table_page))
                            next_table_t = next_table_range.Tables(1)
                            find_next_table = self.parse_table(next_table_t)
                            if len(find_next_table) > 1 and len(find_table) > 0 \
                                    and find_table[0][0] == find_next_table[0][0]:
                                find_next_table.pop(0)
                        find_table.extend(find_next_table)
                if len(find_table) == 0:  # 表不合规则
                    continue
                table_index = (table_page, table_s)
                if table_index in result.keys():
                    table_item = result[table_index]
                    if table_item['key']['find_s'] < find_s:  # 关键字离表格更近
                        table_item['key']['id'] = key['id']
                        table_item['key']['key'] = key['key']
                        table_item['key']['find_s'] = find_s
                    continue

                table_row = len(find_table)
                if table_row > 10:
                    wx.LogMessage(u'表格行数超过10行，排除')
                    continue

                other_table = False
                for item in find_table[0]:
                    if u'税' in item or u'分红' in item or u'行业' in item \
                            or u'产品' in item or u'损益' in item or u'支出' in item\
                            or u'接待' in item or u'资产' in item or u'成本' in item:
                        other_table = True
                        break
                if other_table is False:
                    for item in find_table:
                        if u'营业税' in item[0] or u'应收' in item[0] or u'合同' in item[0] \
                                or u'逾期' in item[0] or u'费用' in item[0] or u'情况' in item[0]:
                            other_table = True
                            break
                if other_table is True:
                    wx.LogMessage(u'取到其他表格，排除')
                    continue

                if len(find_table[-1]) == 3 and len(find_table[0]) == 4:
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
                if len(find_table) == 1 and u'客户名称' in find_table[0]:
                    wx.LogWarning(u'找到只有标题的空表，排除')
                    continue
                elif len(find_table) == 2 and u'客户名称' in find_table[0] and u'/' == find_table[1][0]:
                    wx.LogWarning(u'找到只有标题的空表，排除')
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
                table_item['table']['before2'] = before2.Text.replace(' ', '')
                table_item['table']['before1'] = before1.Text.replace(' ', '')
                table_item['table']['after1'] = after1.Text.replace(' ', '')
                table_item['table']['after2'] = after2.Text.replace(' ', '')
                table_item['table']['row'] = table_row
                table_item['table']['col'] = table_col
                table_item['table']['content'] = find_table
                result[table_index] = table_item

        return result


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
        self.keys = None
        self.word = None

    def init_word(self):
        self.word = Word()
        self.keys = SqlData().get_keys()

    def open(self):
        self.wkbk = self.excel.Workbooks.Open(self.config['excel_file'])
        self.wksht = self.wkbk.Worksheets(1)

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
                                                     Address=os.path.join(config['word_dir'], item[:-3]+'doc'),
                                                     TextToDisplay=item[:-4])
                    if link.TextToDisplay != item[:-4]:
                        link.TextToDisplay = item[:-4]
                self.wksht.Range('D%d:D%d' % (begin_row, end_row)).Value = item[0:4]
                self.wksht.Range('E%d:E%d' % (begin_row, end_row)).Value = item[5:11]
                begin_row = end_row + 1
                count += 1
        wx.LogMessage(u'总共找到%d个文件' % (count-1))
        self.save()
        wx.LogMessage(u'成功将找到的文件初始化到Excel中')

    def save(self):
        self.wkbk.Save()

    def next_doc(self):
        self.current_begin = self.current_end + 1
        old = self.wksht.Cells(self.current_begin, 1).Value
        if old is None:
            wx.LogMessage(u'未找到更多的记录，处理结束')
            return False
        row = self.current_begin
        while True:
            row += 1
            if self.wksht.Cells(row, 1).Value != old:
                break
        self.current_end = row - 1
        return True

    def is_processed(self):
        if self.wksht.Cells(self.current_begin, 3).Value is not None\
                and self.wksht.Cells(self.current_begin, 3).Value == u'是':
            return True
        return False

    def adjust_row(self, row_ct):
        begin_row = self.current_begin
        end_row = self.current_end
        if end_row - begin_row + 1 == row_ct:  # 正好
            wx.LogMessage(u'调整表格：行数正好符合')
        elif end_row - begin_row + 1 > row_ct:  # 表格有多
            self.wksht.Range('A%d:E%d' % (begin_row+row_ct, end_row)).EntireRow.Delete()
            self.current_end = begin_row + row_ct - 1
            wx.LogMessage(u'调整表格：行数有多，删除%d行' % (end_row-self.current_end))
        else:  # 还有空余行，但不够用
            self.current_end = begin_row + row_ct - 1
            self.wksht.Range('A%d:E%d' % (end_row+1, self.current_end)).EntireRow.Insert()
            self.wksht.Range('A%d:E%d' % (begin_row, begin_row)).Copy()
            self.wksht.Paste(self.wksht.Range('A%d:E%d' % (end_row+1, self.current_end)))
            wx.LogMessage(u'调整表格：行数不够，添加%d行' % (self.current_end-end_row))

    def process_doc(self):
        doc_name = self.wksht.Cells(self.current_begin, 2).Hyperlinks(1).TextToDisplay + '.doc'
        wx.LogMessage(u'正在处理%s' % doc_name)
        self.word.open(os.path.join(self.config['word_dir'].decode('gbk'), doc_name))
        result = self.word.parse(self.keys)
        wx.LogMessage(u'查找完毕')
        self.write_data(result)
        wx.LogMessage(u'处理%s完毕' % doc_name)
        self.wkbk.Save()
        self.word.close()

    def write_data(self, tables):
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
            self.wksht.Cells(self.current_begin, 20).Value = u'第一个明细表不适用'
        self.adjust_row(row_ct)
        count = 1
        end_row = self.current_begin - 1
        for key in sorted_key:
            table = tables[key]
            table_row = table['table']['row']
            table_col = table['table']['col']
            begin_row = end_row + 1
            end_row = end_row + table_row
            self.wksht.Range('F%d:F%d' % (begin_row, end_row)).Value = table['key']['id']
            self.wksht.Range('G%d:G%d' % (begin_row, end_row)).Value = table['key']['key']
            self.wksht.Range('I%d:I%d' % (begin_row, end_row)).Value = table['table']['row']
            self.wksht.Range('J%d:J%d' % (begin_row, end_row)).Value = table['table']['before2']
            self.wksht.Range('K%d:K%d' % (begin_row, end_row)).Value = table['table']['before1']
            self.wksht.Range('L%d:L%d' % (begin_row, end_row)).Value = table['table']['after1']
            self.wksht.Range('M%d:M%d' % (begin_row, end_row)).Value = table['table']['after2']
            self.wksht.Range('N%d:N%d' % (begin_row, end_row)).Value = u'表%d' % count
            self.wksht.Range('O%d:O%d' % (begin_row, end_row)).Value = table['table']['page']
            begin_col = 16
            if table_col < 4:
                begin_col = 20 - table_col
            row_i = begin_row
            for row in table['table']['content']:
                col_i = begin_col
                for item in row:
                    self.wksht.Cells(row_i, col_i).Value = item
                    col_i += 1
                row_i += 1
            count += 1
        self.wksht.Range('C%d:C%d' % (self.current_begin, self.current_end)).Value = u'是'
        self.wksht.Range('H%d:H%d' % (self.current_begin, self.current_end)).Value = len(tables)


class WorkerThread(Thread):

    def __init__(self):
        Thread.__init__(self)
        self.running = False

    def run(self):
        pythoncom.CoInitialize()
        self.running = True
        excel = Excel()
        excel.open()
        excel.init_word()
        while self.running:
            if excel.next_doc() is False:
                break
            if excel.is_processed() is True:
                continue
            excel.process_doc()
        wx.LogMessage('work end')

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
        
    def start_click(self, event):
        wx.LogMessage('work begin')
        self.stop_btn.Enable()
        self.start_btn.Disable()
        self.thread = WorkerThread()
        self.thread.start()

    def stop_click(self, event):
        self.start_btn.Enable()
        self.stop_btn.Disable()
        self.thread.stop()
        wx.LogMessage('waiting for work stop')

    def excel_init_click(self, event):
        wx.LogMessage('excel init begin')
        self.excel_init_btn.Disable()
        try:
            excel = Excel()
            excel.open()
            excel.init()
        finally:
            self.excel_init_btn.Enable()
        wx.LogMessage('excel init end')

    def excel_clear_click(self, event):
        dlg = wx.MessageDialog(None, u"是否确认清空Excel", u"确认信息", wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_NO:
            return
        wx.LogMessage('excel clear begin')
        self.excel_clear_btn.Disable()
        try:
            excel = Excel()
            excel.open()
            excel.clear()
        finally:
            self.excel_clear_btn.Enable()
        wx.LogMessage('excel clear end')


class ConfigPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.config = SqlData().get_config()

        self.excel_label = wx.StaticText(self, label=u'EXCEL文件')
        self.excel_file = wx.TextCtrl(self, value=self.config['excel_file'])
        self.excel_btn = wx.Button(self, label=u'选择文件')
        self.Bind(wx.EVT_BUTTON, self.excel_click, self.excel_btn)

        self.pdf_label = wx.StaticText(self, label=u'PDF目录')
        self.pdf_dir = wx.TextCtrl(self, value=self.config['pdf_dir'])
        self.pdf_btn = wx.Button(self, label=u'选择目录')
        self.Bind(wx.EVT_BUTTON, self.pdf_click, self.pdf_btn)

        self.word_label = wx.StaticText(self, label=u'WORD目录')
        self.word_dir = wx.TextCtrl(self, value=self.config['word_dir'])
        self.word_btn = wx.Button(self, label=u'选择目录')
        self.Bind(wx.EVT_BUTTON, self.word_click, self.word_btn)

        self.no_label = wx.StaticText(self, label=u'无关键字年报目录')
        self.no_dir = wx.TextCtrl(self, value=self.config['no_dir'])
        self.no_btn = wx.Button(self, label=u'选择目录')
        self.Bind(wx.EVT_BUTTON, self.no_click, self.no_btn)

        self.save_btn = wx.Button(self, label=u'保存设置')
        self.Bind(wx.EVT_BUTTON, self.save_click, self.save_btn)

        self.top_sizer = wx.BoxSizer(wx.VERTICAL)
        self.path_sizer = wx.GridBagSizer(hgap=10, vgap=10)
        self.path_sizer.Add(self.excel_label, pos=(0, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.excel_file, pos=(0, 1), span=(1, 28), flag=wx.EXPAND)
        self.path_sizer.Add(self.excel_btn, pos=(0, 29), span=(1, 1), flag=wx.EXPAND)
        self.path_sizer.Add(self.pdf_label, pos=(1, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.pdf_dir, pos=(1, 1), span=(1, 28), flag=wx.EXPAND)
        self.path_sizer.Add(self.pdf_btn, pos=(1, 29), span=(1, 1), flag=wx.EXPAND)
        self.path_sizer.Add(self.word_label, pos=(2, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.word_dir, pos=(2, 1), span=(1, 28), flag=wx.EXPAND)
        self.path_sizer.Add(self.word_btn, pos=(2, 29), span=(1, 1), flag=wx.EXPAND)
        self.path_sizer.Add(self.no_label, pos=(3, 0), span=(1, 1), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTRE_VERTICAL)
        self.path_sizer.Add(self.no_dir, pos=(3, 1), span=(1, 28), flag=wx.EXPAND)
        self.path_sizer.Add(self.no_btn, pos=(3, 29), span=(1, 1), flag=wx.EXPAND)
        self.path_sizer.Add(self.save_btn, pos=(4, 29), span=(1, 1), flag=wx.EXPAND)
        self.static_sizer = wx.StaticBoxSizer(wx.StaticBox(self, -1, label=u'路径配置'), wx.VERTICAL)
        self.static_sizer.Add(self.path_sizer, proportion=0, flag=wx.EXPAND, border=10)

        self.top_sizer.Add(self.static_sizer, 0, wx.EXPAND, 20)
        self.SetSizerAndFit(self.top_sizer)

    def save_click(self, event):
        self.config['excel_file'] = self.excel_file.GetValue()
        self.config['pdf_dir'] = self.pdf_dir.GetValue()
        self.config['word_dir'] = self.word_dir.GetValue()
        self.config['no_dir'] = self.no_dir.GetValue()
        if SqlData().save_config(self.config):
            wx.MessageDialog(self, u'保持配置成功', u'消息', wx.OK_DEFAULT).ShowModal()
        else:
            wx.MessageDialog(self, u'保持配置失败', u'消息', wx.OK_DEFAULT).ShowModal()

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

    def no_click(self, event):
        dir_dialog = wx.DirDialog(self, u'选择没有关键字年报目录', '',
                                  wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if dir_dialog.ShowModal() == wx.ID_CANCEL:
            return
        self.no_dir.SetValue(dir_dialog.GetPath())

 
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
