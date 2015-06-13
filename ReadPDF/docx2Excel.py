# -*- coding: utf-8 -*-

import os
import win32com.client
import traceback
import sys

# search = ['公司前五大客户', '公司前5大客户资料', '客户资料']
findText = u'公司前[五5][名大]客户'

excel = None
wkbk = None
wksht = None
word = None


def check_table(table):

    col_ct = table.Columns.Count
    row_ct = table.Rows.Count
    if row_ct < 6 or col_ct < 3:
        return None
    
    print 'row_ct %d, col_ct %d'%(row_ct, col_ct)

    result = []
    col_i = 1
    while col_i <= col_ct and col_i <= 5:
        row_content = []
        row_i = 1
        while row_i <= row_ct and row_i <= 7:
            try:
                text = table.Cell(Row=row_i, Column=col_i).Range.Text
            except:
                print 'fail to acess Cell(%d, %d)'%(row_i, col_i)
            text = text[:-1]
            if len(text) == 0:
                continue
            row_content.append(text)
            row_i = row_i + 1
        result.append(row_content)
        col_i = col_i + 1
    return result


if __name__ == "__main__":
    dirname = u'word'
    new_dir = u'part_word'
    xls_name = 'Docx2Excel.xls'

    try:
        rpath = os.path.split(os.path.realpath(__file__))[0]
    except NameError:  # We are the main py2exe script, not a module
        rpath = os.path.split(os.path.realpath(sys.argv[0]))[0]

    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True
        wkbk = excel.Workbooks.Open(os.path.join(rpath, xls_name))
        wksht = wkbk.Worksheets(1)
        
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = True
        word.ScreenUpdating = True
        word.DisplayAlerts = True
    except:
        print "Unexpected error:", traceback.print_exc()
        raw_input('Press any key to quit >')
        sys.exit(0)

    if wksht.Cells(1,1).Value is None:
        dirs = os.listdir(dirname)
        row = 0
        for item in dirs:
            if item.endswith('.docx'):
                row = row + 1
                wksht.Hyperlinks.Add(Anchor=wksht.Range("A%d"%row),
                     Address=os.path.join(dirname, item),
                     TextToDisplay=item)

    row = 1
    while True:
        if wksht.Cells(row, 1).Value is None:
            break
        if wksht.Cells(row, 2).Value is not None:
            row = row + int(wksht.Cells(row, 2).Value)
            continue

        doc_name = wksht.Cells(row, 1).Hyperlinks(1).TextToDisplay
        new_name = doc_name[:-5] + '_part.docx'
        new_name = os.path.join(rpath, new_dir, new_name)
        if os.path.isfile(new_name):
            #os.remove(new_name)
            new_doc = None
        else:
            new_doc = word.Documents.Add()            

        doc = word.Documents.Open(os.path.join(rpath, dirname, doc_name))
        
        cur_pos = doc.Content.Start
        found = 0
        add_row = 1
        while True:
            doc.Activate()
            sele = word.Selection
            sele.SetRange(cur_pos, cur_pos) #设置光标
            find = sele.Find
            find.MatchWildcards = True
            if find.Execute(FindText=findText, Forward=True) is not True:
                break

            # 找到之后扩展成整个句子
            sele.ExtendMode = True
            sele.Extend()
            find_s = sele.Start
            find_e = sele.End
            if find_s < cur_pos:
                break
            find_key = doc.Range(find_s, find_e).Text

            # 向后读到一个表为止
            table_range = sele.Next(Unit=15)  # 15 => table
            if table_range is None:
                break
            table_range.Select()
            table_s = sele.Start
            table_e = sele.End
            table_t = table_range.Tables(1)

            # 复制到一个新的word文档中
            if new_doc is not None:
                doc.Range(find_s, table_e).Copy()
                new_doc.Activate()
                sele = word.Selection            
                new_pos = new_doc.Content.End
                sele.SetRange(new_pos, new_pos)
                sele.Paste()
                sele.InsertAfter('==============================================\n')

            found = found + 1
            cur_pos = table_e

            find_key = doc.Range(find_s, find_e).Text
            find_content = doc.Range(find_e+1, table_s-1).Text
            find_table = check_table(table_t)

            if find_table is None:  # 表不合规则
                continue
            table_row = 0
            for item in find_table:
                if table_row < len(item):
                    table_row = len(item)
            
            for i in range(row+add_row, row+add_row+table_row+1):
                rangeObj = wksht.Range('C%d:K%d'%(i, i))
                rangeObj.EntireRow.Insert()
            wksht.Cells(row+add_row, 2).Value = find_key
            wksht.Cells(row+add_row, 3).Value = find_content
            col_i = 2
            for col in find_table:
                row_i = row + add_row + 1
                for item in col:
                    wksht.Cells(row_i, col_i).Value = item
                    row_i = row_i + 1
                col_i = col_i + 1
            add_row = add_row + table_row + 1

        wksht.Cells(row, 2).Value = add_row
        wkbk.Save()
        row = row + add_row # 下一个

        if new_doc is not None:
            if found:
                new_doc.SaveAs(new_name)
            new_doc.Close()
        doc.Close()

    word.Quit()
