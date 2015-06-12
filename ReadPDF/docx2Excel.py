# -*- coding: utf-8 -*-

import os
import win32com.client

search = ['公司前五大客户', '公司前5大客户资料', '客户资料']

if __name__ == "__main__":
    docxs = []
    dirname = u'100份pdf及对应word'
    new_dir = u'part_word'
    xls_name = 'Docx2Excel.xls'
    rpath = os.path.split(os.path.realpath(__file__))[0]

    dirs = os.listdir(dirname)
    for item in dirs:
        if item.endswith('.docx'):
            docxs.append(item)

    #excel = win32com.client.Dispatch('Excel.Application')
    #excel.Visible = True
    #xls = excel.Workbooks.Open(os.path.join(rpath, xls_name))

    word = win32com.client.Dispatch('Word.Application')
    # word.Visible = True
    # word.ScreenUpdating = True
    word.DisplayAlerts = False


    for item in docxs:
        new_name = item[:-5] + '_part.docx'
        new_name = os.path.join(rpath, new_dir, new_name)
        if os.path.isfile(new_name):
            continue
            #os.remove(new_name)
        print item
        doc = word.Documents.Open(os.path.join(rpath, dirname, item))
        new_doc = word.Documents.Add()
        findText = u'[五5][名大]客户'
        cur_pos = 1
        found = False
        while True:
            doc.Activate()
            sele = word.Selection
            sele.SetRange(cur_pos, cur_pos) #设置光标
            find = sele.Find
            find.MatchWildcards = True
            if find.Execute(FindText=findText, Forward=True) is not True:
                break
            found = True
            # 找到之后扩展成整个句子
            sele.ExtendMode = True
            sele.Extend()
            find_s = sele.Start
            find_e = sele.End
            if find_s < cur_pos:
                break

            # 向后读到一个表为止
            table_t = sele.Next(Unit=15)  # 15 => table
            if table_t is None:
                break
            table_t.Select()
            table_s = sele.Start
            table_e = sele.End
            sele.SetRange(table_s, table_e)

            cur_pos = table_e

            # 复制到一个新的word文档中
            doc.Range(find_s, table_e).Copy()
            new_doc.Activate()
            sele = word.Selection            
            new_pos = new_doc.Content.End
            sele.SetRange(new_pos, new_pos)
            sele.Paste()
            sele.InsertAfter('==============================================\n')

        if found:
            new_doc.SaveAs(new_name)
        
        new_doc.Close()
        doc.Close()

    word.Quit()
