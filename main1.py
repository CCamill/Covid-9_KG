import os
import re
import win32com.client as wc
import docx

file_name = "D:\桌面\covid_document"

def doc_to_docx(path):
    path_list = os.listdir(path)
    doc_list = [os.path.join(path,str(i)) for i in path_list if str(i).endswith('doc')]
    word = wc.Dispatch('Word.Application')
    wordlist_path=[]
    print("正在读取文件目录....")
    for doc_path in doc_list:
        doxc_path_save=doc_path.replace(path, "D:\\桌面\\11.21\\")
        save_path = str(doxc_path_save).replace('doc','docx')
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(save_path,12, False, "", True, "", False, False, False, False)
        doc.Close()
        wordlist_path.append(save_path)
        print('{} Save sucessfully '.format(save_path))
    word.Quit()
    print("文件转换已完成,开始分析文件....")

    return

def get_table(docPath):
    for index,str in enumerate(docPath):
        if str == "于":  start_index = index
        if str == "的":  end_index = index
    docNmae = docPath[start_index + 1:end_index]

    patientDict = {}
    patientDict['名称'] = docNmae
    docStr = docx.Document(docPath)
    numTables = docStr.tables    # 获取Word文档中所有表格

    my_list = [[],[],[],[]]    # 把数据放到列表中
    for table in numTables:
    # 行列个数
        row_count = len(table.rows)#行数
        col_count = len(table.columns)#列数

        for i in range(row_count):
            row = table.rows[i].cells
            for j in range(col_count):
                content = row[j].text
                my_list[i].append(content)
    print(patientDict['名称'] + "密接登记表")
    for row in my_list:
        print(row)

if __name__ == '__main__':
    # docPath = r"D:\桌面\11.21_docx\关于余利娟的调查报告.docx"
    # path = r'D:\桌面\covid_docx\11.21'
    # docxs = os.listdir(path)
    # for docx in docxs:
    #     get_table(docx)
    get_table(r'D:\桌面\covid_docx\11.21\关于赵荣粉的调查报告.docx')
