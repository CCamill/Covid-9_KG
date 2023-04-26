#获取密接人员信息

import win32com.client as wc
import os
import pandas as pd
import docx


def closesoft():
    print('''挂载程序关闭中……
          ''')
    wcc = wc.constants

    try:
        wps = wc.gencache.EnsureDispatch('kwps.application')
    except:
        wps = wc.gencache.EnsureDispatch('wps.application')
    else:
        wps = wc.gencache.EnsureDispatch('word.application')
    try:
        wps.Documents.Close()
        wps.Documents.Close(wcc.wdDoNotSaveChanges)
        wps.Quit()
    except:
        pass


def doc_to_docx(path):
    path_list = os.listdir(path)
    doc_list = [os.path.join(path,str(i)) for i in path_list if str(i).endswith('doc') and '$' not in str(i)]
    word = wc.gencache.EnsureDispatch('kwps.application')
    wordlist_path=[]
    print("正在读取文件目录....")
    for doc_path in doc_list:
        # doxc_path_save=doc_path.replace(path, "D:\\桌面\\11.21\\")
        # save_path = str(doxc_path_save).replace('doc','docx')
        #
        # doc = word.Documents.Open(doc_path)
        # doc.SaveAs(save_path,12, False, "", True, "", False, False, False, False)
        # doc.Close()
        # wordlist_path.append(save_path)
        # print('{} Save sucessfully '.format(save_path))
        word = wc.Dispatch("Word.Application")  # 打开word应用程序
        doc = word.Documents.Open(doc_path)  # 打开word文件

        a = os.path.split(doc_path)  # 分离路径和文件
        b = os.path.splitext(a[-1])[0]  # 拿到文件名

        doc.SaveAs("{}\\{}.docx".format(r'D:\桌面\covid_docx\11.21', b), 12)  # 另存为后缀为".docx"的文件，其中参数12或16指docx文件
        doc.Close()  # 关闭原来word文件
    word.Quit()
    print("文件转换已完成,开始分析文件....")

    return wordlist_path



def GetData_frompath(save_path):
    document = docx.Document(save_path)
    col_keys = [] # 获取列名
    col_values = [] # 获取列值
    index_num = 0
    for table in document.tables:
        for row_index,row in enumerate(table.rows):
            for col_index,cell in enumerate(row.cells):

                if (col_index==4 and row_index<7) :
                    continue
                # print(' pos index is ({},{})'.format(row_index, col_index))
                # print('cell text is {}'.format(cell.text))
                if row_index<7:
                    if index_num % 2==0:
                        col_keys.append(cell.text)
                    else:
                        col_values.append(cell.text)
                    fore_str = cell.text
                    index_num +=1
                else:
                    if col_index >0 :
                        break
                    else:
                        if index_num % 2 == 0:
                            col_keys.append(cell.text)
                        else:
                            col_values.append(cell.text)
                        fore_str = cell.text
                        index_num += 1

    #col_keys.pop(3)
    print(col_values)
    #col_values.pop(3)
    colLen = len(col_values)
    print(f'col keys is {col_keys}')
    print(f'col values is {col_values}')
    #col_values[colLen] = '\t' + col_values[colLen]


    return  col_keys,col_values



def create_csv(wordlist_path):
    pd_data = []
    for index,single_path in enumerate(wordlist_path):
        col_names,col_values = GetData_frompath(single_path)
        if index == 0:
            pd_data.append(col_names)
            pd_data.append(col_values)
        else:
            pd_data.append(col_values)

    df = pd.DataFrame(pd_data)
    #csv_path="C:\\Users\Lenovo\PycharmProjects\infoexarct\\finsh.csv"
    # df.to_csv(csv_path, encoding='utf_8_sig',index=False,header=None)
    #df.to_excel('data.xlsx', sheet_name='工作表1',index=False,header=None)
    #print("程序执行结束.....")


#wordlist_path=doc_to_docx(path)

if __name__ == "__main__":
    path = r"D:\桌面\covid_document\11.21"
    wordlist_path = doc_to_docx(path)
    print(wordlist_path)

    # create_csv(doc_list)
