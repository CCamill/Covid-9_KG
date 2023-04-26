'''
提取docx文档中的表格到csv文件
'''
from docx import Document
import csv
import re
import os

def get_name(text):
    PATTERN = u'(?:于)[\u4e00-\u9fa5]{2,11}(?:的)'
    reg = re.compile(PATTERN)
    name_list = str(reg.findall(text))
    name = name_list[3:len(name_list) - 3]
    return name

def data_write_csv(file_name, datas):#file_name为写入CSV文件的路径，datas为要写入数据列表
    file_csv = open(file_name,'w+',newline='')#追加
    writer = csv.writer(file_csv) #, delimiter=' ', quotechar=' ', quoting=csv.QUOTE_MINIMAL)
    for data in datas:
        writer.writerow(data)
    print("保存文件成功，处理结束")


def data2csv(docFile):
    date = docFile.split('\\')[-1]
    saveCatalog = os.path.join(r'csvfile', date)
    if not os.path.exists(saveCatalog):
        os.makedirs(saveCatalog)


    path_list = os.listdir(docFile)
    print(docFile)
    docx_list = [os.path.join(docFile, str(i)) for i in path_list if str(i).endswith('docx') and '$' not in str(i)]
    print(docx_list)
    for docx in docx_list:
        document = Document(docx)
        tables = document.tables
        if len(tables) != 0: #判断是否有表格
            patientName = get_name(docx)
            savePath = os.path.join(saveCatalog, "关于{}的密接登记表.csv".format(patientName))
            data = []
            for table in tables[:]:
                for i, row in enumerate(table.rows[:]):   # 读每行
                    row_content = []
                    for cell in row.cells[:]:  # 读一行中的所有单元格
                        c = cell.text
                        row_content.append(c.strip("\n"))
                    data.append(row_content)
            print(data)
            if ['涉密材料\n严禁外泄'] in data:
                data.remove(['涉密材料\n严禁外泄'])
            data_write_csv(savePath, data)
        else:   print(docx , '文件没有表格')

if __name__ == '__main__':
    # docFile = r'D:\桌面\covid_docx\11.21'
    # data2csv(docFile)
    path_dates = []
    for i in os.listdir(r'D:\桌面\新建文件夹\covid_docx'):
        path_dates.append(os.path.join(r'D:\桌面\新建文件夹\covid_docx', i))
    for docFile in path_dates:
        data2csv(docFile)