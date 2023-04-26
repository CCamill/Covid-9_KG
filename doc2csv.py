'''
把doc文档转换为docx文档
再从docx文档中的表格提取到csv文件
'''
import os
import time
import re
from win32com import client
from docx import Document
import csv
from format_filename import formalize_filename

class config:
    docxCatalog = r'D:\桌面\covid_docx'
    csvCatalog = r'D:\桌面\covid_csv'
    docCatalog = r'D:\桌面\covid_document'

def get_name(text):
    PATTERN = u'(?:于)[\u4e00-\u9fa5]{2,11}(?:的)'
    reg = re.compile(PATTERN)
    name_list = str(reg.findall(text))
    name = name_list[3:len(name_list) - 3]
    return name


def data_write_csv(file_name, datas):  # file_name为写入CSV文件的路径，datas为要写入数据列表
    file_csv = open(file_name, 'w+', newline='')  # 追加
    writer = csv.writer(file_csv)  # , delimiter=' ', quotechar=' ', quoting=csv.QUOTE_MINIMAL)
    for data in datas:
        writer.writerow(data)
    print("保存文件成功，处理结束")


def data2csv(docFile):
    date = docFile.split('\\')[-1]
    saveCatalog = os.path.join(r'D:\桌面\covid_csv', date)
    if not os.path.exists(saveCatalog):
        os.makedirs(saveCatalog)

    path_list = os.listdir(docFile)
    print(docFile)
    docx_list = [os.path.join(docFile, str(i)) for i in path_list if str(i).endswith('docx') and '$' not in str(i)]
    print(docx_list)
    for docx in docx_list:
        document = Document(docx)  # 读入文件
        tables = document.tables  # 获取文件中的表格集
        if len(tables) != 0:  # 判断是否有表格
            patientName = get_name(docx)
            savePath = os.path.join(saveCatalog, "关于{}的密接登记表.csv".format(patientName))
            data = []
            for table in tables[:]:
                for i, row in enumerate(table.rows[:]):  # 读每行
                    row_content = []
                    for cell in row.cells[:]:  # 读一行中的所有单元格
                        c = cell.text
                        row_content.append(c.strip("\n"))
                    data.append(row_content)
            print(data)
            if ['涉密材料\n严禁外泄'] in data:
                data.remove(['涉密材料\n严禁外泄'])
            data_write_csv(savePath, data)
        else:
            print(docx, '文件没有表格')


def doc_to_docx(path_date):
    date = path_date.split('\\')[-1]
    save_path = os.path.abspath(config.docCatalog + '\\' + date)  #docx文件保存到源目录
    # save_path = os.path.abspath(config.docxCatalog + '\\' + date)  #docx文件保存到其他目录
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    word = client.Dispatch("kwps.Application")  # 打开word应用程序
    patients = [i for i in os.listdir(path_date) if str(i).endswith('doc') and '$' not in str(i)]

    try:
        for patient in patients:
            file_path = os.path.join(path_date,patient)
            print('当前文档', file_path)
            patient_name = get_name(formalize_filename(patient))
            out_name = os.path.join(save_path, "关于{}的调查报告.docx".format(patient_name))  #
            if os.path.exists(out_name):
                print("{} 文件已存在".format(out_name))
                os.remove(file_path)
                print('已移除{}'.format(file_path))
            else:
                doc = word.Documents.Open(file_path, PasswordDocument=1111)  # 打开word文件
                doc.SaveAs(out_name, 12, False, "", True, "", False,
                           False, False,
                           False)
                doc.Close  # 关闭原来word文件
                print("成功转换：{}".format(out_name))
                os.remove(file_path)
                print('已移除{}'.format(file_path))
    except Exception as e:
        print(e, '转换失败')
    word.Quit()


if __name__ == "__main__":
    # 支持文件夹批量导入
    # 路径格式 ./document/date/*.doc
    dates = os.listdir(config.docCatalog)
    doc_path_dates = []
    for i in dates:
        doc_path_dates.append(os.path.join(config.docCatalog, i))
    for path_date in doc_path_dates:
        # print(list_dir)
        doc_to_docx(path_date)

    docx_path_dates = []
    for i in os.listdir(config.docCatalog):
        docx_path_dates.append(os.path.join(config.docCatalog, i))
    for docFile in docx_path_dates:
        data2csv(docFile)
