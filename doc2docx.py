# -*-
# 将doc文档转换为docx文档
# -*-
# @Author : yyzhang
import os
import time
from win32com import client
import re
from format_filename import formalize_filename
def get_name(text):
    PATTERN = u'(?:于)[\u4e00-\u9fa5]{2,12}(?:的)'
    reg = re.compile(PATTERN)
    name_list = str(reg.findall(text))
    name = name_list[3:len(name_list) - 3]
    return name

def doc_to_docx(path_date):
    word = client.Dispatch("kwps.Application")  # 打开word应用程序
    date = path_date.split('\\')[-1]
    filename_list = [i for i in os.listdir(path_date) if i.split(".")[-1] == "doc"]

    try:
        for file in filename_list:
            print("开始转换:", file)
            patient_name = get_name(formalize_filename(file))

            out_name = os.path.join(path_date, '关于{}的调查报告.docx'.format(patient_name))  #
            if not os.path.exists(out_name):
                file_path = os.path.join(path_date, file)
                doc = word.Documents.Open(file_path, PasswordDocument='1111')  # 打开word文件
                doc.SaveAs("{}".format(out_name), 12, False, "", True, "", False,
                           False, False,
                           False)  # 转换后的文件,12代表转换后为docx文件
                doc.Close  # 关闭原来word文件
                os.remove(os.path.join(path_date, file))
                print("转换后：", out_name)
            else:
                print(out_name + '已存在')
                os.remove(os.path.join(path_date, file))
    except Exception as e:
        print(e)
    word.Quit()

if __name__ == "__main__":
    # 支持文件夹批量导入
    path_date =r'D:\桌面\knowledge extract\document\11.23'
    doc_to_docx(path_date)


