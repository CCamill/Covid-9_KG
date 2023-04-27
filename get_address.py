import os
import re
import docx


def get_name(text):
    PATTERN = u'(?:于)[\u4e00-\u9fa5]{2,4}(?:的)'
    reg = re.compile(PATTERN)
    name_list = str(reg.findall(text))
    name = name_list[3:len(name_list) - 3]
    return name

def get_address(docx_path):
    doc = docx.Document(docx_path)
    paragraphs = []
    patient_name = get_name(docx_path)
    for para in doc.paragraphs:
        paragraphs.append(para.text)

    if patient_name in paragraphs[3]:
        text = paragraphs[3]  # str
    elif patient_name in paragraphs[4]:
        text = paragraphs[4]  # str
    elif patient_name in paragraphs[2]:
        text = paragraphs[2]  # str
    else:
        return '未知'

    text = text.replace('：', ':')
    text = text.replace('；', ',')
    text = text.replace(';', ',')
    text = text.replace(' ', '')
    text = text.replace('，', ',')
    text = text.replace('。', ',')
    text = text.replace('（', '(')
    text = text.replace('）', ')')
    print(text)
    try:
        p = u'(?:址:)[\u4e00-\u9fa5a-zA-Z0-9().-]{1,35}(?:,)'
        reg = re.compile(p)
        res = reg.findall(text)
        address = res[0][2:-1]
        return address
    except Exception as e:
        return '未知'

if __name__ == '__main__':
    docx_path = r'document/11.28/关于李利锋的调查报告.docx'
    address = get_address(docx_path)
    print(address)



