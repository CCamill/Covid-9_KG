import re
import docx

def get_name(text):
    PATTERN = u'(?:于)[\u4e00-\u9fa5]{2,4}(?:的)'
    reg = re.compile(PATTERN)
    name_list = str(reg.findall(text))
    name = name_list[3:len(name_list) - 3]
    return name

def get_phonenumber(docx_path):
    doc = docx.Document(docx_path)
    paragraphs = []
    patient_name = get_name(docx_path)
    for para in doc.paragraphs:
        paragraphs.append(para.text)
    text = ''
    if patient_name in paragraphs[3]:
        text = paragraphs[3]  # str
    elif patient_name in paragraphs[4]:
        text = paragraphs[4]  # str
    elif patient_name in paragraphs[2]:
        text = paragraphs[2]  # str
    else:
        return '未知'

    text = text.replace('：',':')
    text = text.replace(' ','')
    text = text.replace('，',',')
    text = text.replace('。',',')
    print(text)
    try:
        reg = re.compile(u'(?:电话.*)+\d{6,11}')
        phone = reg.findall(text)[0]
        try:
            reg = re.compile('1\d{10}')
            phoneNumber = reg.findall(phone)[0]
            return phoneNumber
        except Exception as e:
            reg = re.compile('\d{7}')
            phoneNumber = reg.findall(phone)[0]
            return phoneNumber
    except Exception as e:
        return '未知'
docx_path = r'document/12.04/关于申国亮的调查报告.docx'
phonenumber = get_phonenumber(docx_path)
print(phonenumber)