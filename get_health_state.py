import os

import docx
import re

cn_sym = ['，','。','、','、','；','：','？']

def get_health_state(docx_path):
    doc = docx.Document(docx_path)
    paragraphs = []
    text = ''
    for para in doc.paragraphs[2:6]:
        if '健康状况' in para.text:
            text = para.text
    if text != '':
        text = text.replace('：',':')
        text = text.replace(' ','')
        text = text.replace('，',',')
        text = text.replace('、',',')
        print('原文本: ',text)
        index = text.index(':')
        state = text[index+1:]
        if state[-1] in cn_sym:
            print(state[0:-1])
            print('\n')
        else:
            print(state)
            print('\n')
    else:
        print('未知\n')

path = r'document/12.04/关于徐红勤的调查报告.docx'
get_health_state(path)