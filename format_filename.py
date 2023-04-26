import os
import re

def str_insert(filename, str):
    index = filename.find(str)
    slist = list(filename)
    slist.insert(index, '的')
    sout = ''.join(slist)
    return sout

factor_str = ['.','(',')','（','）','_','无症状感染者','密接','—','-',',',' ','副本','协查']

def formalize_filename(filename):
    filename = re.sub(r'[0-9]+', '', filename)
    filetype = filename.split('.')[-1]
    filename = filename.replace(filetype, '')
    filename = filename.replace('、', '和')

    # if filename[-1] != '告' or '明':
    #     if '报告' in filename:
    #         index = filename.find('告')
    #         flist = list(filename)[0:index + 2]
    #         filename = ''.join(flist)
    #     if '说明' in filename:
    #         index = filename.find('明')
    #         flist = list(filename)[index + 2:-1]
    #         filename = ''.join(flist)

    for i in factor_str:
        filename = filename.replace(i, '')

    if filename[1] != '于':
        filename = '关于' + filename
    if '的' not in filename:
        if '调查' in filename:
            char = '调查'
            filename = str_insert(filename, char)
        elif '流调' in filename:
            char = '流调'
            filename = str_insert(filename, char)
        elif '情况' in filename:
            char = '情况'
            filename = str_insert(filename, char)
        else:
            filename = filename + '的调查报告'

    filename = ''.join(filename) + '.' + filetype
    return filename