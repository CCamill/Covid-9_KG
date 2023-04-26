import os
from format_filename import formalize_filename
for date in os.listdir(r'D:\桌面\knowledge extract\document'):
    path_date = os.path.join(r'D:\桌面\knowledge extract\document' , date)
    for pation in [i for i in os.listdir(path_date) if i.split(".")[-1] == "docx"]:
        pation_path = os.path.join(path_date,pation)
        new_name = os.path.join(path_date,formalize_filename(pation))
        if not os.path.exists(new_name):
            print('正在处理' ,pation_path)
            os.renames(pation_path,new_name)
        else:
            print(new_name,'已存在')