'''
将csv中的人物关系输出到neo4j图数据库中
'''
from get_information import *
from py2neo import *
from doc2csv import *


class config:
    docxCatalog = r'document'
    csvCatalog = r'csvfile'
    docCatalog = r'D:\桌面\covid_document'

url = 'bolt://localhost:7687'
key = 'chang8677'
usr = 'neo4j'
graph = Graph(url, auth=(usr, key))
matcher = NodeMatcher(graph)  # 创建关系需要用到

def csv_to_reader(doc_path):
    with open(doc_path, "r") as csvfile:
        reader = csv.reader(csvfile)
        rows = [row for row in reader]
        return rows

def csv_to_dictrreader(doc_path):
    with open(doc_path, "r") as csvfile:
        reader = csv.DictReader(csvfile)
        rows = [row for row in reader]
        return rows
def creat_relation(a_have,b_have,row,properties,keys,flag):
    if a_have and b_have:
        if row[flag[0][-1]] != '':
            rel_a = Relationship(a_have, row[flag[0][-1]], b_have, **properties)
            graph.create(rel_a)
            print('关系创建成功')
        if row[flag[0][-1]] == '' or row[flag[0][-1]] == '-':
            rel_a = Relationship(a_have, '未知', b_have, **properties)
            graph.create(rel_a)
            print('关系创建成功')
        #
        # if '备注' in keys:
        #     if row['备注'] != '':
        #         rel_a = Relationship(a_have, row['备注'], b_have, **properties)
        #         graph.create(rel_a)
        #         print('关系创建成功')
        #     if row['备注'] == '' or row['备注'] == '-':
        #         rel_a = Relationship(a_have, '未注明', b_have, **properties)
        #         graph.create(rel_a)
        #         print('关系创建成功')
        # elif '备注（关系或者接触方式）' in keys:
        #     if '备注（关系或者接触方式）' in keys:
        #         if row['备注（关系或者接触方式）'] != '':
        #             rel_a = Relationship(a_have, row['备注（关系或者接触方式）'], b_have, **properties)
        #             graph.create(rel_a)
        #             print('关系创建成功')
        #         if row['备注（关系或者接触方式）'] == '' or row['备注（关系或者接触方式）'] == '-':
        #             rel_a = Relationship(a_have, '未注明', b_have, **properties)
        #             graph.create(rel_a)
        #             print('关系创建成功')
        else:
            print('不存在关系')

def main(csv_path,docx_path):
    print('当前文档路径',csv_path,docx_path)
    patient_inf = get_information(docx_path)
    label = date
    patient = Node(label, name=patient_inf['姓名'],
                   type='接受调查者',
                   sex=patient_inf['性别'],
                   age=patient_inf['年龄'],
                   id=patient_inf['身份证号'],
                   phonenumber=patient_inf['联系电话'],
                   HomeAddress = patient_inf['家庭住址'],
                   health_state= patient_inf['健康状况'],
                   NAT = patient_inf['核酸检测结果'])

    if not patient_inf['身份证号'] and not matcher.match(label).where(id=patient_inf['身份证号']).first():
        graph.create(patient)
        print(patient['name'], '节点创建成功')
    elif not matcher.match(label).where(name=patient_inf['姓名']).first():
        graph.create(patient)
        print(patient['name'], '节点创建成功')
    else:
        patient = graph.nodes.match(label, name=patient_inf['姓名']).first()
        patient['type'] = '接受调查者'
        patient['NAT'] = patient_inf['核酸检测结果']
        patient['health_state'] = patient_inf['健康状况']
        patient['age'] = patient_inf['年龄']
        patient['HomeAddress'] = patient_inf['家庭住址']
        patient['id'] = patient_inf['身份证号']
        graph.push(patient)
        print('type 修改成功')
        print(patient['name'], '节点已存在')

    flag = csv_to_reader(csv_path)
    if len(flag) == 0:
        print('文档 关于{}的调查报告.docx 不存在表格'.format(patient_inf['姓名']))

    elif flag[0][0] == '排序' or flag[0][0] == '序号':
        rows = csv_to_dictrreader(csv_path)
        if len(rows) == 0:
            patient['other'] = '无密接信息'
            graph.push(patient)
        elif rows[0]['姓名'] == '':
            patient['other'] = '无密接信息'
            graph.push(patient)
        else:
            for row in rows:
                keys = list(row.keys())
                if row['姓名'] != '':
                    if matcher.match(label).where(name=row['姓名'],type = '接受调查者').first():
                        print(row['姓名'],'节点已存在')
                    else:
                        node = Node(label, name=row['姓名'],type='密接人员',)
                        #keys.remove('排序')
                        if '性别' in keys and row['性别'] != '':
                            node['sex'] = row['性别']
                            graph.push(node)

                        if '联系方式' in keys and row['联系方式'] != '':
                            node['phonenumber'] = row['联系方式']
                            graph.push(node)
                        elif '手机号' in keys and row['手机号'] != '':
                            node['phonenumber'] = row['手机号']
                            graph.push(node)

                        if '现住址' in keys and row['现住址'] != '':
                            node['present_address'] = row['现住址']
                            graph.push(node)
                        elif '家 庭 住 址' in keys and row['家 庭 住 址'] != '':
                            node['present_address'] = row['家 庭 住 址']
                            graph.push(node)

                        if not matcher.match(label).where(name=row['姓名']).first():
                            graph.create(node)
                            print(node['name'], '节点创建成功')
                        else:
                            print(node['name'], '节点已存在')

                    a_have = graph.nodes.match(label, name=patient_inf['姓名']).first()
                    b_have = graph.nodes.match(label, name=row['姓名']).first()
                    properties = {'类别': '密接', '日期': date}

                    creat_relation(a_have, b_have, row, properties, keys,flag)

    elif flag[0][0] != '排序' and flag[0][0] != '序号':
        rows = csv_to_reader(csv_path)
        if len(rows) == 0:
            patient['other'] = '无密接信息'
            graph.push(patient)
        elif rows[0][0] == '':
            patient['other'] = '无密接信息'
            graph.push(patient)
        else:
            for row in rows:
                if len(row) <= 1:
                    pass
                elif row[1] == '':
                    pass
                else:
                    node = Node(label, name=row[1])
                    if row[2] =='':
                        node['sex'] = '未知'
                        graph.push(node)
                    else:
                        node['sex'] = row[2]
                        graph.push(node)
                    if row[3] =='':
                        node['phonenumber'] = '未知'
                        graph.push(node)
                    else:
                        node['phonenumber'] = row[3]
                        graph.push(node)
                    if row[4] =='':
                        node['present_address'] = '未知'
                        graph.push(node)
                    else:
                        node['present_address'] = row[4]
                        graph.push(node)

                    if not matcher.match(label).where(name=row[1]).first():
                        graph.create(node)
                        print(node['name'], '节点创建成功')
                    else:
                        print(node['name'], '节点已存在')

                    a_have = graph.nodes.match(label, name=patient_inf['姓名']).first()
                    b_have = graph.nodes.match(label, name=row[1]).first()
                    if a_have and b_have:
                        if row[-1] != '':
                            rel_a = Relationship(a_have, row[-1], b_have)
                            graph.create(rel_a)
                            print('关系创建成功')
                        else:
                            print('不存在关系')


if __name__ == "__main__":
    date_paths = os.listdir(config.docxCatalog)

    for date in date_paths:
        csv_date_path = os.path.join(config.csvCatalog, date)
        docx_date_path = os.path.join(config.docxCatalog, date)
        patients = [i for i in os.listdir(docx_date_path) if str(i).endswith('docx') and '$' not in str(i)]
        for patient in patients:
            patient_name = get_name(patient)
            csv_path = os.path.join(csv_date_path, '关于{}的密接登记表.csv'.format(patient_name))
            docx_path = os.path.join(docx_date_path, '关于{}的调查报告.docx'.format(patient_name))
            if os.path.exists(docx_path) and os.path.exists(csv_path):
                main(csv_path,docx_path)
            else:
                print('文档 关于{}的调查报告.docx 不存在表格'.format(patient_name))