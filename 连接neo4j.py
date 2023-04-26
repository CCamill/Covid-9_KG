# 连接Neo4j数据库输入地址、用户名、密码
from py2neo import *
url = 'bolt://localhost:7687'
key = 'chang8677'
usr = 'neo4j'
graph = Graph(url,auth = (usr,key))
matcher = NodeMatcher(graph) #创建关系需要用到

nodes = matcher.match()
print(nodes)
for node in nodes:
    graph.delete(node)
# a = Node('owl_Class', name = '常金为') # 创建一个label = 'owl_Class'，属性name值为'banana'，此时节点并没有真的传到Neo4j
# b = Node('owl_Class', name = '牛市成')
# c = Node('owl_Class', name = '李侍寝')
# graph.create(a) # 将节点a创建到数据库
# graph.create(b)
# graph.create(c)

# 2.创建关系(已有节点)
# cRelation = matcher.match('ns1_Mammal',name = 'monkey').first()
# aRelation = matcher.match('owl_Class',name = 'banana').first()
# a_eat_c = Relationship(c,'eat',a)
# graph.create(a_eat_c) # 将创建传递到图上
# a_son_c = Relationship(a,'son',c)
# graph.create(a_son_c)

# node=matcher.match('owl_Class').where(name = '常金为').first() # 评估匹配并返回匹配的第一个节点
# node.setdefault('age',default='20')
# node['number']='001'  #添加“number”属性对应的属性值为“002”
# node['age'] = '20'
# node['home'] = '河南'
# node['sex'] = '男'
# print('node:')
# print(node) #打印节点信息
# graph.push(node)    #将更改或者添加的放到图中
#

# d = Node("owl_Class",name = 'zlm', age = '22' , sex = '男')
# graph.create(d)

# node = matcher.match('owl_Class').where(name = 'zlm').first()
# node['sex'] = '女'
# graph.push(node)

'''
根据ID删除节点
'''
# node = matcher.match('owl_Class').where( name = '常金为').first()
# print(node['name'])
# node_id = node.identity
# print(node_id)
# graph.delete(matcher[0])


# nodes=matcher.match('owl_Class').where(name = '李侍寝') #先匹配
# # graph.delete(node)  #a代表前面已经定义的结点
# print('节点长度',len(nodes))
# # print(help(node))
# print('节点是否存在',nodes.exists())
# print('第一个',nodes.first())
# for node in nodes:
#     if node.identity == 3:
#         print(node.identity, '已删除')
#         graph.delete(node)
#     print(node)



