from py2neo import *

url = 'bolt://localhost:7687'
key = 'chang8677'
usr = 'neo4j'
graph = Graph(url, auth=(usr, key))
matcher = NodeMatcher(graph)  # 创建关系需要用到
rmatcher = RelationshipMatcher(graph)

name = '孙绵雨'
date = '12.04'
cyber = r"match (n:`{}`)<-[r*0..]-(b) where n.name = '{}' return n,b".format(date,name)

nodes = graph.run(cyber).data()
rel = []
for node in nodes:
    r = rmatcher.match({node['n'],node['b']}).first()
    rel.append(r)
for r in list(set(rel)):
    print(r)
