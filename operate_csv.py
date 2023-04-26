import csv


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

path = r'csvfile/11.21/关于刘青选的密接登记表.csv'

for i in csv_to_reader(path):
    print(i)
print(csv_to_reader(path)[0][0])