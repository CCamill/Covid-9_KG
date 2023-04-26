import datetime#导入时间
area={
    '11':'北京市','12':'天津市','13':'河北省','14':'山西省','15':'内蒙古自治区',
    '21':'辽宁省','22':'吉林省','23':'黑龙江省',
    '31':'上海市','32':'江苏省','33':'浙江省','34':'安徽省','35':'福建省','36':'江西省','37':'山东省',
    '41':'河南省','42':'湖北省','43':'湖南省',
    '44':'广东省','45':'广西壮族自治区','46':'海南省',
    '50':'重庆市','51':'四川省','52':'贵州省','53':'云南省','54':'西藏自治区',
    '61':'陕西省','62':'甘肃省','63':'青海省','64':'宁夏回族自治区','65':'新疆维吾尔族自治区',
    '81':'香港特别行政区','82':'澳门特别行政区','83':'台湾地区'
}#全国省的字典,全局变量
factorArr=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]#17位相乘系数数组
judgeLastNumArr=[1,0,'X',9,8,7,6,5,4,3,2]#%11对应的余数得到的校验码，下标为余数
class IdcardValidator():
    def __init__(self,number):#一个参数，构造函数
        self.number=str(number)#参数为身份证号码数组
        # self.n_flag=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]#17位相乘数组
    def getLocation(self):#地区切片
        if self.isTruth() is False:
            return False
        loc=self.number[0:2]#前2位
        return area[loc]
    def getSex(self):#性别
        if self.isTruth() is False:
            return False
        sex=self.number[-2]#倒数第二位
        sex=int(sex)
        if sex%2:
            return '男'# 男
        return '女' # 女
    def getAge(self):
        if self.isTruth() is False:
            return False
        birthday=self.number[6:14]#出生年月日

        birthYear=birthday[0:4]#前四位
        birthMonth=birthday[4:6]
        birthDate=birthday[6:8]
        now=datetime.datetime.now()
        diffYear=now.year-int(birthYear)#int换算
        diffMonth=now.month-int(birthMonth)
        diffDay=now.day-int(birthDate)
        if diffMonth<0 or diffMonth==0 and diffDay<0:
            return diffYear -1
        return diffYear
    def isTruth(self):#辨别真假
        if len(self.number) is not 18:
            return False
        num17 = self.number[0:17]
        last_num = self.number[-1]  # 截取前17位和最后一位
        moduls = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
        num17 = map(int, num17)
        num_tuple = zip(num17, moduls)  # [(1, 4), (2, 5), (3, 6)]
        num = map(lambda x: x[0] * x[1], num_tuple)
        mod = sum(num) % 11
        remainder = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        factorArr = [1, 0, 'X', 9, 8, 7, 6, 5, 4, 3, 2]
        last_remainder = dict(zip(remainder, factorArr))
        if last_num == str(last_remainder[mod]):
            return True
        else:
            return False

def get_information_by_id(id):
    demo=IdcardValidator(id)
    return demo
