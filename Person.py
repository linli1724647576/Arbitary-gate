import openpyxl
import re
import math

'''
匹配逻辑
1、 同校区匹配  
2、 互相打分 a + b ( a + b - (b-a)*60)
'''

'''
0 序号
5 学历
6 姓名
7 性别
8 学院
9 感情状态
10 学历
11 社科or工科
12 学号
13 手机号
14 微信号
15 年龄
16 身高
17 校区
18 内向 or 外向

# @权重： 性格→生活与消费习惯→爱好→身高→年龄→专业→饮食
'''
class Person():
    def __init__(self,data):
        self.id = data[0].value
        self.sex = data[7].value  # 性别
        self.name = data[6].value # 姓名
        self.background = data[10].value # 学历 硕士 or 博士
        self.phone_number = data[13].value  # 电话号码
        self.wx = data[14].value  # 微信
        self.student_id = data[12].value  # 学号
        self.college = data[8].value # 学院
        self.personal = data[28].value   # 希望匹配到个人 Or 群
        self.adjust = data[29].value  # 是否调剂
        number = re.findall("\d+", data[16].value)
        self.height = [(int)(number[0]),(int)(number[-1])]  # 身高
        self.campus = data[17].value # 所在校区
        self.subject = data[11].value # 理工科Or社科
        self.to_or_eight = data[28].value #八人组or二人组

        self.personality1 = data[18].value # 性格 内向 or 外向
        self.personality2 = data[19].value # 性格 实际 or 感受
        self.personality3 = data[20].value # 性格 理智 or 感性
        self.personality4 = data[21].value # 性格 计划 or 灵活
        try:
            self.age = (float)(data[15].value) #年龄
        except Exception as e:
            self.age = 0

        self.hobbies = data[47].value.split("┋")

        self.match_age_range = [data[32].value,data[33].value]  #匹配年龄范围
        number = re.findall("\d+", data[34].value)
        self.match_height_range = [(int)(number[0]),(int)(number[-1])]  #匹配身高范围

        self.match_campus_range = data[30].value  # 匹配校区
        self.match_subject_range = data[35].value # 匹配学科(我都可以，可以优化）

        # 匹配性格
        if(type(data[37].value)==int):
            self.match_personality1 = data[37].value
        else:
            self.match_personality1 =  None
        if (type(data[38].value) == int):
            self.match_personality2 = data[38].value
        else:
            self.match_personality2 = None
        if (type(data[39].value) == int):
            self.match_personality3 = data[39].value
        else:
            self.match_personality3 = None
        if (type(data[39].value) == int):
            self.match_personality4 = data[39].value
        else:
            self.match_personality4 = None

        self.diet = data[42].value  #匹配饮食
        self.life = data[43].value  #匹配生活观念
        self.consume = data[44].value #匹配消费

        self.w = data[45].value.split("→") #权重


    def print_person(self):
        print("==================")
        print("id = ",self.id)
        print("sex = ", self.sex)
        print("name = ",self.name)
        print("background = ",self.background)
        print("phone_number = ", self.phone_number)
        print("wx = ", self.wx)
        print("student_id = ", self.student_id)
        print("college = ", self.college)
        # print("Professional_grade = ", self.Professional_grade)
        print("personal = ", self.personal)
        print("adjust = ", self.adjust)
        print("height = ", self.height)
        print("campus = ", self.campus)
        print("subject = ", self.subject)
        print("hobbies = ",self.hobbies)
        print("personality1 = ", self.personality1)
        print("personality2 = ", self.personality2)
        print("personality3 = ", self.personality3)
        print("personality4 = ", self.personality4)
        print("age = ",self.age)

        print("match_age_range = ",self.match_age_range)
        print("match_height_range = ",self.match_height_range)
        print("match_campus_range = ",self.match_campus_range)
        print("match_subject_range = ",self.match_subject_range)
        print("match_personality1 = ", self.match_personality1)
        print("match_personality2 = ", self.match_personality2)
        print("match_personality3 = ", self.match_personality3)
        print("match_personality4 = ", self.match_personality4)
        print("diet = ",self.diet)
        print("life = ", self.life)
        print("consume = ",self.consume)

        print(self.w)

    def get_score(self,Person):
        wt = [0,0,0,0,0,0,0]
        if("年龄" in self.w):
            wt[0] = (7 - self.w.index("年龄")) * 0.2
        else:
            wt[0] = 0.1

        if ("身高" in self.w):
            wt[1] = (7 - self.w.index("身高")) * 0.25
        else:
            wt[1] = 0.1

        if ("专业" in self.w):
            wt[2] = (7 - self.w.index("专业")) * 0.2
        else:
            wt[2] = 0.1

        if ("性格" in self.w):
            wt[3] = (7 - self.w.index("性格")) * 0.15
        else:
            wt[3] = 0.1

        if ("生活与消费习惯" in self.w):
            wt[4] = (7 - self.w.index("生活与消费习惯")) * 0.1
        else:
            wt[4] = 0.1

        if ("饮食" in self.w):
            wt[5] = (7 - self.w.index("饮食")) * 0.1
        else:
            wt[5] = 0.1

        if ("爱好" in self.w):
            wt[6] = (7 - self.w.index("爱好")) * 0.1
        else:
            wt[6] = 0.1

        # 排除跨校区
        if(not self.match_campus_range.__contains__(Person.campus)):
            return 0

        # 排除硕士和博士匹配
        if(self.background != Person.background):
            return 0

        # 年龄评分
        age_score = 0 if Person.age<(float)(self.match_age_range[0])or Person.age>(float)(self.match_age_range[1]) else 100
        # if(Person.age<(float)(self.match_age_range[0])or Person.age>(float)(self.match_age_range[1])):
        #     age_score = 0
        # else:
        #     age_score = 100

        # 身高评分
        height_score = 100 if max(Person.height[0], self.match_height_range[0]) < min(self.match_height_range[0], Person.height[1]) else 0
        # if (max(Person.height[0], Person.height[1]) < min(self.match_height_range[0], self.match_height_range[1])):
        #     height_score = 100
        # else:
        #     height_score = 0

        # 专业评分
        subject_score = 100 if self.match_subject_range == Person.subject else  0
        # if(self.match_subject_range == Person.subject):
        #     subject_score = 100
        # else:
        #     subject_score = 0

        # 性格评分
        if self.match_personality1 == None:
            personality_score1 = 25
        else:
            personality_score1 = 25 if self.match_personality1 == Person.personality1 else  0
        if self.match_personality2 == None:
            personality_score2 = 25
        else:
            personality_score2 = 25 if self.match_personality2 == Person.personality2 else  0
        if self.match_personality3 == None:
            personality_score3 = 25
        else:
            personality_score3 = 25 if self.match_personality3 == Person.personality3 else  0
        if self.match_personality4 == None:
            personality_score4 = 25
        else:
            personality_score4 = 25 if self.match_personality4 == Person.personality4 else  0
        personality_score = personality_score1 + personality_score2 + personality_score3 + personality_score4

        #生活与消费习惯评分呢
        life_score = 100 if self.life == Person.life else  0
        consume_score = 100 if self.consume == Person.consume else 0
        life_and_consume_score = life_score + consume_score

        # 饮食评分
        diet_score = 100 if self.diet == Person.diet else 0

        # 爱好评分
        habbit_score = 100
        for i in self.hobbies:
            for j in Person.hobbies:
                if(i==j):
                    habbit_score += 20

        score = age_score*wt[0] + height_score*wt[1] + subject_score*wt[2] + personality_score*wt[3] + life_and_consume_score*wt[4] + diet_score*wt[5] + habbit_score*wt[6]
        return (int)(score)



if __name__=="__main__":
    all_data = []
    excel = openpyxl.load_workbook(f'./2022_table.xlsx')
    sheet = excel.worksheets[0]
    for row in sheet.iter_rows():
        if row[0].value is None:
            break
        all_data.append(row)
    all_data = all_data[1:]
    person_list = []
    for data in all_data:
        p = Person(data)
        p.print_person()
        person_list.append(p)
    for i in range(len(person_list)):
        for j in range(i+1,len(person_list)):
            person_list[i].get_score(person_list[j])
