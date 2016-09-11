# -*- coding: utf-8 -*-
'''
说明：
* 脚本仅计算了智育学分
* 将第20行代码换成自己的CET-4、CET-6成绩
* 公选超过25学分的"神人"自己看着办吧。

'''

import xdrlib
import xlrd
import sys

dst = {}

dst[u'优秀'] = 95.0
dst[u'良好'] = 85.0
dst[u'中等'] = 75.0
dst[u'及格'] = 65.0

# liushiyu cet4_and_6_score = 591 + 577 # CET 4 + CET 6
# yuanqi cet4_and_6_score = 494 + 0 # CET 4 + CET 6
# dairui cet4_and_6_score = 525 + 447 # CET 4 + CET 6
# cet4_and_6_score = 439 + 451 # suzhaoxin
cet4_and_6_score = 425 + 483 # zhou guang yuan

if __name__ == "__main__":
    filename = sys.argv[1]
    # list结构{"课程名称": ["成绩"， "学分", "成绩*学分", "课程类型"}
    obligatory = {} # 必修
    elective_private = {} # 限选
    elective_public  = {} # 公选

    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]

    n_rows = table.nrows

    for row in range(3, n_rows ):
        _name = table.cell(row, 3).value
        _score = table.cell(row, 4).value
        if _score not in dst.keys():
            _score = float(table.cell(row, 4).value)
        else:
            _score = dst[_score]
        _credit = float(table.cell(row, 9).value)
        _type = table.cell(row, 7).value
        
        '''
        if _score == 0.0: # 成绩为0的不算
            continue 
        '''

        if _type == u'必修':
            obligatory[_name] = [ _score, _credit, _score * _credit, _type ]
        elif _type == u'限选':
            elective_private[_name] = [ _score, _credit, _score * _credit, _type ]
        elif _type == u'任选':
            elective_public[_name] = [ _score, _credit, _score * _credit, _type ]

    obligatory_scores = [ value[2] for value in obligatory.values() ]
    obligatory_credits = [ value[1] for value in obligatory.values() ]
    elective_private_scores = [ value[2] for value in elective_private.values() ]
    elective_private_credits = [ value[1] for value in elective_private.values() ]
    elective_public_scores = [ value[2] for value in elective_public.values() ]
    elective_public_credits = [ value[1] for value in elective_public.values()]

    print "\n必修：\n"
    for item in obligatory:
        print ("%s  %d  %d" % (item, obligatory[item][0], obligatory[item][1]))
    print "\n限选：\n"
    for item in elective_private:
        print ("%s  %d  %d" % (item, elective_private[item][0], elective_private[item][1]))
    print "\n公选：\n"
    for item in elective_public:
        print ("%s  %d  %d" % (item, elective_public[item][0], elective_public[item][1]))

    obligatory_plus_elective_private = (sum(obligatory_scores) + sum(elective_private_scores)) / (sum(obligatory_credits) + sum(elective_private_credits))
    elective_public_ = sum(elective_public_scores) * 0.0005
    cet4_and_6 = cet4_and_6_score / 710.0 * 100.0 * 0.0005 * 2

    print "\n必修课：" + str(sum(obligatory_scores) / sum(obligatory_credits)) + \
            " 学分：" + str(sum(obligatory_credits)) + \
            " 总分：" + str(sum(obligatory_scores))

    print "\n限选课：" + str(sum(elective_private_scores) / sum(elective_private_credits)) + \
            " 学分：" + str(sum(elective_private_credits)) + \
            " 总分：" + str(sum(elective_private_scores))

    print "\n必修课+限选课：" + str(obligatory_plus_elective_private) + \
            " 学分："+ str(sum(obligatory_credits) + sum(elective_private_credits)) + \
            " 总分：" + str(sum(obligatory_scores) + sum(elective_private_scores))

    print "\n公选课：" + str(elective_public_) + " 学分：" + str(sum(elective_public_credits)) + " 总分：" + str(sum(elective_public_scores))

    print "\n四六级: " + str(cet4_and_6) + " (" +str(cet4_and_6_score) + ") 学分：4"
    print "\n四六级+公选：" + str(elective_public_ + cet4_and_6) + " 学分：" + str(sum(elective_public_credits) + 4)

    final_scores = obligatory_plus_elective_private + elective_public_ + cet4_and_6 # 四级

    print "\n智育学分：" + str(final_scores)


'''
科技加分
'''
