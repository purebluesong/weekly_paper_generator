# coding=utf-8
from docx_convert import FinalPaperWeekPlan
import json


def convert(filename):
    with open(filename, 'r+') as f:
        datas = json.load(f, encoding='utf-8')['data']
        for index, data in enumerate(datas):
            data['week'] = u"第 %02d 周" % (index+1)
            data['duration'] = u"%s 到 %s" % (data.pop('start_time'), data.get('report_time'))
            data['content'] = '    ' + data.get('content') or ''
            data['summary'] = '    ' + data.get('summary') or ''
            data['nextweekplan'] = '    ' + data.get('nextweekplan') or ''
            data['problem'] = '    ' + data.get('problem') or ''
        return datas


def make(start=0):
    paper = FinalPaperWeekPlan()
    base_data = {
        "name": u'张瑞鹏',
        "tel": '15754605351',
        "email": 'purebluesong@gmail.com',
        "base": u"上海谷露软件",
        "mentor": u"方雷",
        "inner_mentor": u"关毅",
    }

    for data in convert('weeklyplan.json')[start:]:
        data.update(base_data)
        paper.convert(
            data=data,
            destnation='weekreports/',
            titleparams=(data.get('name'), data.get('week'))
        )


if __name__ == '__main__':
    make()
