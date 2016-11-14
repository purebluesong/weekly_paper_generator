# coding=utf-8
from docx_convert import FinalPaperWeekPlan
import json


def convert(filename):
    with open(filename, 'r+') as f:
        data = json.load(f, encoding='utf-8')
        config = data['config']
        datas = data['data']
        for index, data in enumerate(datas):
            data['week'] = u"第 %02d 周" % (index+1)
            data['duration'] = u"%s 到 %s" % (data.pop('start_time'), data.get('report_time'))
            data['content'] = '    ' + data.get('content') or ''
            data['summary'] = '    ' + data.get('summary') or ''
            data['nextweekplan'] = '    ' + data.get('nextweekplan') or ''
            data['problem'] = '    ' + data.get('problem') or ''
        return datas, config


def make(start=0):
    paper = FinalPaperWeekPlan()
    datas, config = convert('weeklyplan.json')
    for data in datas[start:]:
        data.update(config)
        paper.convert(
            data=data,
            destnation='weekreports/',
            titleparams=(data.get('name'), data.get('week'))
        )


if __name__ == '__main__':
    make()
