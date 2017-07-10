# coding=UTF-8
from __future__ import division
from youtrack.connection import Connection, httplib2
import string
import sys
import csv
import codecs
from datetime import datetime

reload(sys)
sys.setdefaultencoding('utf-8')

f = open('sprint.txt')
# f = open('/data/sprint')
read = f.readlines()
sprint = read[1]
user = read[-2]
pwd = read[-1]

yt = Connection('https://server.propersoft.cn/youtrack', user, pwd)
print 'connected'

now = datetime.now()
date = now.strftime('%Y-%m-%d_%H%M%S')
# with open('/data/youtrack/'+ date + '_youtrack.csv', 'wb') as csvfile:
with open('../../'+ date + '_youtrack.csv', 'wb') as csvfile:
    csvfile.write(codecs.BOM_UTF8)
    spamwriter = csv.writer(csvfile, dialect='excel')
    spamwriter.writerow(['sprint'] + ['姓名'] + ['绩效点数'] + ['调整值']+ ['奖惩值'] + ['验证工时(分钟)']+ ['总工时(分钟)'] + ['计算明细'] + ['记录明细'])
    allissue = yt.getAllIssues('#' + sprint)
    verified = yt.getAllIssues('#Verified #' + sprint)
    statdic = {}
    class Counter(dict):
        def __missing__(self, key):
            return None
    c = Counter(statdic)
    detailDic = {}

    for issue in verified:
        if hasattr(issue, 'Estimation'):
            es = string.atoi(issue.Estimation)
        else:
            es = 0

        items = yt.getWorkItems(issue.id)

        ttdic = {}
        for item in items:
            if ttdic.has_key(item.authorLogin):
                ttdic[item.authorLogin] = ttdic[item.authorLogin] + string.atoi(item.duration)
            else:
                ttdic[item.authorLogin] = string.atoi(item.duration)
            des = item.description if hasattr(item, 'description') else ''
            timeStamp = string.atoi(item.date) / 1000
            dateArray = datetime.fromtimestamp(timeStamp)
            itemDate = dateArray.strftime("%Y-%m-%d")
            if detailDic.has_key(item.authorLogin):
                detailDic[item.authorLogin] += issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'
            else:
                detailDic[item.authorLogin] = issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'

        keys = ttdic.keys()
        if len(keys) == 0:
            des = issue.Assignee if hasattr(issue, 'Assignee') else ''
            if not c.has_key(des):
                c[des] = {}
            c[des][issue.id] = es
        elif len(keys) == 1:
            if not c.has_key(keys[0]):
                c[keys[0]] = {}
            c[keys[0]][issue.id] = es
        else:
            total = sum(ttdic.values())
            for k in keys:
                if not c.has_key(k):
                    c[k] = {}
                c[k][issue.id] = es * ttdic[k] / total

    sstatdic = {}
    sdetailDic = {}
    duration = {}
    for issue in allissue:
        if hasattr(issue, 'Estimation'):
            es = string.atoi(issue.Estimation)
        else:
            es = 0

        items = yt.getWorkItems(issue.id)

        sttdic = {}
        for item in items:
            if sttdic.has_key(item.authorLogin):
                sttdic[item.authorLogin] = sttdic[item.authorLogin] + string.atoi(item.duration)
            else:
                sttdic[item.authorLogin] = string.atoi(item.duration)
            des = item.description if hasattr(item, 'description') else ''
            timeStamp = string.atoi(item.date) / 1000
            dateArray = datetime.fromtimestamp(timeStamp)
            itemDate = dateArray.strftime("%Y-%m-%d %H:%M:%S")

            if duration.has_key(item.authorLogin):
                duration[item.authorLogin] += string.atoi(item.duration)
            else:
                duration[item.authorLogin] = 0

            if sdetailDic.has_key(item.authorLogin):
                sdetailDic[item.authorLogin] += issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'
            else:
                sdetailDic[item.authorLogin] = issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'

        skeys = sttdic.keys()
        if len(skeys) == 0:
            des = issue.Assignee if hasattr(issue, 'Assignee') else ''
            # if des == '':
            #     sstatdic = {0}
            if not sstatdic.has_key(des):
                sstatdic[des] = {}
            sstatdic[des][issue.id] = es
        elif len(skeys) == 1:
            if not sstatdic.has_key(skeys[0]):
                sstatdic[skeys[0]] = {}
            sstatdic[skeys[0]][issue.id] = es
        else:
            total = sum(sttdic.values())
            for k in skeys:
                if not sstatdic.has_key(k):
                    sstatdic[k] = {}
                sstatdic[k][issue.id] = es * sttdic[k] / total


    for assignee in sstatdic:
        if assignee == '':
            continue
        s = 0
        ss = 0
        list = []
        tasks = ''
        for issue in sstatdic[assignee].keys():
            tt = duration[assignee] if duration.has_key(assignee) else '0'
        ssh = tt
        des = c[assignee].keys() if (c.has_key(assignee)) else ''
        for issue in des:
            list.append('%s%s%i%s'% (issue , ':' , (c[assignee][issue]),', '))
            tasks = ''.join(list)
            s += c[assignee][issue]
        sd = str(round(s / 60 / 8, 2))
        sh = str(round(s, 2))
        detail = detailDic[assignee] if detailDic.has_key(assignee) else ''
        spamwriter.writerow([sprint,assignee,sd,'','',sh, ssh, tasks, detail])