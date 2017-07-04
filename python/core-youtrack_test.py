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
#print sys.getdefaultencoding()
# httplib2.debuglevel=4
# httplib.debuglevel=4

user = 'wangwei'
pwd = 'oaABC123!'
f = open('sprint.txt')
read = f.readlines()
sprint = read[-1]

yt = Connection('https://server.propersoft.cn/youtrack', user, pwd)
print 'connected'
# print '========'
# for assignee in statdic:
#     print assignee + ':'
#     sum = 0
#     for issue in statdic[assignee].keys():
#         print '\t' + issue + ':' + str(statdic[assignee][issue])
#         # print statdic[assignee][issue]
#         sum += statdic[assignee][issue]
#     print '\tSUM:' + str(round(sum/60/8,2)) + ' (' + str(round(sum,2)) + ')'
# i = []
now = datetime.now()
date = now.strftime('%Y-%m-%d_%H%M%S')
with open('../../'+ date + '_yourack.csv', 'wb') as csvfile:
    csvfile.write(codecs.BOM_UTF8)
    spamwriter = csv.writer(csvfile, dialect='excel')
    spamwriter.writerow(['sprint'] + ['姓名'] + ['绩效点数'] + ['调整值']+ ['奖惩值'] + ['验证工时(分钟)']+ ['总工时(分钟)'] + ['计算明细'] + ['记录明细'])
    allissue = yt.getAllIssues('#' + sprint)
    # statdic = {}
    # print allissue
    verified = yt.getAllIssues('#Verified #' + sprint)
    statdic = {}
    detailDic = {}
    # print verified
    # for issue in allissue:
    for issue in verified:
        if hasattr(issue, 'Estimation'):
            es = string.atoi(issue.Estimation)
        else:
            es = 0
        # print issue.id + ':' + issue.Assignee + ',' + str(es)

        items = yt.getWorkItems(issue.id)
        # items = yt.getWorkItems(issue)
        ttdic = {}
        for item in items:
            if ttdic.has_key(item.authorLogin):
                ttdic[item.authorLogin] = ttdic[item.authorLogin] + string.atoi(item.duration)
            else:
                ttdic[item.authorLogin] = string.atoi(item.duration)

            des = item.description if hasattr(item, 'description') else ''
            timeStamp = string.atoi(item.date) / 1000
            dateArray = datetime.fromtimestamp(timeStamp)
            itemDate = dateArray.strftime("%Y-%m-%d %H:%M:%S")
            if detailDic.has_key(item.authorLogin):
                # detail[item.authorLogin] += item.authorLogin + ' ' + item.duration + ' ' + item.url + ' ' + item.description
                detailDic[item.authorLogin] += issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'
            else:
                detailDic[item.authorLogin] = issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'
        # dj = ''.join(ttdic)
        # i.append(item)
        # print  i
        # print  ttdic
        keys = ttdic.keys()
        # print i
        # for k in keys:
        #     print '\t' + k + ',' + str(ttdic[k])
        if len(keys) == 0:
            if not statdic.has_key(issue.Assignee):
                statdic[issue.Assignee] = {}
            statdic[issue.Assignee][issue.id] = es
        elif len(keys) == 1:
            if not statdic.has_key(keys[0]):
                statdic[keys[0]] = {}
            statdic[keys[0]][issue.id] = es
        else:
            total = sum(ttdic.values())
            for k in keys:
                if not statdic.has_key(k):
                    statdic[k] = {}
                statdic[k][issue.id] = es * ttdic[k] / total

    # i = []
    for assignee in statdic:
        s = 0
    # print assignee
    #     items = yt.getWorkItems(issue)
    #     print items
    #     if assignee == item.authorLogin:
            # i.append(items)
            # print i
        # print '#' * 100
        # spamwriter.writerow([assignee])
        list = []
        # allitem = []
        for issue in statdic[assignee].keys():
            list.append('%s%s%i%s'% (issue , ':' , (statdic[assignee][issue]),', '))
            task = ''.join(list)
            s += statdic[assignee][issue]
        sd = str(round(s / 60 / 8, 2))
        sh = str(round(s, 2))
        detail = detailDic[assignee] if detailDic.has_key(assignee) else ''
        spamwriter.writerow([sprint,assignee,sd,'','',sh, '', task, detail])
        # print i
        # print '#' * 100