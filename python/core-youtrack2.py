# coding=UTF-8

from __future__ import division
from youtrack.connection import Connection, httplib2
import httplib
import string

# httplib2.debuglevel=4
# httplib.debuglevel=4

user = 'wangwei'
pwd = 'oaABC123!'

f = open('sprint.txt')
read = f.readlines()
sprint = read[-1]

yt = Connection('https://server.propersoft.cn/youtrack', user, pwd)
print 'connected'

verified = yt.getAllIssues('#Verified #' + sprint)
statdic = {}
for issue in verified:
    if hasattr(issue, 'Estimation'):
        es = string.atoi(issue.Estimation)
    else:
        es = 0
    # print issue.id + ':' + issue.Assignee + ',' + str(es)

    items = yt.getWorkItems(issue.id)
    ttdic = {}
    for item in items:
        if ttdic.has_key(item.authorLogin):
            ttdic[item.authorLogin] = ttdic[item.authorLogin] + string.atoi(item.duration)
        else:
            ttdic[item.authorLogin] = string.atoi(item.duration)

    keys = ttdic.keys()
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

# print '========'
for assignee in statdic:
    print assignee + ':'
    sum = 0
    for issue in statdic[assignee].keys():
        print '\t' + issue + ':' + str(statdic[assignee][issue])
        # print statdic[assignee][issue]
        sum += statdic[assignee][issue]
    print '\tSUM:' + str(round(sum/60/8,2)) + ' (' + str(round(sum,2)) + ')'
