# coding=UTF-8
from __future__ import division
from youtrack.connection import Connection, httplib2
import string
import sys
import csv
import ConfigParser
import codecs
import email,os
import smtplib
from email.MIMEText import MIMEText
from email.MIMEMultipart import MIMEMultipart
from datetime import datetime

reload(sys)
sys.setdefaultencoding('utf-8')

# sprintpath = "/data/sprint"
sprintpath = "sprint.txt"
conf = ConfigParser.ConfigParser()
conf.read(sprintpath)
sprints = []
mail_devs = []
sprint_dev = conf.get("conf", "sprint_dev")
sprint_test = conf.get("conf", "sprint_test")
date_dev = conf.get("conf", "date_dev")
date_test = conf.get("conf", "date_test")
mail_dev = conf.get("conf", "mail_dev")
mail_test = conf.get("conf", "mail_test")
cc = conf.get("conf", "cc")
smtp_user = conf.get("conf", "smtp_user")
smtp_pwd = conf.get("conf", "smtp_pwd")

# mail_devs.append(mail_dev)
if datetime.now().strftime('%Y-%m-%d') == date_dev:
    sprints.append(sprint_dev)

if datetime.now().strftime('%Y-%m-%d') == date_test:
    sprints.append(sprint_test)

print sprints


user = conf.get("conf", "user")
pwd = conf.get("conf", "pwd")
youtrackpath = conf.get("conf", "path")
yt = Connection('https://cloud.propersoft.cn/youtrack', user, pwd)
print 'connected'
for sprint in sprints:

    now = datetime.now()
    date = now.strftime('%Y-%m-%d %H:%M:%S')
    # with open(youtrackpath + sprint +'.csv', 'wb') as csvfile:
    with open(youtrackpath + sprint + '.csv', 'wb') as csvfile:
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
                    duration[item.authorLogin] = string.atoi(item.duration)

                if sdetailDic.has_key(item.authorLogin):
                    sdetailDic[item.authorLogin] += issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'
                else:
                    sdetailDic[item.authorLogin] = issue.id + ' ' + itemDate + ' ' + des + ' ' + item.duration + ';'

            skeys = sttdic.keys()
            if len(skeys) == 0:
                des = issue.Assignee if hasattr(issue, 'Assignee') else ''
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


    # def sendMail(subject, receivers, cc, content, atts):
    def sendMail(subject, receivers, cc, content, atts):
        SENDER = 'gdpr@propersoft.cn'
        msg = MIMEMultipart('related')
        msg['Subject'] = unicode(subject, "UTF-8")
        msg['From'] = SENDER
        msg['To'] = receivers
        msg['Cc'] = cc

        # 邮件内容
        if os.path.isfile(content):
            if (content.split('.')[-1] == 'html'):
                cont = MIMEText(open(content).read(), 'html', 'utf-8')
            else:
                cont = MIMEText(open(content).read(), 'plain', 'utf-8')
        else:
            cont = MIMEText(content, 'plain', 'utf-8')
        msg.attach(cont)

        # 构造附件
        if atts != -1 and atts != '':
            for att in atts.split(','):
                os.path.isfile(att)
                name = os.path.basename(att)
                att = MIMEText(open(att).read(), 'base64', 'utf-8')
                att["Content-Type"] = 'application/octet-stream'
                # 将编码方式为utf-8的name，转码为unicode，然后再转成gbk(否则，附件带中文名的话会出现乱码)
                att["Content-Disposition"] = 'attachment; filename=%s' % name.decode('utf-8').encode('gbk')
                msg.attach(att)
        smtp = smtplib.SMTP_SSL('smtp.exmail.qq.com', port=465)
        smtp.login(smtp_user, smtp_pwd)
        for recev in receivers.split(','):
            smtp.sendmail(SENDER, recev, msg.as_string())
        if cc != '':
            for c in cc.split(','):
                smtp.sendmail(SENDER, c, msg.as_string())
        smtp.close()


    # python performance-statistic.py 2017-03-26 xxx xxx '#core-1703a' core-1703a 'hexin@propersoft.cn, gdpr@propersoft.cn'
def main(args):
    # execdate = date_dev
    if datetime.now().strftime('%Y-%m-%d') == date_dev:
        receivers = mail_dev
        sendMail(sprint_dev, receivers, cc, '执行时间: '+ date, youtrackpath + sprint_dev + '.csv')
    else:
        print sprint_dev + ' Not at execute date:'
        print args

    if datetime.now().strftime('%Y-%m-%d') == date_test:
        receivers = mail_test
        sendMail(sprint_test, receivers, cc, '执行时间: '+ date, youtrackpath + sprint_test + '.csv')
    else:
        print sprint_test + ' Not at execute date:'
        print args

if __name__ == '__main__':
    main(sys.argv)