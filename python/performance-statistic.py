# coding=UTF-8

from __future__ import division
from __future__ import print_function
from __future__ import print_function

import codecs
import csv
import os
import smtplib
import string
import sys
from datetime import datetime

from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

from youtrack.connection import Connection

reload(sys)
sys.setdefaultencoding('utf-8')

now = datetime.now()
date = now.strftime('%Y-%m-%d %H:%M:%S')
path = 'E:\Python\performance/'

# 任务类型因子
factor_type_design = 1.5
factor_type_implement = 1.0
factor_type_check = 0.5
# 效率因子
factor_efficiency_delays = 0.1
# 质量因子
factor_quality_rejects = 0.1


def create_result_dic_element():
    return {
        'performance_result': 0,  # 绩效结果
        # 'verified_issues': [],      # 所有已验证任务集合
        # 'all_issues': [],           # 所有任务集合
        'spend_times': 0,  # 总工时
        'compute_details': '',  # 计算明细
        'all_records': ''  # 记录明细
    }


def compute_performance_result(estimation, issue_type, delays, rejects):
    delays = float(delays)
    rejects = float(rejects)
    if issue_type == 'Design':
        factor_type = factor_type_design
    elif issue_type == 'Implement':
        factor_type = factor_type_implement
    else:
        factor_type = factor_type_check

    if delays > 0:
        factor_efficiency = 1.0 - delays * factor_efficiency_delays
        factor_efficiency = factor_efficiency if factor_efficiency > 0 else 0
    else:
        factor_efficiency = 1.0

    if rejects > 0:
        factor_quality = 1.0 - rejects * factor_quality_rejects
        factor_quality = factor_quality if factor_quality > 0 else 0
    else:
        factor_quality = 1.0

    result = estimation * factor_type * factor_efficiency * factor_quality
    return result


def fill_performance_result(dic, issue_id, estimation, issue_type, delays, rejects):
    performance_result = compute_performance_result(estimation, issue_type, delays, rejects)
    dic['performance_result'] += performance_result
    dic['compute_details'] += issue_id + ':' + str(performance_result) \
        + '(estimation: ' + str(estimation) + ', type: ' + issue_type + ', delays: ' \
        + str(delays) + ', rejects: ' + str(rejects) + '),'


def get_performance_result_dic(user, pwd, query, sprint):
    yt = Connection('https://cloud.propersoft.cn/youtrack', user, pwd)
    print('connected')
    print(sprint)

    # 每人绩效结果字典，key 为用户名，value 为 create_result_dic_element 方法创建的属性字典。
    result_dic = {}

    all_issues = yt.getAllIssues(query)
    for issue in all_issues:
        if not hasattr(issue, 'Assignee'):
            continue

        # 任务责任人
        assignee = issue.Assignee
        if assignee not in result_dic:
            result_dic[assignee] = create_result_dic_element()

        # 评估工时
        if hasattr(issue, 'Estimation'):
            estimation = string.atoi(issue.Estimation)
        else:
            estimation = 0

        issue_type = issue.Type
        delays = issue.Delays if hasattr(issue, 'Delays') else 0
        rejects = issue.Rejects if hasattr(issue, 'Rejects') else 0

        items = yt.getWorkItems(issue.id)
        # 每个任务下每人实际工时字典
        author_duration_dic = {}
        for item in items:
            if item.authorLogin not in result_dic:
                result_dic[item.authorLogin] = create_result_dic_element()

            result_dic[item.authorLogin]['spend_times'] += string.atoi(item.duration)

            desc = item.description if hasattr(item, 'description') else ''
            time_stamp = string.atoi(item.date) / 1000
            date_array = datetime.fromtimestamp(time_stamp)
            item_date = date_array.strftime("%Y-%m-%d")
            result_dic[item.authorLogin][
                'all_records'] += issue.id + ' ' + item_date + ' ' + desc + ' ' + item.duration + ';'

            author_duration_dic[item.authorLogin] = string.atoi(item.duration)

        if issue.State == 'Verified':
            authors = author_duration_dic.keys()
            if len(authors) == 0:
                # 任务未填写实际工时，工作量直接计算到任务负责人头上
                fill_performance_result(result_dic[assignee], issue.id, estimation, issue_type, delays, rejects)
            elif len(authors) == 1:
                # 仅一人填写实际工时，工作量计算到填写人头上
                fill_performance_result(result_dic[authors[0]], issue.id, estimation, issue_type, delays, rejects)
            else:
                # 多人协作任务工作量分配
                total = sum(author_duration_dic.values())
                for author in authors:
                    fill_performance_result(result_dic[author], issue.id,
                                            estimation * author_duration_dic[author] / total,
                                            issue_type, delays, rejects)

    return result_dic


def export_csv(sprint, result_dic):
    with open(path + sprint + '.csv', 'wb') as csv_file:
        csv_file.write(codecs.BOM_UTF8)
        spam_writer = csv.writer(csv_file)
        spam_writer.writerow(['sprint'] + ['姓名'] + ['绩效点数'] + ['总工时(分钟)'] + ['计算明细'] + ['记录明细'])

        for assignee in result_dic.keys():
            performance_result = str(round(result_dic[assignee]['performance_result'] / 8 / 60, 2))
            spend_times = str(result_dic[assignee]['spend_times'])
            compute_details = result_dic[assignee]['compute_details']
            all_records = result_dic[assignee]['all_records']
            spam_writer.writerow([sprint, assignee, performance_result, spend_times, compute_details, all_records])


def send_mail(subject, receivers, cc, content, atts):
    sender = 'gdpr@propersoft.cn'
    msg = MIMEMultipart('related')
    msg['Subject'] = unicode(subject, "UTF-8")
    msg['From'] = sender
    msg['To'] = receivers
    msg['Cc'] = cc

    # 邮件内容
    if os.path.isfile(content):
        if content.split('.')[-1] == 'html':
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
    smtp.login('gdpr@propersoft.cn', '31353260Gdpr')
    for recev in receivers.split(','):
        smtp.sendmail(sender, recev, msg.as_string())
    if cc != '':
        for c in cc.split(','):
            smtp.sendmail(sender, c, msg.as_string())
    smtp.close()


# python performance-statistic-v2.py 2017-10-17 xxx xxx '#core-1703a' core-1703a 'a@p.cn, b@p.cn' a@p.cn
def main(args):
    exec_date = args[1]
    user = args[2]
    pwd = args[3]
    query = args[4]
    sprint = args[5]

    today = datetime.now().strftime('%Y-%m-%d')
    if today == exec_date:
        subject = sprint + ' 本轮统计结果'
    else:
        subject = '每日统计  ' + sprint + '_' + today

    result_dic = get_performance_result_dic(user, pwd, query, sprint)
    export_csv(sprint, result_dic)
    if len(args) == 8:
        receivers = args[6]
        cc = args[7]
        send_mail(subject, receivers, cc, '执行时间: ' + date, path + sprint + '.csv')


if __name__ == '__main__':
    main(sys.argv)
