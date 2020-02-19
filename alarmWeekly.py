#encoding:utf-8
'''
@File    :   alarmWeekly.py
@Time    :   2019/12/13 10:31:52
@Author  :   wang
@Version :   1.0
@Contact :   wang@imprexion.com.cn
'''
# Start typing your code from here
from email.header import Header
from email.utils import parseaddr, formataddr
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import logging
import smtplib
import mysql.connector

dbUser = 'root'
dbPassword = 'xxx'
dbName = 'alarms'

envTest = ['office_test_Num10', 'K1', 'K2', 'K3']

from_addr = 'xxx@imprexion.com.cn'
from_password = 'xxx'
# 输入收件人地址:
toAddrs = 'xxx@imprexion.com.cn,\
    xxx@imprexion.com.cn,\
    xxx@imprexion.com.cn'

# 输入 SMTP 服务器地址:
# smtp_server = 'smtp.imprexion.com.cn'
smtp_server = 'smtp.mxhichina.com'

mysqlConn = None
mysqlCursor = None

mallNamePrefixs = ['SZ_YLZX_', 'GY_YXC_']
mallNames = {'SZ_YLZX_': '深圳印力中心', 'GY_YXC_': '贵阳印象城'}

tableStart = '''<table  cellpadding="0" style="border-collapse:collapse;width:888.0px;border-color:#666666;border-width:1.0px;border-style:solid;">
 <colgroup ><col  style="width:108.0px;">
 <col  span="3" style="width:87.0px;">
 </colgroup><tbody ><tr  height="40">
  <td  align="center" class="xl67" colspan="10" height="40" valign="middle" width="888" style="padding-top:1.0px;padding-right:1.0px;padding-left:1.0px;color:#000000;font-size:18.7px;font-weight:400;font-style:normal;text-decoration:none solid #000000;font-family:微软雅黑,sans-serif;border:1.0px solid #666666;background-image:none;background-position:.0% .0%;background-size:auto;background-attachment:scroll;background-origin:padding-box;background-clip:border-box;background-color:#bdd7ee;">每周告警统计</td>
 </tr >'''

rowStart = '''
 <tr  height="24">  
  '''
domainStart = '''<td  align="center" class="xl65" valign="middle" style="padding-top:1.0px;padding-right:1.0px;padding-left:1.0px;color:#000000;font-size:16.0px;font-weight:400;font-style:normal;text-decoration:none solid #000000;font-family:微软雅黑,sans-serif;border:1.0px solid #666666;background-image:none;background-position:.0% .0%;background-size:auto;background-attachment:scroll;background-origin:padding-box;background-clip:border-box;background-color: ;">'''
domainEnd = '</td>'
rowEnd = '</tr>'
tableEnd = '</tbody></table>'

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s: %(message)s',
    datefmt='%a, %d %b %Y %H:%M:%S')


def initMysql(loginUser, passwd, db):
    global mysqlConn, mysqlCursor
    mysqlConn = mysql.connector.connect(user=loginUser,
                                        password=passwd,
                                        database=db)
    mysqlCursor = mysqlConn.cursor()
    if mysqlConn == None or mysqlCursor == None:
        return False
    else:
        return True


def unInitMysql():
    global mysqlConn, mysqlCursor
    if mysqlCursor != None:
        mysqlCursor.close()
    if mysqlConn != None:
        mysqlConn.close()


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


# 如果是1号--9号则变成01--09
def addZeroPrefix(sDay):
    if sDay < 10:
        return '0' + str(sDay)
    return str(sDay)


def dateTimeWeekDouble(dtStamp):
    if not isinstance(dtStamp, datetime):
        return

    # 根据当前周的星期天来判断，否则可能不对
    weekIdx = dtStamp.weekday()
    dtStamp = dtStamp + timedelta(days=(6 - weekIdx))

    dtSecond = int(dtStamp.timestamp())
    dtDay = int(dtSecond / 86400)
    modDay = dtDay % 14
    if modDay < 7:
        return False
    else:
        return True


def getCurrentWeekWorkSpan():
    dtStart = datetime.now()
    weekIdx = dtStart.weekday()
    startDayDouble = dateTimeWeekDouble(dtStart)
    if None == startDayDouble:
        return ''

    weekIdx = dtStart.weekday()
    dtStart = dtStart - timedelta(days=weekIdx)
    dtEnd = dtStart
    if startDayDouble:
        # 双休
        dtEnd = dtStart + timedelta(days=4)
    else:
        dtEnd = dtStart + timedelta(days=5)

    return addZeroPrefix(dtStart.month) + '.' + addZeroPrefix(
        dtStart.day) + '-' + str(dtEnd.year) + '.' + addZeroPrefix(
            dtEnd.month) + '.' + addZeroPrefix(dtEnd.day)


def getEmailSubject():
    dtNow = datetime.now()
    dtStart = dtNow - timedelta(days=7)
    # _12.09-2019.12.14
    return '告警统计_' + addZeroPrefix(dtStart.month) + '.' + addZeroPrefix(
        dtStart.day) + '-' + str(dtNow.year) + '.' + addZeroPrefix(
            dtNow.month) + '.' + addZeroPrefix(dtNow.day)


def composeEmail():
    retContent = ''
    blankLine = '<div><br></div>'
    wrapInfoStart = '<div style="font-family: ' + 'Microsoft YaHei UI' + \
        ', Tahoma; line-height: normal; clear: both;">'
    wrapInfoEnd = '</div>'
    retContent = retContent + '<body>Hi，各位好:<br/>以下是本周各个商场告警统计情况反馈，请查阅，谢谢！！'
    # 准备添加表格，先添加一个空白行
    retContent = retContent + blankLine
    for mallPrefix in mallNamePrefixs:
        # 添加两个空白行
        retContent = retContent + blankLine + blankLine
        table1Content = getTableAlarms(mallPrefix)
        retContent = retContent + mallNames[mallPrefix]+'告警表格如下：'
        retContent = retContent + table1Content
        retContent = retContent + blankLine

        table2Content = getTableAlerts(mallPrefix)
        retContent = retContent + mallNames[mallPrefix]+'异常表格如下：'
        retContent = retContent + table2Content

    retContent = retContent + "</body>"
    # 准备添加hard和suggest，先添加一个空白行
    retContent = retContent + blankLine
    return retContent


def getAlarms(mallPrefix):
    global mysqlCursor
    mysqlCursor.execute(
        'select event_cases.endpoint,events.status,event_cases.note,event_cases.metric,events.timestamp \
            from events left join event_cases on events.event_caseId=event_cases.id \
                where endpoint like %s', [mallPrefix+'%']
    )
    values = mysqlCursor.fetchall()
    return values


def getAlerts(mallPrefix):
    dtNow = datetime.now()
    sinceDate = datetime.now() - timedelta(days=7)

    global mysqlCursor
    mysqlCursor.execute(
        'select endpoint,note,metric,timestamp from alerts where timestamp > %s and endpoint like %s',
        [sinceDate.strftime("%Y-%m-%d %H:%M:%S"), mallPrefix+'%'])
    values = mysqlCursor.fetchall()
    return values


def getPackageName(event_metric):
    if event_metric.find("packageName=") == -1:
        return ''
    else:
        return event_metric[event_metric.find("packageName=")+len("packageName="):]


def getAlertPackageName(alert_metric):
    if None == alert_metric:
        return ''
    if alert_metric.find("packageName=") == -1:
        return ''
    else:
        nameEnd = alert_metric.find(",")
        if nameEnd == -1:
            return alert_metric[alert_metric.find("packageName=")+len("packageName="):]
        else:
            return alert_metric[alert_metric.find("packageName=")+len("packageName="):nameEnd]


def getTableAlarms(mallPrefix):
    alarms = getAlarms(mallPrefix)
    alarmsTable = tableStart

    # 添加表格头
    tableHeader = rowStart
    tableHeader = tableHeader + '<th>' + '主机名' + '</th>'
    tableHeader = tableHeader + '<th>' + 'note' + '</th>'
    tableHeader = tableHeader + '<th>' + '告警发生时间' + '</th>'
    tableHeader = tableHeader + rowEnd
    alarmsTable = alarmsTable + tableHeader

    for alarm in alarms:
        if len(alarm) != 5:
            continue

        if alarm[0] in envTest:
            continue

        alarmRow = rowStart
        alarmRow = alarmRow + domainStart + alarm[0] + domainEnd
        if alarm[1] == 1:
            alarmRow = alarmRow + domainStart + \
                getPackageName(alarm[3]) + alarm[2] + '已恢复' + domainEnd
        else:
            alarmRow = alarmRow + domainStart + \
                getPackageName(alarm[3]) + alarm[2] + domainEnd
        alarmRow = alarmRow + domainStart + alarm[4].strftime(
            "%Y-%m-%d %H:%M") + domainEnd
        alarmRow = alarmRow + rowEnd

        alarmsTable = alarmsTable + alarmRow

    alarmsTable = alarmsTable + tableEnd
    return alarmsTable


def getTableAlerts(mallPrefix):
    alerts = getAlerts(mallPrefix)
    alertsTable = tableStart

   # 添加表格头
    tableHeader = rowStart
    tableHeader = tableHeader + '<th>' + '主机名' + '</th>'
    tableHeader = tableHeader + '<th>' + 'note' + '</th>'
    tableHeader = tableHeader + '<th>' + '异常发生时间' + '</th>'
    tableHeader = tableHeader + rowEnd
    alertsTable = alertsTable + tableHeader

    for alert in alerts:
        if len(alert) != 4:
            continue

        if alert[0] in envTest:
            continue

        alarmRow = rowStart
        alarmRow = alarmRow + domainStart + alert[0] + domainEnd
        alarmRow = alarmRow + domainStart + \
            getAlertPackageName(alert[2]) + alert[1] + domainEnd
        alarmRow = alarmRow + domainStart + alert[3].strftime(
            "%Y-%m-%d %H:%M") + domainEnd
        alarmRow = alarmRow + rowEnd

        alertsTable = alertsTable + alarmRow

    alertsTable = alertsTable + tableEnd
    return alertsTable


def procAlarmWeekly():
    if False == initMysql(dbUser, dbPassword, dbName):
        logging.error(u'连接数据库失败:')
        return False

    sendAlarmWeekly(composeEmail())
    unInitMysql()


def sendAlarmWeekly(bodyContent):
    if '' == bodyContent.strip():
        return

    msg = MIMEText(bodyContent, 'html', 'utf-8')
    msg['From'] = _format_addr('告警统计机器人 <%s>' % from_addr) # 显示的发件人
    # msg['To'] = _format_addr('研发部 <%s>' % to_addr)          # 单个显示的收件人
    msg['To'] = toAddrs                                          # 多个显示的收件人
    msg['Subject'] = Header(getEmailSubject(), 'utf-8').encode() # 显示的邮件标题
    server = smtplib.SMTP_SSL(smtp_server, 465)
    # server.set_debuglevel(1)
    server.login(from_addr, from_password)
    server.sendmail(from_addr, toAddrs.split(','), msg.as_string())
    server.quit()


if __name__ == '__main__':
    procAlarmWeekly()