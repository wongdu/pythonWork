#encoding:utf-8
'''
@File    :   weeklyStatistic.py
@Time    :   2019/11/30 15:33:15
@Author  :   wang
@Version :   1.0
@Contact :   wang@imprexion.com.cn
'''
# Start typing your code from here

from imapclient import IMAPClient
import email
from email.header import decode_header, Header
from email.utils import parseaddr, formataddr
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import logging
import sys
import time
import html
import re
from datetime import datetime, timedelta
import smtplib

hostname = 'imap.imprexion.com.cn'
user = 'xxx@imprexion.com.cn'
password = 'xxx'

# 输入收件人地址:
to_addr = 'xxx@imprexion.com.cn'
# 输入 SMTP 服务器地址:
smtp_server = 'smtp.mxhichina.com'

AllStaffs = {
    'chen.quanlin': '陈全林',
    'li.liang': '李亮',
    'zhao.tao': '赵涛',    
    'du.long': '杜龙',
}

logging.basicConfig(
    level=logging.INFO,
    format=
    '%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s: %(message)s',
    datefmt='%a, %d %b %Y %H:%M:%S')


def weekDouble():
    dtStamp = datetime.now()

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


# 保留两位数字
def get_two_float(f_str, n):
    f_str = str(f_str)  # f_str = '{}'.format(f_str) 也可以转换为字符串
    a, b, c = f_str.partition('.')
    c = (c + "0" * n)[:n]  # 如论传入的函数有几位小数，在字符串后面都添加n为小数0
    return ".".join([a, c])


# 用户名+密码登陆
def loginEmailbox():
    c = IMAPClient(hostname, use_uid=True, ssl=False)
    try:
        c.login(user, password)
        logging.info(u'登录成功')
    except c.Error:
        logging.error(u'用户名或密码错误')
        sys.exit(1)
    return c


# 退出登陆
def logoutEmailbox(c):
    c.logout()


# 获取所有的周报邮件
def fetchAllWeeklyEmails(c, folder='INBOX'):
    c.select_folder(folder, readonly=True)
    # 根据大小周从上班前一天开始检索
    sinceDate = None
    if weekDouble():
        sinceDate = datetime.now() - timedelta(days=3)
    else:
        sinceDate = datetime.now() - timedelta(days=2)

    result = c.search([u'SINCE', sinceDate])
    return c.fetch(result, ['BODY.PEEK[]'])


# 获取上周所发的周报统计邮件
def fetchLastWeeklyStatisticEmail(c, folder='INBOX'):
    c.select_folder(folder, readonly=True)
    # 从上周的当前时间开始检索
    result = c.search([u'SINCE', datetime.now() - timedelta(days=7)])
    return c.fetch(result, ['BODY.PEEK[]'])


def listFolders(c):
    return c.list_folders()


def decodeString(strParam):
    value, charset = decode_header(strParam)[0]
    if charset:
        value = value.decode(charset)
    return value


def subAfterTable(strWeeklyTable):
    if '' == strWeeklyTable.strip():
        return ''
    suggestIdx = -1
    if strWeeklyTable.count('意见和建议') > 1:
        suggestIdx = strWeeklyTable.rfind('意见和建议')
    elif strWeeklyTable.count('意见和建议') == 1:
        suggestIdx = strWeeklyTable.find('意见和建议')

    if suggestIdx == -1:
        return strWeeklyTable

    idxEndSuggestTitle = strWeeklyTable.find('</tr>', suggestIdx)
    if idxEndSuggestTitle == -1:
        return strWeeklyTable
    idxEndSuggestContent = strWeeklyTable.find(
        '</tr>', idxEndSuggestTitle + len('</tr>'))
    idxEndBody = strWeeklyTable.find('</tbody>',
                                     idxEndSuggestTitle + len('</tr>'))
    if idxEndBody == -1:
        return strWeeklyTable

    if idxEndSuggestContent != -1 and idxEndSuggestContent > idxEndSuggestTitle and idxEndSuggestContent < idxEndBody:
        return strWeeklyTable[:idxEndSuggestContent +
                              len('</tr>')] + strWeeklyTable[idxEndBody:]
    elif idxEndSuggestContent == -1:
        return strWeeklyTable
    else:
        return strWeeklyTable


def getWeeklyContent(strWeeklyTable):
    idxLastSend = strWeeklyTable.find('发件人')
    if idxLastSend != -1:
        strWeeklyTable = strWeeklyTable[:idxLastSend]
    strWeeklyTable = subAfterTable(strWeeklyTable)
    # 消除\n，避免不能使用正则表达式
    strWeeklyTable = strWeeklyTable.replace('\n', ' ')
    if '' == strWeeklyTable.strip():
        return []
    # 1、去掉&nbsp;
    strWeeklyTable = html.unescape(strWeeklyTable)
    # 1、将&nbsp转换成的\xa0用空格代替;
    strWeeklyTable = strWeeklyTable.replace(u'\u3000', u' ').replace(
        u'\xa0', ' ').replace(u'\t', u' ')
    if '' == strWeeklyTable.strip():
        return []

    lContent = []
    while len(strWeeklyTable) > 0:
        matchObj = re.match(r'(.*?)>(.*?)<', strWeeklyTable, re.M | re.I)
        if matchObj:
            lContent.append(matchObj.group(2))
            strWeeklyTable = strWeeklyTable[matchObj.span()[1]:]
        else:
            break

    return lContent


# 将list中"风险和困难 "这样的字符串修改为"风险和困难"
def subUnnecessaryChars(lContent, strDetail):
    if type(lContent).__name__ != 'list':
        return []
    if type(strDetail).__name__ != 'str':
        return lContent

    if '' == strDetail:
        return lContent

    for idx in range(len(lContent)):
        if lContent[idx].find(strDetail) != -1:
            lContent[idx] = strDetail

    return lContent


# 获取'风险和困难'、'意见和建议'的值
def getHardAndSuggest(lRealContent):
    subUnnecessaryChars(lRealContent, "风险和困难")
    subUnnecessaryChars(lRealContent, "意见和建议")
    hardIdx = None
    suggestIdx = None
    if lRealContent.count('风险和困难') > 1:
        lTemp = lRealContent[:]
        lTemp.reverse()
        hardIdx = lTemp.index("风险和困难")
        hardIdx = -hardIdx - 1
        lTemp = None
    elif lRealContent.count('风险和困难') == 1:
        hardIdx = lRealContent.index('风险和困难')

    if lRealContent.count('意见和建议') > 1:
        lTemp = lRealContent.copy()
        lTemp.reverse()
        suggestIdx = lTemp.index("意见和建议")
        suggestIdx = -suggestIdx - 1
        lTemp = None
    elif lRealContent.count('意见和建议') == 1:
        suggestIdx = lRealContent.index('意见和建议')

    if hardIdx == None:
        return "", ""

    if suggestIdx == None:
        return lRealContent[int(hardIdx) + 1:], ""

    hardContent = lRealContent[int(hardIdx) + 1:int(suggestIdx)]
    suggestContent = lRealContent[int(suggestIdx) + 1:]
    return hardContent, suggestContent


def guessCharset(msg):
    # 先从msg对象获取编码:
    charset = msg.get_charset()
    if charset is None:
        # 如果获取不到，再从Content-Type字段获取:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


def parsePart(emaileMessage):
    contents = []
    if emaileMessage.is_multipart():
        # 如果邮件对象是一个MIMEMultipart, get_payload()返回list，包含所有的子对象:
        parts = emaileMessage.get_payload()
        for n, part in enumerate(parts):
            # 递归获取每一个子对象:
            contents += (parsePart(part))
    else:
        # 邮件对象不是一个MIMEMultipart, 就根据content_type判断:
        content_type = emaileMessage.get_content_type()
        if content_type == 'text/html':
            # 纯HTML内容:
            content = emaileMessage.get_payload(decode=True).strip()
            # 要检测文本编码:
            charset = guessCharset(emaileMessage)
            if charset:
                contents.append(content.decode(charset))
            else:
                contents.append(content.decode('gbk'))
        elif content_type == 'text/plain':
            # 纯文本内容
            content = emaileMessage.get_payload(decode=True).strip()
            # 要检测文本编码
            charset = guessCharset(emaileMessage)
            if charset:
                contents.append(content.decode(charset))
            else:
                contents.append(content.decode('gbk'))
        else:
            # to do 不是文本,作为附件处理，要获取文件名称
            pass

    return contents


#解析单个邮件，返回：主题、发件人、收件人、困难、建议
def parseMsgDict(msgDict):
    emaileMessage = email.message_from_string(
        msgDict[b'BODY[]'].decode())  # 生成Message类型

    subject = decodeString(emaileMessage['Subject'])

    mailFrom = email.header.make_header(
        email.header.decode_header(emaileMessage['From']))
    mailTo = email.header.make_header(
        email.header.decode_header(emaileMessage['To']))

    strMailFrom = str(mailFrom)
    idxLeftFromName = strMailFrom.rfind("<")
    idxRightFromName = strMailFrom.rfind("@")
    # 不能正常获取发件人名字就直接返回
    if idxLeftFromName == -1 or idxRightFromName == -1 or idxLeftFromName >= idxRightFromName:
        return subject, mailFrom, mailTo, None, None
    else:
        mailFrom = strMailFrom[idxLeftFromName + 1:idxRightFromName]

    strMailTo = str(mailTo)
    idxLeftToName = strMailTo.rfind("<")
    idxRightToName = strMailTo.rfind("@")
    # 不能正常获取收件人名字就直接返回
    if idxLeftToName == -1 or idxRightToName == -1 or idxLeftToName >= idxRightToName:
        return subject, mailFrom, mailTo, None, None
    else:
        mailTo = strMailTo[idxLeftToName + 1:idxRightToName]

    # 如果邮件不是发送到研发部就直接返回
    if mailTo != 'rd':
        return subject, mailFrom, mailTo, None, None
    # 如果subject不包含周报这两个字符就直接返回
    if type(subject).__name__ != 'str' or subject.find("周报") == -1:
        return subject, mailFrom, mailTo, None, None
    # 跳过周报统计
    if type(subject).__name__ != 'str' or subject.find("周报统计") != -1:
        return subject, mailFrom, mailTo, None, None

    mailContents = parsePart(emaileMessage)
    tableContent = ''
    if len(mailContents) == 1:
        if type(mailContents[0]).__name__ != 'str':
            return subject, mailFrom, mailTo, '', ''
        tableContent = mailContents[0]
    else:
        tableContent = mailContents[1]

    leftWeeklyTable = tableContent.find('<table')
    if leftWeeklyTable == -1:
        return subject, mailFrom, mailTo, '', ''
    rightWeeklyTable = tableContent.find('</table>',
                                         leftWeeklyTable + len('<table'))
    if rightWeeklyTable == -1:
        return subject, mailFrom, mailTo, '', ''

    if leftWeeklyTable > rightWeeklyTable:
        return subject, mailFrom, mailTo, '', ''

    lTemp = getWeeklyContent(tableContent[leftWeeklyTable:rightWeeklyTable])
    lRealContent = []
    for textValue in lTemp:
        if len(textValue) > 0 and '' != textValue.strip():
            lRealContent.append(textValue)

    hard, suggest = getHardAndSuggest(lRealContent)

    mailContent = '\n'.join(mailContents)
    return subject, mailFrom, mailTo, hard, suggest


def parseForLastWeeklyStatisticTable(msgDict):
    emaileMessage = email.message_from_string(
        msgDict[b'BODY[]'].decode())  # 生成Message类型

    subject = decodeString(emaileMessage['Subject'])
    mailFrom = email.header.make_header(
        email.header.decode_header(emaileMessage['From']))
    mailTo = email.header.make_header(
        email.header.decode_header(emaileMessage['To']))

    strMailFrom = str(mailFrom)
    idxLeftFromName = strMailFrom.rfind("<")
    idxRightFromName = strMailFrom.rfind("@")
    # 不能正常获取发件人名字就直接返回
    if idxLeftFromName == -1 or idxRightFromName == -1 or idxLeftFromName >= idxRightFromName:
        return ''
    else:
        mailFrom = strMailFrom[idxLeftFromName + 1:idxRightFromName]

    strMailTo = str(mailTo)
    idxLeftToName = strMailTo.rfind("<")
    idxRightToName = strMailTo.rfind("@")
    # 不能正常获取收件人名字就直接返回
    if idxLeftToName == -1 or idxRightToName == -1 or idxLeftToName >= idxRightToName:
        return ''
    else:
        mailTo = strMailTo[idxLeftToName + 1:idxRightToName]

    if mailFrom != 'wang.xxx':
        return ''
    # 如果邮件不是发送到研发部就肯定不是周报统计，直接跳过
    # if mailTo != 'rd':
    #     return ''

    # 不是周报统计就跳过
    if type(subject).__name__ != 'str' or subject.find("周报统计") == -1:
        return ''

    mailContents = parsePart(emaileMessage)
    if len(mailContents) == 2:
        tableContent = mailContents[1]
    elif len(mailContents) == 1:
        tableContent = mailContents[0]
    else:
        return ''

    # tableContent = mailContents[1]
    leftWeeklyStatisticTable = tableContent.find('<table')
    if leftWeeklyStatisticTable == -1:
        return ''
    rightWeeklyStatisticTable = tableContent.find(
        '</table>', leftWeeklyStatisticTable + len('<table'))
    if rightWeeklyStatisticTable == -1:
        return ''
    if leftWeeklyStatisticTable > rightWeeklyStatisticTable:
        return ''

    return tableContent[leftWeeklyStatisticTable:rightWeeklyStatisticTable +
                        len('</table>')]


def getLastWeeklyStatisticTable(imapClient):
    msgDicts = fetchLastWeeklyStatisticEmail(imapClient)
    for msgId, msgDict in msgDicts.items():
        weeklyStatisticTable = parseForLastWeeklyStatisticTable(msgDict)
        if '' != weeklyStatisticTable.strip():
            # 找到就退出
            return checkFirstWeeklyInNewMonth(weeklyStatisticTable)

    return ''


# 获取三个colspan的值
def getColspan(weeklyStatisticTable, startIdx):
    leftQutoIdx = weeklyStatisticTable.find('"', startIdx)
    rightQutoIdx = weeklyStatisticTable.find('"', leftQutoIdx + len('"'))
    if leftQutoIdx == -1 or rightQutoIdx == -1 or leftQutoIdx >= rightQutoIdx:
        return -1
    return int(weeklyStatisticTable[leftQutoIdx + len('"'):rightQutoIdx])


# 如果是1号--9号则变成01--09
def addZeroPrefix(sDay):
    if sDay < 10:
        return '0' + str(sDay)
    return str(sDay)


# 获取当前周最后一个工作日是本月的第几个，即索引
# 返回0表示是上个月的最后一个工作日
def getCurrMonthLastworkIndex():
    t = datetime.now()
    currMon = t.month
    doubleRest = weekDouble()
    idx = 0
    while True:
        dtToJudge = datetime.now() - timedelta(days=(7 * idx))
        if doubleRest:
            dtToJudge = dtToJudge - timedelta(days=2)
        else:
            dtToJudge = dtToJudge - timedelta(days=1)
        if dtToJudge.month == currMon:
            idx = idx + 1
            doubleRest = bool(1 - doubleRest)
        else:
            break

    return idx


# 获取当前周最后一个工作日后面(不计算当前周)，还有几个本月的最后工作日
# 思路：从当前位置开始统计总共有几个最后工作日，
# 即当前月的总共的最后工作日个数-获取的该值=当前最后工作日的索引
def getLastworkAfterCurrWeek():
    dtStart = datetime.now()
    weekIdx = dtStart.weekday()
    startDayDouble = dateTimeWeekDouble(dtStart)
    if None == startDayDouble:
        return 0

    currMon = dtStart.month
    # 不管如何，在下面开始while循环开始计数前，跳到下周的最后一个工作日
    if startDayDouble:
        # 当前为大周双休
        if weekIdx == 4:
            # 今天周五恰好是双休最后一个工作日，
            # +8天跳到下周六(+2跳到本周周天然后+6跳到下周小周的周六)
            dtStart = dtStart + timedelta(days=8)
        elif weekIdx > 4:
            # 今天是周末，可能周六或周天，先到当前周最后一天：6 - weekIdx，
            # 然后下周小周上6天班所以+6
            dtStart = dtStart + timedelta(days=(6 - weekIdx + 6))
        else:
            # 下面这行只是跳到双休的周五，应该+2跳到本周周天然后+6跳到下周小周的周六
            # dtStart = dtStart + timedelta(days=(4 - weekIdx)) 
            dtStart = dtStart + timedelta(days=(6 - weekIdx + 6)) # 其实同上面周末情形一样
    else:
        # 当前为小周单休
        if weekIdx == 5:
            # 今天周六恰好是单休最后一个工作日，
            # +6天跳到下周五(+1跳到本周周日然后+5跳到下周大周的周五)
            dtStart = dtStart + timedelta(days=6)
        elif weekIdx == 6:
            # 今天是周日，已经到了当前周最后一天，然后下周大周上5天班
            dtStart = dtStart + timedelta(days=5)
        else:
            # 今天不是周末，周一到周五，先到当前周最后一天：6 - weekIdx，
            # 然后下周大周上5天班所以+5
            dtStart = dtStart + timedelta(days=(6 - weekIdx + 5))

    idx = 0
    while True:
        if dtStart.month != currMon:
            break
        restDouble = dateTimeWeekDouble(dtStart)
        if None == restDouble:
            return 0
        if restDouble and dtStart.weekday() == 4:
            dtStart = dtStart + timedelta(days=8)
            idx = idx + 1
        elif restDouble == False and dtStart.weekday() == 5:
            dtStart = dtStart + timedelta(days=6)
            idx = idx + 1

    return idx


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


# 获取当前月每周最后一个工作日的数量
# 注意区别：获取当前月中当前周后面(不计算当前周)，还有几个一周最后工作日
def getCurrMonLastWorkDay():
    dtStart = datetime.now().replace(day=1)
    weekIdx = dtStart.weekday()
    startDayDouble = dateTimeWeekDouble(dtStart)
    if None == startDayDouble:
        return []

    currMon = dtStart.month
    lastWorkDay = []
    # 如果1号是最后一个工作日就直接加到返回的列表中，
    # 如果不是就跳到当前周的最后一个工作日，
    # 区别求当前周后面(不计入内)的最后工作日个数，因为那个函数是跳过当前周的最后工作日
    if startDayDouble:
        # 当前为大周双休
        if weekIdx == 4:
            # 1号恰好是最后一个工作日
            lastWorkDay.append(
                addZeroPrefix(currMon) + addZeroPrefix(dtStart.day))
            dtStart = dtStart + timedelta(days=8)
        elif weekIdx > 4:
            # 1号是周末，可能周六或周天，先到当前周最后一天，然后下周小周上6天班
            dtStart = dtStart + timedelta(days=(6 - weekIdx + 6))
        else:
            # 跳到当前双休周的周五，下面while循环自然会判断是否添加到列表中
            dtStart = dtStart + timedelta(days=(4 - weekIdx))
    else:
        # 当前为小周单休
        if weekIdx == 5:
            # 1号恰好是最后一个工作日
            lastWorkDay.append(
                addZeroPrefix(currMon) + addZeroPrefix(dtStart.day))
            dtStart = dtStart + timedelta(days=6)
        elif weekIdx == 6:
            # 1号是星期天，先到当前周最后一天，然后下周大周上5天班
            dtStart = dtStart + timedelta(days=5)
        else:
            # 跳到当前单休周的周六，下面while循环自然会判断是否添加到列表中
            dtStart = dtStart + timedelta(days=(5 - weekIdx))

    while True:
        if dtStart.month != currMon:
            break
        restDouble = dateTimeWeekDouble(dtStart)
        if None == restDouble:
            return []
        if restDouble and dtStart.weekday() == 4:
            lastWorkDay.append(
                addZeroPrefix(currMon) + addZeroPrefix(dtStart.day))
            dtStart = dtStart + timedelta(days=8)
        elif restDouble == False and dtStart.weekday() == 5:
            lastWorkDay.append(
                addZeroPrefix(currMon) + addZeroPrefix(dtStart.day))
            dtStart = dtStart + timedelta(days=6)

    return lastWorkDay


def updateQutoContent(weeklyTime, startIdx, newContent):
    leftQutoIdx = weeklyTime.find('"', startIdx)
    rightQutoIdx = weeklyTime.find('"', leftQutoIdx + len('"'))
    return weeklyTime[:leftQutoIdx +
                      len('"')] + str(newContent) + weeklyTime[rightQutoIdx:]


# 更新表格中第二tr部分的月份
def procWeeklyMonth(weeklyMonth, sPreLastMon, sLastMon, lastColSpan):
    if '' == weeklyMonth.strip():
        return ''

    currLenMonLastWorkDay = len(getCurrMonLastWorkDay())
    if 0 == currLenMonLastWorkDay:
        return ''

    # 先把月份更新了，因为都是'xx月份'这样的字符串，并不会改变索引位置
    currMon = datetime.now().month
    # currMon = addZeroPrefix(currMon)
    weeklyMonth = weeklyMonth.replace(sPreLastMon + u'月份',
                                      sLastMon + u'月份').replace(
                                          sLastMon + u'月份', currMon + u'月份')

    idxPreLastColspan = weeklyMonth.find('colspan=')
    idxLastColspan = weeklyMonth.find('colspan=',
                                      idxPreLastColspan + len('colspan='))
    if idxPreLastColspan == -1 or idxLastColspan == -1:
        return ''

    # 同上面的注释，注意顺序不能反了，否则索引位置可能发生了变化
    weeklyMonth = updateQutoContent(weeklyMonth, idxLastColspan,
                                    currLenMonLastWorkDay)
    weeklyMonth = updateQutoContent(weeklyMonth, idxPreLastColspan,
                                    lastColSpan)

    return weeklyMonth


def getLastMonthTdIndex(weeklyDay, preLastcols):
    retIdx = 0
    while preLastcols > 0:
        currIdx = weeklyDay.find('<td', retIdx)
        retIdx = currIdx
        preLastcols = preLastcols - 1

    return weeklyDay.find('<td', retIdx)


def getLastMonthWeekly(weeklyMonth):
    lContent = []
    while len(weeklyMonth) > 0:
        matchObj = re.match(r'(.*?)>(.*?)<', weeklyMonth, re.M | re.I)
        if matchObj:
            lContent.append(matchObj.group(2))
            weeklyMonth = weeklyMonth[matchObj.span()[1]:]
        else:
            break

    return lContent


# 更新表格第三行：每周最后一个工作日信息，即最后一个工作日日期
def procWeeklyDay(weeklyDay, preLastColSpan, lastColSpan):
    if '' == weeklyDay.strip():
        return ''

    tdNum = weeklyDay.count('td')
    if tdNum != (preLastColSpan + lastColSpan) * 2:
        return ''

    lastWeeklyDay = getLastMonthWeekly(
        weeklyDay[getLastMonthTdIndex(weeklyDay, preLastColSpan):])
    currWeeklyDay = getCurrMonLastWorkDay()

    idxFirstStartTd = weeklyDay.find('<td')
    idxFirstStartTdComplete = weeklyDay.find('>', idxFirstStartTd + len('<td'))

    firstEndTd = weeklyDay.find('</td>')
    halfDecorateEnd = weeklyDay.find('>', firstEndTd + len('</td>'))
    #每一项的前半部分属性、后半部分属性
    firstHalfDecorate = weeklyDay[firstEndTd + len('</td>'):halfDecorateEnd +
                                  len('>')]
    secondHalfDecorate = '</td>'

    lastEndTd = weeklyDay.rfind('</td>')
    # 最后一个td结束后的内容
    lastAfter = weeklyDay[lastEndTd + len('</td>'):]
    retWeeklyDay = weeklyDay[idxFirstStartTdComplete +
                             len('>'):] + lastWeeklyDay[0] + secondHalfDecorate
    for v in lastWeeklyDay[1:]:
        retWeeklyDay = retWeeklyDay + firstHalfDecorate + v + secondHalfDecorate
    for v in currWeeklyDay:
        retWeeklyDay = retWeeklyDay + firstHalfDecorate + v + secondHalfDecorate

    retWeeklyDay = retWeeklyDay + lastAfter
    return retWeeklyDay


def updateAllColspan(weeklyStatisticTable, idxFirstColspan):
    idxThirdColspan = weeklyStatisticTable.rfind('colspan=')
    if idxThirdColspan == -1:
        return ''

    lastColSpan = getColspan(weeklyStatisticTable, idxThirdColspan)
    if lastColSpan < 1:
        return ''

    weeklyStatisticTable = updateQutoContent(
        weeklyStatisticTable, idxFirstColspan,
        lastColSpan + len(getCurrMonLastWorkDay()) + 1)


# 当一个新的月的第一个每周最后工作日出现，需要更新统计表信息
def updateMonthTable(weeklyStatisticTable):
    # 不处理空白字符串
    if '' == weeklyStatisticTable.strip():
        return ''
    # 常规判断
    if weeklyStatisticTable.count('table') != 2 or weeklyStatisticTable.count(
            'table') != 2:
        return ''
    if weeklyStatisticTable.count('colspan=') != 3:
        return ''

    idxFirstColspan = weeklyStatisticTable.find('colspan=')
    if idxFirstColspan == -1:
        # 0、更新表格中第一tr部分总的colspan
        weeklyStatisticTable = updateAllColspan(weeklyStatisticTable,
                                                idxFirstColspan)

    # 1、获取上个月和前一个月colspan的value
    idxThirdColspan = weeklyStatisticTable.rfind('colspan=')
    idxSecondColspan = weeklyStatisticTable.find(
        'colspan=', idxFirstColspan + len('colspan='))
    lastColSpan = getColspan(weeklyStatisticTable, idxThirdColspan)
    preLastColSpan = getColspan(weeklyStatisticTable, idxSecondColspan)
    if lastColSpan < 1 or preLastColSpan < 1:
        return ''

    # 2、获取某个月份的每周的最后一个工作日
    firstEndTr = weeklyStatisticTable.find('</tr>')
    secondStartTr = weeklyStatisticTable.find('<tr', firstEndTr + len('</tr>'))
    secondEndTr = weeklyStatisticTable.find('</tr>', firstEndTr + len('</tr>'))
    if firstEndTr == -1 or secondStartTr == -1 or secondEndTr == -1:
        return ''

    dtNow = datetime.now()
    # 3、a: 如果减去30可能就跨过一个月，执行到这里肯定是当前月的1号--9号，
    # 减去15天肯定是到了上个月，而目的也是获取上个月的月份，所以没有问题
    #    b: 同理如果减去45天的话就肯定到上个月的前一个月
    dtLastHalfMonth = datetime.now() - timedelta(days=15)
    dtPreLastHalfMonth = datetime.now() - timedelta(days=45)
    sCurrentMonth = str(dtNow.month)
    sLastMonth = str(dtLastHalfMonth.month)
    sPreLastMonth = str(dtPreLastHalfMonth.month)

    # 4、更新表格第二行：月份信息
    newWeeklyMonth = procWeeklyMonth(
        weeklyStatisticTable[secondStartTr:secondEndTr], sPreLastMonth,
        sLastMonth, lastColSpan)
    if '' == newWeeklyMonth:
        return ''

    weeklyStatisticTable = weeklyStatisticTable[:
                                                secondStartTr] + newWeeklyMonth + weeklyStatisticTable[
                                                    secondEndTr:]

    thirdStartTr = weeklyStatisticTable.find('<tr', secondEndTr + len('</tr>'))
    thirdEndTr = weeklyStatisticTable.find('</tr>', secondEndTr + len('</tr>'))
    if thirdStartTr == -1 or thirdEndTr == -1:
        return ''

    # 5、更新表格第三行：每周最后一个工作日信息
    newWeeklyDay = procWeeklyDay(weeklyStatisticTable[thirdStartTr:thirdEndTr],
                                 preLastColSpan, lastColSpan)
    if '' == newWeeklyMonth:
        return ''

    weeklyStatisticTable = weeklyStatisticTable[:
                                                thirdStartTr] + newWeeklyDay + weeklyStatisticTable[
                                                    thirdEndTr:]

    # 6、  更新每个用户的统计信息，把上个月的前一个月的信息删掉
    # 6.1、更新发送率统计信息，把上个月的前一个月的信息删掉
    weeklyStatisticTable = updateAllUsersStatistic(weeklyStatisticTable,
                                                   preLastColSpan)

    return weeklyStatisticTable


def updateOneRow(oneTableRow, preLastcols):
    if '' == oneTableRow.strip():
        return ''

    lastWeeklyDay = getLastMonthWeekly(
        oneTableRow[getLastMonthTdIndex(oneTableRow, preLastcols):])
    currWeeklyDay = getCurrMonLastWorkDay()

    idxFirstStartTd = oneTableRow.find('<td')
    idxFirstStartTdComplete = oneTableRow.find('>',
                                               idxFirstStartTd + len('<td'))

    firstEndTd = oneTableRow.find('</td>')
    halfDecorateEnd = oneTableRow.find('>', firstEndTd + len('</td>'))
    #每一项的前半部分属性、后半部分属性
    firstHalfDecorate = oneTableRow[firstEndTd + len('</td>'):halfDecorateEnd +
                                    len('>')]
    secondHalfDecorate = '</td>'

    lastEndTd = oneTableRow.rfind('</td>')
    # 最后一个td结束后的内容
    lastAfter = oneTableRow[lastEndTd + len('</td>'):]
    retWeeklyDay = oneTableRow[idxFirstStartTdComplete +
                               len('>'
                                   ):] + lastWeeklyDay[0] + secondHalfDecorate
    for v in lastWeeklyDay[1:]:
        retWeeklyDay = retWeeklyDay + firstHalfDecorate + v + secondHalfDecorate
    for v in currWeeklyDay:
        retWeeklyDay = retWeeklyDay + firstHalfDecorate + ' ' + secondHalfDecorate

    retWeeklyDay = retWeeklyDay + lastAfter
    return retWeeklyDay


# 当新的月份的第一个最后工作日出现，要把上个月的前个月信息删除，包括发送统计率
def updateAllUsersStatistic(weeklyUsersTable, preLastcols):
    if '' == weeklyUsersTable.strip():
        return ''

    firstTr = weeklyUsersTable.find('<tr')
    idx2ndTr = weeklyUsersTable.find('<tr', firstTr + len('<tr'))
    idx3rdTr = weeklyUsersTable.find('<tr', idx2ndTr + len('<tr'))
    idxBase = idx3rdTr + len('<tr')
    while True:
        idxStart = weeklyUsersTable.find('<tr', idxBase)
        idxEnd = weeklyUsersTable.find('</tr>', idxStart)
        if idxStart == -1:
            break
        newOneRow = updateOneRow(
            weeklyUsersTable[idxStart:idxEnd + len('</tr>')], preLastcols)
        temp = weeklyUsersTable[:idxStart] + newOneRow + weeklyUsersTable[
            idxEnd + len('</tr>'):]
        weeklyUsersTable = temp
        idxBase = idxStart + len('<tr')

    return weeklyUsersTable


# 判断是否是将上个月的前一个月的列信息删掉，并添加上当前月的列
def checkFirstWeeklyInNewMonth(weeklyStatisticTable):
    weeklyStatisticTable = weeklyStatisticTable.strip()
    dtNow = datetime.now()
    currDay = dtNow.day
    if currDay > 9:
        return weeklyStatisticTable

    if weekDouble() or currDay < 3 or currDay == 9:
        return weeklyStatisticTable
    elif currDay < 2:
        return weeklyStatisticTable

    # 先将上个月的前一个月的信息删除掉
    # 再将当前月的每周最后一个工作日信息加上
    return updateMonthTable(weeklyStatisticTable)


def updateUserSentFlag(userSend, updateValue):
    currIdx = getLastworkAfterCurrWeek()
    currMonSize = len(getCurrMonLastWorkDay())
    # +1是因为修改的是当前周的统计信息，而减掉的是包含当前周及以后的数量
    # spanSize = currMonSize - currIdx

    # 现在获取的已经就是当前周后面的最后工作日数量，
    # 所以直接从后往前跨过去就好，无需用当前月总的最后工作日数量-当前周后面的最后工作日数量
    spanSize=currIdx
    idxBase = userSend.rfind('</td>')
    while spanSize > 0:
        idxTemp = userSend.rfind('</td>', 0, idxBase - len('</td>'))
        idxBase = idxTemp
        spanSize = spanSize - 1

    idxUpdate = userSend.rfind('">', 0, idxBase)
    retValue = userSend[:idxUpdate +
                        len('">')] + updateValue + userSend[idxBase:]
    return retValue


def updateUserNoSentFlag(userSend):
    currIdx = getLastworkAfterCurrWeek()
    currMonSize = len(getCurrMonLastWorkDay())
    # +1是因为修改的是当前周的统计信息，而减掉的是包含当前周及以后的数量
    # spanSize = currMonSize - currIdx

    # 现在获取的已经就是当前周后面的最后工作日数量，
    # 所以直接从后往前跨过去就好，无需用当前月总的最后工作日数量-当前周后面的最后工作日数量
    spanSize=currIdx
    idxBase = userSend.rfind('</td>')
    while spanSize > 0:
        idxTemp = userSend.rfind('</td>', 0, idxBase - len('</td>'))
        idxBase = idxTemp
        spanSize = spanSize - 1

    idxUpdate = userSend.rfind('">', 0, idxBase)
    retValue = userSend[:
                        idxUpdate] + ' align="center"' + '">' + '<img src="cid:image1" alt="❌">' + userSend[
                            idxBase:]
    return retValue


def updateUserSentRow(toSendTable, userName, updateValue):
    if '' == userName.strip():
        return toSendTable
    idxBase = toSendTable.find(userName)
    idxUpdate = toSendTable.find('</tr>', idxBase)
    # 找不到用户名，可能是新用户，暂不处理，后面添加了新用户的行信息后再处理
    if idxBase == -1 or idxUpdate == -1:
        return toSendTable
    newUserSentFlag = updateUserSentFlag(toSendTable[idxBase:idxUpdate],
                                         updateValue)
    retValue = toSendTable[:idxBase] + newUserSentFlag + toSendTable[idxUpdate:]
    return retValue


def updateUserNoSentRow(toSendTable, userName):
    if '' == userName.strip():
        return toSendTable
    idxBase = toSendTable.find(userName)
    idxUpdate = toSendTable.find('</tr>', idxBase)
    # 因为先添加新用户，然后再更新未发送标记，不过为了跟更新发送标记一致，所以也判断处理
    # 也有一种情况是：删除了该用户，还想更新未发送标记，那就直接返回
    if idxBase == -1 or idxUpdate == -1:
        return toSendTable
    newUserSentFlag = updateUserNoSentFlag(toSendTable[idxBase:idxUpdate])
    retValue = toSendTable[:idxBase] + newUserSentFlag + toSendTable[idxUpdate:]
    return retValue


# 把新加的用户的所有的统计信息先清空
def procNewUserStatistic(userContent, userName):
    if '' == userName.strip():
        return userContent
    if '' == userContent.strip():
        return ''

    firstTd = userContent.find('<td')
    if -1 == firstTd:
        return ''

    matchName = re.match(r'(.*?);">(.*?)</td>', userContent[firstTd:],
                         re.M | re.I)
    if matchName == None:
        return ''

    nameTup = matchName.span(2)
    # 先把名字修改了
    tempValue = userContent[:nameTup[0] + firstTd] + userName
    rightPart = userContent[nameTup[1] + firstTd:]

    while len(rightPart) > 0:
        matchObj = re.match(r'(.*?)">(.*?)</td>', rightPart,
                            re.S | re.M | re.I)
        if matchObj:
            tempTup = matchObj.span(2)
            # print(idx,matchObj.group(2), ":", rightPart[tempTup[0]:tempTup[1]],";",nameTup[0],nameTup[1])
            tempValue = tempValue + rightPart[:tempTup[0]] + ' '
            rightPart = rightPart[tempTup[1]:]
        else:
            tempValue = tempValue + rightPart
            break

    return tempValue


def procNewUser(toSendTable, userName):
    if '' == userName.strip():
        return toSendTable
    # 如果表格中已经有了该用户，那就无需操作，并不是添加新用户
    if -1 != toSendTable.find(userName):
        return toSendTable
    idxLastEndTr = toSendTable.rfind('</tr>')
    idxLast2ndEndTr = toSendTable.rfind('</tr>', 0,
                                        idxLastEndTr - len('</tr>'))
    idxLast3rdEndTr = toSendTable.rfind('</tr>', 0,
                                        idxLast2ndEndTr - len('</tr>'))

    # print("00000000000",toSendTable[idxLast3rdEndTr+3:idxLast3rdEndTr+8])
    # print("00000000000",toSendTable[idxLast3rdEndTr+3:idxLast3rdEndTr+8].encode())
    # 其实下面4个len可以不用加，加上是为了直观理解
    tempValue = toSendTable[:idxLast2ndEndTr +
                            len('</tr>')] + procNewUserStatistic(
                                toSendTable[idxLast3rdEndTr +
                                            len('</tr>'):idxLast2ndEndTr +
                                            len('</tr>')],
                                userName) + toSendTable[idxLast2ndEndTr +
                                                        len('</tr>'):]
    return tempValue


def procDeleteUser(toSendTable, userName):
    if '' == userName.strip():
        return toSendTable
    # 如果表格中没有该用户，那就无需操作
    idxUserName = toSendTable.find(userName)
    if -1 == idxUserName:
        return toSendTable

    idxCurrRowStart = toSendTable.rfind('<tr', 0, idxUserName - len(userName))
    idxNextRowStart = toSendTable.find('<tr', idxUserName + len(userName))
    if -1 == idxCurrRowStart or -1 == idxNextRowStart:
        return toSendTable

    tempValue = toSendTable[:idxCurrRowStart +
                            len('<tr')] + toSendTable[idxNextRowStart:]
    return tempValue


def drawSlant(sendTableToAddSlant):
    if -1 == sendTableToAddSlant.find('id="td1"'):
        idxFirstTd = sendTableToAddSlant.find('<td')
        idx2ndTd = sendTableToAddSlant.find('<td', idxFirstTd + len('<td'))
        if -1 == idx2ndTd:
            return sendTableToAddSlant

        sendTableToAddSlant = sendTableToAddSlant[:idx2ndTd + len(
            '<td')] + ' id="td1" ' + sendTableToAddSlant[idx2ndTd +
                                                         len('<td'):]
    # 加上画斜线的控制
    # line函数中的if判断必须另起一行，所以在if前加了\n
    sendTableToAddSlant = '<script Language="javascript">\
	function a(x, y, color) {\
		document.write("<img border=\'0\' style=\'position: absolute; left: "\
				+ (x + 8) + "; top: " + (y + 8) + ";background-color: "\
				+ color + "\' src=\'px.gif\' width=1 height=1>")\
	} </script>' + sendTableToAddSlant + '<script>\
    function line(x1, y1, x2, y2, color) {\
        var tmp\n\
        if (x1 >= x2) {\
            tmp = x1;\
            x1 = x2;\
            x2 = tmp;\
            tmp = y1;\
            y1 = y2;\
            y2 = tmp;\
        }\
        for (var i = x1; i <= x2; i++) {\
            x = i;\
            y = (y2 - y1) / (x2 - x1) * (x - x1) + y1;\
            a(x, y, color);\
        }\
    }\
    line(td1.offsetLeft, td1.offsetTop, td1.offsetLeft + td1.offsetWidth, td1.offsetTop + td1.offsetHeight, "#DC143C") </script>'

    return sendTableToAddSlant


def generateHard(dictHard):
    hardsTitle = '<div style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-variant-ligatures: normal; clear: both;"><span style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-family:微软雅黑; font-variant-ligatures: normal; font-weight: 700;">2、研发周工作“风险与困难”反馈统计：</span></div>'
    preDecorate = '<div style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-variant-ligatures: normal; clear: both;">'
    if len(dictHard):
        return hardsTitle + preDecorate + '无' + '</div>'

    retValue = hardsTitle
    for k, v in dictHard.items():
        tempValue = retValue + preDecorate + AllStaffs[k] + ': ' + '</div>'
        retValue = tempValue
        for hard in v:
            tempValue = retValue + preDecorate + hard + '</div>'
            retValue = tempValue

    return retValue


def generateSuggest(dictSuggest):
    suggestsTitle = '<div style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-variant-ligatures: normal; clear: both;"><span style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-family: 微软雅黑; font-variant-ligatures: normal; font-weight: 700;">3、研发周工作“意见与建议”反馈统计：</span></div>'
    preDecorate = '<div style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-variant-ligatures: normal; clear: both;">'
    if len(dictSuggest):
        return suggestsTitle + preDecorate + '无' + '</div>'

    retValue = suggestsTitle
    for k, v in dictSuggest.items():
        tempValue = retValue + preDecorate + AllStaffs[k] + ': ' + '</div>'
        retValue = tempValue
        for hard in v:
            tempValue = retValue + preDecorate + hard + '</div>'
            retValue = tempValue

    return retValue


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


def getCurrentWeekWorkDaySpan():
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
    # _12.09-2019.12.14
    return '周报统计_' + getCurrentWeekWorkDaySpan()


def composeEmail(tableContent, hardInfo, suggestInfo):
    # 发送的表格的提示信息
    sentDescrip = '<div style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-variant-ligatures: normal; clear: both;"><span style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-family:微软雅黑; font-variant-ligatures: normal; font-weight: 700;">1、周报发送情况详细结果如下：</span></div>'
    retContent = ''
    blankLine = '<div><br></div>'
    wrapInfoStart = '<div style="font-family: ' + 'Microsoft YaHei UI' + ', Tahoma; line-height: normal; clear: both;">'
    wrapInfoEnd = '</div>'
    retContent = retContent + '<body>Hi，各位好:<br/>以下是本周研发周报统计情况反馈，请查阅，谢谢！！'
    # 添加两个空白行
    retContent = retContent + blankLine + blankLine
    retContent = retContent + sentDescrip
    retContent = retContent + tableContent
    retContent = retContent + "</body>"
    # 准备添加hard和suggest，先添加一个空白行
    retContent = retContent + blankLine

    retContent = retContent + wrapInfoStart
    # 开始添加hard
    retContent = retContent + hardInfo
    retContent = retContent + blankLine
    retContent = retContent + suggestInfo
    retContent = retContent + wrapInfoEnd
    return retContent


def sendEmail(bodyContent):
    if '' == bodyContent.strip():
        return

    msg = MIMEText(bodyContent, 'html', 'utf-8')
    msg['From'] = _format_addr('周报统计机器人 <%s>' % user)
    msg['To'] = _format_addr('研发部 <%s>' % to_addr)
    msg['Subject'] = Header(getEmailSubject(), 'utf-8').encode()
    # server = smtplib.SMTP(smtp_server, 25)
    server = smtplib.SMTP_SSL(smtp_server, 465)
    # server.set_debuglevel(1)
    server.login(user, password)
    server.sendmail(user, [to_addr], msg.as_string())
    server.quit()


def sendEmailWithPic(bodyContent):
    if '' == bodyContent.strip():
        return

    msg = MIMEMultipart('related')
    content = MIMEText(bodyContent, 'html', 'utf-8')

    msg['From'] = _format_addr('周报统计机器人 <%s>' % user)
    msg['To'] = _format_addr('研发部 <%s>' % to_addr)
    msg['Subject'] = Header(getEmailSubject(), 'utf-8').encode()
    # 如果有编码格式问题导致乱码，可以进行格式转换：
    # content = content.decode('utf-8').encode('gbk')
    msg.attach(content)

    # fp = open('C:\\Users\\test\\Desktop\\cha.png', 'rb')
    fp = open('/usr/sbin/cha.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', 'image1')  # 这个id用于上面html获取图片
    msg.attach(msgImage)

    # server = smtplib.SMTP(smtp_server, 25)
    server = smtplib.SMTP_SSL(smtp_server, 465)
    # server.set_debuglevel(1)
    server.login(user, password)
    server.sendmail(user, [to_addr], msg.as_string())
    server.quit()


def sendWeekly():
    imapClient = loginEmailbox()

    folders = listFolders(imapClient)
    logging.info(u'包含文件夹如下:')
    for folder in folders:
        logging.info(folder[-1])

    toSendWeeklyTable = getLastWeeklyStatisticTable(imapClient)
    if len(toSendWeeklyTable) == 0:
        logging.warning(u'获取上周周报统计表格失败:')
        return

    mHard = {'': []}
    mSuggest = {'': []}
    lSent = []

    msgDicts = fetchAllWeeklyEmails(imapClient)
    for msgId, msgDict in msgDicts.items():
        subject, mailFrom, mailTo, hardContents, subjectContents = parseMsgDict(
            msgDict)
        # 只关注周报邮件，跳过普通邮件
        if hardContents == None or subjectContents == None:
            continue
        # 某些同事发了图片格式的周报，解析不了'困难'和'建议'就都是''，不是周报的这两项都是None就不处理
        # 后面再更新新用户的发送标记，因为新用户还没有加到table中
        if hardContents == '' or subjectContents == '':
            # 周报可能是图片格式
            toSendWeeklyTable = updateUserSentRow(toSendWeeklyTable,
                                                  AllStaffs[mailFrom], '✅')
            continue
        # 打印周报详情：发件人、困难、建议
        # print(mailFrom, hardContents, subjectContent)
        if mailFrom not in lSent:
            lSent.append(mailFrom)

        # '' or [] or ['无']
        hardLen = len(hardContents)
        if 0 == hardLen:
            pass
        elif 1 == hardLen:
            singleHard = hardContents[0]
            singleHard = singleHard.strip()
            if '无' == singleHard or '暂无' == singleHard or '没有' == singleHard or '暂时没有' == singleHard:
                pass
            else:
                mHard[mailFrom] = hardContents
        else:
            mHard[mailFrom] = hardContents

        suggestLen = len(subjectContents)
        if 0 == suggestLen:
            pass
        elif 1 == suggestLen:
            singleSuggest = subjectContents[0]
            singleSuggest = singleSuggest.strip()
            if '无' == singleSuggest or '暂无' == singleSuggest or '没有' == singleSuggest or '暂时没有' == singleSuggest:
                pass
            else:
                mSuggest[mailFrom] = subjectContents
        else:
            mSuggest[mailFrom] = subjectContents

        toSendWeeklyTable = updateUserSentRow(toSendWeeklyTable,
                                              AllStaffs[mailFrom], '✅')
        if '' == toSendWeeklyTable.strip():
            logging.warning(u'更新用户发送标记失败: ')
            logging.warning(mailFrom)
            return

    # 0、新用户列表
    # newUsers = ['杜蓉蓉','王菲飞']
    newUsers = []
    # 1、添加新用户
    for newUser in newUsers:
        toSendWeeklyTable = procNewUser(toSendWeeklyTable, newUser)
    # 2、更新新用户的发送标记
    for newUser in newUsers:
        if newUser in lSent:
            toSendWeeklyTable = updateUserSentRow(toSendWeeklyTable, newUser,
                                                  '✅')

    # 0、待删除用户列表，以后无需统计
    # toDeleteUsers = ['杜蓉蓉','王菲飞']
    toDeleteUsers = []
    # 1、删除用户行信息
    for delUser in toDeleteUsers:
        toSendWeeklyTable = procDeleteUser(toSendWeeklyTable, delUser)

    # 更新未发送周报标记
    for k, v in AllStaffs.items():
        if k not in lSent:
            toSendWeeklyTable = updateUserNoSentRow(toSendWeeklyTable,
                                                    AllStaffs[k])

    # 更新发送率
    sentPercent = get_two_float(100 * len(lSent) / len(AllStaffs), 2)
    toSendWeeklyTable = updateUserSentRow(toSendWeeklyTable, '发送率',
                                          sentPercent + '%')

    # procUserHoliday
    toSendWeeklyTable = updateUserSentRow(toSendWeeklyTable, '', '请假')

    # 单元格斜线
    toSendWeeklyTable = drawSlant(toSendWeeklyTable)

    # 发送周报
    sendEmailWithPic(composeEmail(toSendWeeklyTable,generateHard(mHard),generateSuggest(mSuggest)))
    # file_name = 'write.html'
    # with open(file_name,'a',encoding='utf-8') as file_obj:
    #     file_obj.write(toSendWeeklyTable)

    logoutEmailbox(imapClient)


if __name__ == '__main__':
    sendWeekly()
