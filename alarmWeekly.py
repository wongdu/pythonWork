#encoding:utf-8
'''
@File    :   alarmWeekly.py
@Time    :   2019/12/13 10:31:52
@Author  :   wang
@Version :   1.0
@Contact :   wang@imprexion.com.cn
'''
# Start typing your code from here

from datetime import datetime, timedelta
import mysql.connector
import logging

user = 'root'
password = 'xxx'
database = 'alarms'

mysqlConn = None
mysqlCursor = None

logging.basicConfig(
    level=logging.INFO,
    format=
    '%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s: %(message)s',
    datefmt='%a, %d %b %Y %H:%M:%S')


def initMysql(loginUser, passwd, db):
    global mysqlConn, mysqlCursor
    mysqlConn = mysql.connector.connect(user=loginUser, password=passwd, database=db)
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


def getAlarms():
    dtNow = datetime.now()
    sinceDate = datetime.now() - timedelta(days=7)

    global mysqlCursor
    mysqlCursor.execute('select * from event_cases where timestamp > %s',
                        (sinceDate.strftime("%Y-%m-%d %H:%M:%S"), ))
    values = mysqlCursor.fetchall()
    return values


def getAlerts():
    dtNow = datetime.now()
    sinceDate = datetime.now() - timedelta(days=7)

    global mysqlCursor
    mysqlCursor.execute('select endpoint,note,timestamp from alerts where timestamp > %s',
                        (sinceDate.strftime("%Y-%m-%d %H:%M:%S"), ))
    values = mysqlCursor.fetchall()
    return values


def procAlarmWeekly():
    if False == initMysql(user, password, database):
        logging.error(u'连接数据库失败:')
        return False

    alarms = getAlarms()
    alerts = getAlerts()


def sendAlarmWeekly():
    pass


if __name__ == '__main__':
    procAlarmWeekly()