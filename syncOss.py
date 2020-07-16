#encoding:utf-8
'''
@File    :   weeklyStatistic.py
@Time    :   2019/11/30 15:33:15
@Author  :   wang
@Version :   1.0
@Contact :   412824500@qq.com
'''
# Start typing your code from here

from aliyunsdkcore import client
from aliyunsdksts.request.v20150401 import AssumeRoleRequest

import configparser
import json
import os
import oss2
import requests
import sys
import time

ossUri=""
AccessKeyId=""
AccessKeySecret=""
SecurityToken=""
StsRoleArn="acs:ram::1839163195371488:role/xiaoxin"
# sts_role_arn
Endpoint=""
BucketName=""
listBucketName=[]
listObjectNamePrefix=[]
ossLocalDir=""
currDateOssDir=""

class StsToken(object):
    """AssumeRole返回的临时用户密钥
    :param str access_key_id: 临时用户的access key id
    :param str access_key_secret: 临时用户的access key secret
    :param int expiration: 过期时间，UNIX时间，自1970年1月1日UTC零点的秒数
    :param str security_token: 临时用户Token
    :param str request_id: 请求ID
    """
    def __init__(self):
        self.access_key_id = ''
        self.access_key_secret = ''
        self.expiration = 0
        self.security_token = ''
        self.request_id = ''

def syncOssFiles():
    global listObjectNamePrefix
    for k in listObjectNamePrefix:
        syncOssWithObjectNamePrefix(currDateOssDir,k)

def syncOssWithObjectNamePrefix(parentDir,objPrefix):
    global AccessKeyId,AccessKeySecret,StsRoleArn,Endpoint,BucketName,SecurityToken
    print(parentDir+"/"+objPrefix)
    if False == checkAndCreateDir(parentDir+"/"+objPrefix):
        return

    # token = fetch_sts_token(AccessKeyId, AccessKeySecret, StsRoleArn)
    # auth = oss2.StsAuth(token.access_key_id, token.access_key_secret, token.security_token)
    auth = oss2.StsAuth(AccessKeyId, AccessKeySecret, SecurityToken)
    bucket = oss2.Bucket(auth, Endpoint, BucketName)
    # for obj in oss2.ObjectIterator(bucket, prefix = 'objPrefix"+"/', delimiter = '/'):
    for obj in oss2.ObjectIterator(bucket, prefix = objPrefix+'/', delimiter = '/'):
        if obj.is_prefix():  # 文件夹
            print('directory: ' + obj.key)
        else:                # 文件
            print('file: ' + obj.key)


def fetch_sts_token(access_key_id, access_key_secret, role_arn):
    print("access_key_id:",access_key_id)
    print("access_key_secret:",access_key_secret)
    clt = client.AcsClient(access_key_id, access_key_secret, 'cn-shenzhen')
    req = AssumeRoleRequest.AssumeRoleRequest()

    req.set_accept_format('json')
    req.set_RoleArn("acs:ram::1839163195371488:role/xiaoxin")
    req.set_RoleSessionName('session-name')

    # body = clt.do_action_with_exception(req)
    body = clt.do_action(req)

    j = json.loads(oss2.to_unicode(body))
    print(type(j),j)

    token = StsToken()
    token.access_key_id = j['Credentials']['AccessKeyId']
    token.access_key_secret = j['Credentials']['AccessKeySecret']
    token.security_token = j['Credentials']['SecurityToken']
    token.request_id = j['RequestId']
    token.expiration = oss2.utils.to_unixtime(j['Credentials']['Expiration'], '%Y-%m-%dT%H:%M:%SZ')

    return token

def getOssInfo():
    global ossUri
    global AccessKeyId, AccessKeySecret,SecurityToken
    ret = requests.get(ossUri)
    result=ret.json()
    if result["StatusCode"] != 200:
        return False
    AccessKeyId=result["AccessKeyId"]
    AccessKeySecret=result["AccessKeySecret"]
    SecurityToken=result["SecurityToken"]
    return True

def initConfig():
    global ossUri,listObjectNamePrefix,ossLocalDir,Endpoint,BucketName
    config = configparser.ConfigParser()
    config.read("d:/selfProject/pythonWork/ossCfg.ini")
    print('sections:' , ' ' , config.sections())
    ossUri=config.get("default","ossUri")
    Endpoint=config.get("common","endpoint")
    BucketName=config.get("common","bucketName")
    listObjectNamePrefix=config.get("common", "objectNamePrefix").split("\n")
    ossLocalDir=os.path.abspath(config.get("default","ossLocalDir"))

    return checkAndCreateDir(ossLocalDir)

def checkAndCreateDir(dir):
    if os.path.exists(dir):
        return True
    else:
        os.makedirs(dir)

    return os.path.exists(dir)

if __name__ == '__main__':
    print("oss version is:",oss2.__version__)
    if False == initConfig():
        sys.exit()
    if False == getOssInfo():
        sys.exit()

    currDate=time.strftime("%Y%m%d", time.localtime())
    currDateOssDir=ossLocalDir+"/"+currDate
    if False == checkAndCreateDir(currDateOssDir):
        sys.exit()
    print(AccessKeyId)
    print(listObjectNamePrefix)
    syncOssFiles()
