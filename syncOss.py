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
import shutil
import sys
import time

ossUri=""
AccessKeyId=""
AccessKeySecret=""
SecurityToken=""

Endpoint=""
BucketName=""
listBucketName=[]
listObjectNamePrefix=[]
ossLocalDir=""
currDateOssDir=""
reserveRecent=1

def syncOssFiles():
    global listObjectNamePrefix
    for k in listObjectNamePrefix:
        syncOssWithObjectNamePrefix(currDateOssDir,k)

def syncOssWithObjectNamePrefix(parentDir,objPrefix):
    global AccessKeyId,AccessKeySecret,Endpoint,BucketName,SecurityToken
    print(parentDir+"/"+objPrefix)
    if False == checkAndCreateDir(parentDir+"/"+objPrefix):
        return

    auth = oss2.StsAuth(AccessKeyId, AccessKeySecret, SecurityToken)
    bucket = oss2.Bucket(auth, Endpoint, BucketName)
    for obj in oss2.ObjectIterator(bucket, prefix = objPrefix+'/', delimiter = '/'):
        if obj.is_prefix():  # 文件夹
            # 去掉子目录的最后的/，因为上面for中在objPrefix加了/，否则不能正常获取子目录下的文件
            syncOssWithObjectNamePrefix(parentDir,obj.key[0:len(obj.key)-1])
        else:                # 文件
            bucket.get_object_to_file(obj.key, parentDir+'/'+obj.key)

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
    global ossUri,listObjectNamePrefix,ossLocalDir,Endpoint,BucketName,reserveRecent
    config = configparser.ConfigParser()
    config.read("d:/selfProject/pythonWork/ossCfg.ini")
    print('sections:' , ' ' , config.sections())
    ossUri=config.get("default","ossUri")
    ossLocalDir=os.path.abspath(config.get("default","ossLocalDir"))
    reserveRecent=int(config.get("default","reserveRecent"))
    Endpoint=config.get("common","endpoint")
    BucketName=config.get("common","bucketName")
    listObjectNamePrefix=config.get("common", "objectNamePrefix").split("\n")

    return checkAndCreateDir(ossLocalDir)

def checkAndCreateDir(dir):
    if os.path.exists(dir):
        return True
    else:
        os.makedirs(dir)

    return os.path.exists(dir)

def clearStale():
    global reserveRecent,ossLocalDir
    if reserveRecent<=0:
        return
    listRecentDate=[]
    for i in range(reserveRecent):
        date=time.strftime("%Y%m%d", time.localtime(time.time()-i*24*60*60))
        listRecentDate.append(date)

    print(listRecentDate)
    dirs = os.listdir(ossLocalDir)
    # print(dirs)
    for dirName in dirs:
        if dirName in listRecentDate:
            continue

        absPath=os.path.abspath(ossLocalDir+"/"+dirName)
        if os.path.isdir(absPath):
            shutil.rmtree(absPath)
        else:
            os.remove(absPath)

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
    clearStale()
