#encoding:utf-8
'''
@File    :   weeklyStatistic.py
@Time    :   2019/11/30 15:33:15
@Author  :   wang
@Version :   1.0
@Contact :   412824500@qq.com
'''
# Start typing your code from here

import configparser
import os
import oss2
import requests
import sys
import time

ossUri=""
AccessKeyId=""
AccessKeySecret=""
SecurityToken=""
listBucketName=[]
listObjectNamePrefix=[]
ossLocalDir=""
currDateOssDir=""

def syncOssFiles():
    global listObjectNamePrefix
    for k in listObjectNamePrefix:
        syncOssWithObjectNamePrefix(currDateOssDir,k)

def syncOssWithObjectNamePrefix(parentDir,objPrefix):
    print(parentDir+"/"+objPrefix)
    if False == checkAndCreateDir(parentDir+"/"+objPrefix):
        return



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
    global ossUri,listObjectNamePrefix,ossLocalDir
    config = configparser.ConfigParser()
    config.read("d:/selfProject/pythonWork/ossCfg.ini")
    print('sections:' , ' ' , config.sections())
    ossUri=config.get("default","ossUri")
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
