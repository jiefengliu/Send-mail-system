
#encoding: gbk
#-*- coding: cp936 -*-

import os
import sys
import time


def writeFileHandle(content):
    filename = os.getcwd()
    filename =u'information'+u'.txt'
    filename =filename.strip()

    f = open(filename,'w')
    f.write(content)
    f.flush()
    f.close()


def readFileHandle():
    path = os.getcwd()
    #读文件
    filename = u'information'+u'.txt'
    filename =filename.strip()
    #filename='xxx.txt'
    isTrue = os.path.isfile(u'information.txt')

    # print (isTrue)
    if isTrue:
        f = open(filename)

        line = f.readline()
        #print (line)
        return isTrue,line
        f.close()
    else:
        #print ('文件不存在')
        return isTrue,''


'''
if __name__ == '__main__':

    #写文件：如果在同目录下没有要写的文件名
    #会新建文件，并写入内容
    #如果存在文件，直接写入内容
    writeFileHandle(u'iphonexb@163.com')
    readFileHandle()

'''




