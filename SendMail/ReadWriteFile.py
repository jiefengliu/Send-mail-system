
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
    #���ļ�
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
        #print ('�ļ�������')
        return isTrue,''


'''
if __name__ == '__main__':

    #д�ļ��������ͬĿ¼��û��Ҫд���ļ���
    #���½��ļ�����д������
    #��������ļ���ֱ��д������
    writeFileHandle(u'iphonexb@163.com')
    readFileHandle()

'''




