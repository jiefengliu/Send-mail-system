# encoding: gbk



import sys
import os
import time
import xlrd
import smtplib
from email.mime.text import MIMEText


from email.utils import COMMASPACE,formatdate
from email import encoders
reload(sys)
sys.setdefaultencoding( "utf-8" )

class EmailHandle:
    #�����ļ�
    def __init__(self):
        self.useraddress = ''
        self.user_password = ''
        self.mail_host = ''


    def readFileFormal(self,fileName):

        fileName = unicode(fileName)
        data= xlrd.open_workbook(fileName)  # ��excel�ļ�
        table = data.sheet_by_name(u'Sheet1')  # ��ȡһ�������� �ж��ַ���
        rownum = table.nrows
        colnum = table.ncols
        flag = False
            #��ȡ���ı���
        title = table.cell(0,0).value

        if colnum != 20:
            flag = True
        return flag

    '''
    #��ȡexcel�ļ���Ϣ
    '''
    def readExcelFile(self,fileName):
        fileName = unicode(fileName)
        data= xlrd.open_workbook(fileName)  # ��excel�ļ�
        table = data.sheet_by_name(u'Sheet1')  # ��ȡһ�������� �ж��ַ���
        rownum = table.nrows
        colnum = table.ncols

        alldata = []  #���е�ֵ
        HeadInfo = []  #��ͷ��Ϣ
        deductInfo = [] #�۷�

        #��ȡ���ı���
        title = table.cell(0,0).value
        # print title
        #��ȡ��ͷ��Ϣ   ����Ҫ��Ӧ��ϵ
        for j in range(colnum):
             if (2<=j and j<=6) or (10<=j and j<=11) or(13 <= j and j<=17) :
                 if (j==2 or j==10 or j==13):
                      deductInfo.append(table.row(1)[j].value)   #��¼���ۺ�Ӧ��
                 temp = table.row(2)[j].value
             else :
                 temp = table.row(1)[j].value
             HeadInfo.append(temp)

        #��û�����쳣�����
        #��ȡ���е�ֵ
        for i in range(3,rownum):
            rowValue = table.row_values(i)
            alldata.append(rowValue)
        # print (alldata)

        return HeadInfo,deductInfo,alldata,title


        #��execl�������γ�html��ʽ

    def produceHtml(self,baseinfo,deductinfo,rowdata):
            #print len(rowdata)
        content = u'''<style type="text/css">
                .tg { border-collapse: collapse;  border-spacing: 0;}
                .tg td {
                        font-family: Arial, sans-serif;
                        font-size: 14px;
                        padding: 10px 5px;
                        border-style: solid;
                        border-width: 1px;
                        overflow: hidden;
                        word-break: normal;}
                    .tg th {
                        font-family: Arial, sans-serif;
                        font-size: 14px;
                        font-weight: normal;
                        padding: 10px 5px;
                        border-style: solid;
                        border-width: 1px;
                        overflow: hidden;
                        word-break: normal;}
                    .tg .tg-s6z2 {
                        text-align: center;}
                </style>
        <table class="tg">
            <tr>
                <th class="tg-031e" rowspan="2">����</th>
                <th class="tg-s6z2" rowspan="2">��������</th>
                <th class="tg-s6z2" colspan="5">��λ����</th>
                <th class="tg-031e" rowspan="2">��Ч����</th>
                <th class="tg-031e" rowspan="2">�Ӱ๤��</th>
                <th class="tg-031e" rowspan="2">����</th>
                <th class="tg-s6z2" colspan="2">Ӧ��</th>
                <th class="tg-031e" rowspan="2">Ӧ��<br>�ϼ�</th>
                <th class="tg-s6z2" colspan="5">���۴���</th>
                <th class="tg-031e" rowspan="2">ʵ������</th>
                <th class="tg-031e" rowspan="2">�ʼ���ַ</th>
            </tr>
            <tr>
                <td class="tg-031e">ְ�����</td>
                <td class="tg-031e">��ͨ����</td>
                <td class="tg-031e">���ѽ���</td>
                <td class="tg-031e">���Խ���</td>
                <td class="tg-031e">��������</td>
                <td class="tg-031e">����</td>
                <td class="tg-031e">����</td>
                <td class="tg-031e">�籣</td>
                <td class="tg-031e">������</td>
                <td class="tg-031e">��˰</td>
                <td class="tg-031e">����</td>
                <td class="tg-031e">���ۺϼ�</td>
            </tr>
            <tr>
                <td class="tg-031e">''' +str(rowdata[0])+ u'''</td>
                <td class="tg-031e">'''+str(rowdata[1])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[2])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[3])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[4])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[5])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[6])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[7])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[8])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[9])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[10])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[11])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[12])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[13])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[14])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[15])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[16])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[17])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[18])+u'''</td>
                <td class="tg-031e">'''+str(rowdata[19])+u'''</td>
            </tr>
        </table>'''

        return content


        #�����ʼ�

    def send_mail(self,mail_host,mail_user,mail_password,receiveAddress,subject,content):  #receiveAdd���ռ��ˣ�subject�����⣻content���ʼ�����
        flag = False
        info = ''
        sender=mail_user+"<"+mail_user+">"
        msg = MIMEText(content,_subtype='html',_charset='utf-8')    #����һ��ʵ������������Ϊhtml��ʽ�ʼ�
        msg['Subject'] = subject    #��������
        msg['From'] = sender
        msg['To'] = receiveAddress
        try:
            s = smtplib.SMTP()
            con  = s.connect(mail_host)  #����smtp������
            login =  s.login(mail_user,mail_password)  #��½������
            send = s.sendmail(sender, receiveAddress, msg.as_string())  #�����ʼ�
            flag = True
            s.close()

        except Exception, e:
            info = str(e)
            flag = False
        return flag,info





    def start(self,mailhost,mailuser,mailpassword,filename):

        headinfo,deductinfo,alldata,title= self.readExcelFile(filename)    #��excel�ļ�
        stateinfo=[]
        for i in range(len(alldata)):
            rowdata = alldata[i]
            content = self.produceHtml(headinfo,deductinfo,rowdata)        #html��ʽ�ļ�
            receiveraddress = rowdata[len(rowdata)-1]
            # print receiveraddress
            if receiveraddress != '':
                 flag,info = self.send_mail(mailhost,mailuser,mailpassword,receiveraddress,title,content)
                 # print flag
                 if flag:
                     # print  receiveadd+"    "+"���ͳɹ�"+"    "+ time.ctime(time.time())
                     sendinfo = receiveraddress+"    "+u"���ͳɹ�"+"    "+ time.ctime(time.time())
                     # print sendinfo
                 else:
                     sendinfo = receiveraddress+"    "+u"����ʧ��"+"    "+ time.ctime(time.time())
                     # print sendinfo
            else:
                #print ('����ϵ�˵�ַ�����ڣ���')
                sendinfo = u'��ϵ�˵�ַ�����ڣ���'
            stateinfo.append(unicode(sendinfo))
        self.producetxt(stateinfo)

    #������־�ļ� ��Ҫ�Ǽ�¼�����ʼ���ʱ�����Ϣ
    def producetxt(self,info):

        timeinfo =  time.strftime('%Y-%m-%d(%X)',time.localtime(time.time()))
        filename = os.getcwd()
        filename =filename +u'/'+timeinfo + u'.txt'
        filename = u'logfile'+u'.txt'
        f = open(filename,'a')
        for i in range(len(info)):
             f.write(info[i]+'\n')
        f.close()


    def issendmail(self,mail_host,mail_user,mail_password):  #receiveAdd���ռ��ˣ�subject�����⣻content���ʼ�����
        flag = False
        info = ''
        try:
            s = smtplib.SMTP()
            s.connect(mail_host)  #����smtp������
            xy =  s.login(mail_user,mail_password)  #��½������
            s.close()
            flag = True
            info = ''
        except Exception, e:
            #print str(e)
            info = str(e)
            flag = False
        return flag,info

    #���鷢���ߵ��ʼ��Ƿ���Ե��뵽���������ж���֤�Ƿ�ɹ�
    def isExisting(self,host,user,password):
        flag = False
        info = ''
        # host = user.strip('')
        # hostlist = host.split('@')
        # host = hostlist[1]
        # host = 'smtp.'+host
        # #print host
        flag ,info= self.issendmail(host,user,password)
        return flag,info

    '''��ʼ�������˺���Ϣ'''
    def initValue(self,smtphost,user,password,filenameStr):
        self.user_address = user.strip()        #�û��˺�
        self.user_password = password.strip()   #�û�����
        self.mail_host = smtphost.strip()       #smtp��ַ
        self.excelFilename = filenameStr.strip()    #��ȡ���ļ���
        # self.start(self.mail_host,self.user_address,self.user_password,self.excelFilename)    #��ʼ



'''
if __name__ == "__main__":
    # print
    #isExisting()

    mail_host=u"smtp.qq.com"  #���÷�����
    mail_user=u"test@yulintu.com"    #�û���
    mail_password=u"yfb123456"   #����
    filename = u'����������ģ��(1).xls'


    myClass = EmailHandle()
    myClass.start(mail_host,mail_user,mail_password,filename)
    flag,info  = myClass.isExisting(mail_host,mail_user,mail_password)
    print flag
    print info
'''





