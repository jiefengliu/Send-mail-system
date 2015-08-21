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
    #配置文件
    def __init__(self):
        self.useraddress = ''
        self.user_password = ''
        self.mail_host = ''


    def readFileFormal(self,fileName):

        fileName = unicode(fileName)
        data= xlrd.open_workbook(fileName)  # 读excel文件
        table = data.sheet_by_name(u'Sheet1')  # 获取一个工作表 有多种方法
        rownum = table.nrows
        colnum = table.ncols
        flag = False
            #获取表格的标题
        title = table.cell(0,0).value

        if colnum != 20:
            flag = True
        return flag

    '''
    #获取excel文件信息
    '''
    def readExcelFile(self,fileName):
        fileName = unicode(fileName)
        data= xlrd.open_workbook(fileName)  # 读excel文件
        table = data.sheet_by_name(u'Sheet1')  # 获取一个工作表 有多种方法
        rownum = table.nrows
        colnum = table.ncols

        alldata = []  #所有的值
        HeadInfo = []  #表头信息
        deductInfo = [] #扣费

        #获取表格的标题
        title = table.cell(0,0).value
        # print title
        #获取表头信息   这里要对应关系
        for j in range(colnum):
             if (2<=j and j<=6) or (10<=j and j<=11) or(13 <= j and j<=17) :
                 if (j==2 or j==10 or j==13):
                      deductInfo.append(table.row(1)[j].value)   #记录代扣和应扣
                 temp = table.row(2)[j].value
             else :
                 temp = table.row(1)[j].value
             HeadInfo.append(temp)

        #还没处理异常的情况
        #获取所有的值
        for i in range(3,rownum):
            rowValue = table.row_values(i)
            alldata.append(rowValue)
        # print (alldata)

        return HeadInfo,deductInfo,alldata,title


        #将execl中数据形成html格式

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
                <th class="tg-031e" rowspan="2">姓名</th>
                <th class="tg-s6z2" rowspan="2">基本工资</th>
                <th class="tg-s6z2" colspan="5">岗位津贴</th>
                <th class="tg-031e" rowspan="2">绩效奖金</th>
                <th class="tg-031e" rowspan="2">加班工资</th>
                <th class="tg-031e" rowspan="2">补贴</th>
                <th class="tg-s6z2" colspan="2">应扣</th>
                <th class="tg-031e" rowspan="2">应发<br>合计</th>
                <th class="tg-s6z2" colspan="5">代扣代缴</th>
                <th class="tg-031e" rowspan="2">实发工资</th>
                <th class="tg-031e" rowspan="2">邮件地址</th>
            </tr>
            <tr>
                <td class="tg-031e">职务津贴</td>
                <td class="tg-031e">交通津贴</td>
                <td class="tg-031e">话费津贴</td>
                <td class="tg-031e">电脑津贴</td>
                <td class="tg-031e">其他津贴</td>
                <td class="tg-031e">考勤</td>
                <td class="tg-031e">其他</td>
                <td class="tg-031e">社保</td>
                <td class="tg-031e">公积金</td>
                <td class="tg-031e">个税</td>
                <td class="tg-031e">其他</td>
                <td class="tg-031e">代扣合计</td>
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


        #发送邮件

    def send_mail(self,mail_host,mail_user,mail_password,receiveAddress,subject,content):  #receiveAdd：收件人；subject：主题；content：邮件内容
        flag = False
        info = ''
        sender=mail_user+"<"+mail_user+">"
        msg = MIMEText(content,_subtype='html',_charset='utf-8')    #创建一个实例，这里设置为html格式邮件
        msg['Subject'] = subject    #设置主题
        msg['From'] = sender
        msg['To'] = receiveAddress
        try:
            s = smtplib.SMTP()
            con  = s.connect(mail_host)  #连接smtp服务器
            login =  s.login(mail_user,mail_password)  #登陆服务器
            send = s.sendmail(sender, receiveAddress, msg.as_string())  #发送邮件
            flag = True
            s.close()

        except Exception, e:
            info = str(e)
            flag = False
        return flag,info





    def start(self,mailhost,mailuser,mailpassword,filename):

        headinfo,deductinfo,alldata,title= self.readExcelFile(filename)    #读excel文件
        stateinfo=[]
        for i in range(len(alldata)):
            rowdata = alldata[i]
            content = self.produceHtml(headinfo,deductinfo,rowdata)        #html格式文件
            receiveraddress = rowdata[len(rowdata)-1]
            # print receiveraddress
            if receiveraddress != '':
                 flag,info = self.send_mail(mailhost,mailuser,mailpassword,receiveraddress,title,content)
                 # print flag
                 if flag:
                     # print  receiveadd+"    "+"发送成功"+"    "+ time.ctime(time.time())
                     sendinfo = receiveraddress+"    "+u"发送成功"+"    "+ time.ctime(time.time())
                     # print sendinfo
                 else:
                     sendinfo = receiveraddress+"    "+u"发送失败"+"    "+ time.ctime(time.time())
                     # print sendinfo
            else:
                #print ('此联系人地址不存在！！')
                sendinfo = u'联系人地址不存在！！'
            stateinfo.append(unicode(sendinfo))
        self.producetxt(stateinfo)

    #生成日志文件 主要是记录发送邮件的时间和信息
    def producetxt(self,info):

        timeinfo =  time.strftime('%Y-%m-%d(%X)',time.localtime(time.time()))
        filename = os.getcwd()
        filename =filename +u'/'+timeinfo + u'.txt'
        filename = u'logfile'+u'.txt'
        f = open(filename,'a')
        for i in range(len(info)):
             f.write(info[i]+'\n')
        f.close()


    def issendmail(self,mail_host,mail_user,mail_password):  #receiveAdd：收件人；subject：主题；content：邮件内容
        flag = False
        info = ''
        try:
            s = smtplib.SMTP()
            s.connect(mail_host)  #连接smtp服务器
            xy =  s.login(mail_user,mail_password)  #登陆服务器
            s.close()
            flag = True
            info = ''
        except Exception, e:
            #print str(e)
            info = str(e)
            flag = False
        return flag,info

    #检验发件者的邮件是否可以登入到服务器，判断验证是否成功
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

    '''初始化基本账号信息'''
    def initValue(self,smtphost,user,password,filenameStr):
        self.user_address = user.strip()        #用户账号
        self.user_password = password.strip()   #用户密码
        self.mail_host = smtphost.strip()       #smtp地址
        self.excelFilename = filenameStr.strip()    #读取的文件名
        # self.start(self.mail_host,self.user_address,self.user_password,self.excelFilename)    #开始



'''
if __name__ == "__main__":
    # print
    #isExisting()

    mail_host=u"smtp.qq.com"  #设置服务器
    mail_user=u"test@yulintu.com"    #用户名
    mail_password=u"yfb123456"   #口令
    filename = u'副本工资条模板(1).xls'


    myClass = EmailHandle()
    myClass.start(mail_host,mail_user,mail_password,filename)
    flag,info  = myClass.isExisting(mail_host,mail_user,mail_password)
    print flag
    print info
'''





