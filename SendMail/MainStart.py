#encoding: gbk



import wx
import ConfigParser
import time
import MainWindow
import EmailTest
import ReadWriteFile



class MainClass(MainWindow.MainFrame):
    def __init__(self,parent,title):
        super(MainWindow.MainFrame,self).__init__(parent,title=title,size=(800,550))
        self.initLayout()
        self.Centre()
        config = ConfigParser.ConfigParser()
        profile = 'configuration.ini'
        config.read(profile)
        self.smtp_host = config.get('Variable','smtp_host')
        # print self.smtp_host
        self.mail_user = config.get('Variable','mail_user')
        # print self.mail_user
        self.user_password = config.get('Variable','password')
        # print self.user_password
        self.addressText.SetValue(self.mail_user)
        self.pwdText.SetValue(self.user_password)
        self.EmailHandle = EmailTest.EmailHandle()


    def handleSmtpAddress(self,userAddress):
        host = userAddress.strip()
        hostlist = host.split('@')
        host = hostlist[1]
        host = 'smtp.'+host
        return host

    def sendOnClick(self,e):
        # print
        self.reviewText.SetValue('')
        self.reviewText.Clear()
        self.testify()

    def messageDialogShow(self,title,message):
        dialogTitle = title
        dialogMessage = message
        dlg = wx.MessageDialog(self,dialogMessage,dialogTitle,wx.YES_NO|wx.ICON_QUESTION)
        if wx.ID_YES == dlg.ShowModal():
            ret = True
        else:
            ret = False
        dlg.Destroy()
        return ret

    def sendEmail(self):
        headinfo,deductinfo,alldata,title= self.EmailHandle.readExcelFile(self.EmailHandle.excelFilename)    #��excel�ļ�
        stateinfo=[]
        for i in range(len(alldata)):

            rowdata = alldata[i]
            content = self.EmailHandle.produceHtml(headinfo,deductinfo,rowdata)        #html��ʽ�ļ�
            receiveraddress = rowdata[len(rowdata)-1]
            # print receiveraddress
            if receiveraddress != '':
                 flag,info = self.EmailHandle.send_mail(self.smtp_host,self.mail_user,self.user_password,receiveraddress,title,content)
                 if flag:
                    #print  receiveadd+"    "+"���ͳɹ�"+"    "+ time.ctime(time.time())
                     sendinfo = u"�� "+str(rowdata[0])+u"  ���ͳɹ�     "+receiveraddress+u"    "+ time.ctime(time.time())
                 else:
                     sendinfo = info +u"  ������"+ str(rowdata[0])+u"   ����ʧ��      "+receiveraddress+u"    "+ time.ctime(time.time())
                     # print info[1],len(info),info[2]
            else:
                #print ('����ϵ�˵�ַ�����ڣ���')
                sendinfo = str(rowdata[0])+u' û���ʼ���ַû�з��ͣ���'
            stateinfo.append(unicode(sendinfo))
            self.showInfo(sendinfo)
        self.EmailHandle.producetxt(stateinfo)

    def showInfo(self,sendinfo):
        self.reviewText.AppendText(u'%s\n'%str(sendinfo))

    def testify(self):

        addr = str(self.addressText.GetValue()).strip()
        pwd = str(self.pwdText.GetValue()).strip()
        filename = str(self.combo.GetValue()).strip()
        isflag = addr.split('@')
        #print len(isflag)
        isfile = filename.split('.xls')
        #print isfile

        if addr == '' or pwd == '' or filename =='' or len(isflag) ==1 or len(isfile) ==1:
            title = u"��ʾ��Ϣ"
            text = u"ĳ���ı���Ϊ�ջ��������ʽ����ȷ,����ȷ����!!  \n Yes ����ı��� \n No ���ֲ��� "
            ret = self.messageDialogShow(title,text)
            if True == ret:
                self.combo.SetValue('')
                return
            else:
                x = 1
                return

        if self.EmailHandle.readFileFormal(filename):
            title = u"��ʾ��Ϣ"
            text = u"���������ļ���ʽ�Ƿ���ϱ�׼��ʽ  \n Yes ����ļ�·���ı������� \n No ���ֲ��� "
            ret = self.messageDialogShow(title,text)
            if True == ret:
                self.combo.SetValue('')
                return
            else:
                x = 1
                return
        # newUserAddress = str(self.addressText.GetValue())
        # newHostAddress = self.handleSmtpAddress(newUserAddress.strip())

        newHostAddress = self.smtp_host
        if newHostAddress  == self.smtp_host:
            x = 1
        else:
            title = u"��ʾ��Ϣ"
            text = u"������ʼ���ַ��smtp��������ַ�������ļ��еĲ�һ��" \
                   u"�Ƿ����?  \n Yes ���� \n No ���������Ƿ���ȷ"
            ret = self.messageDialogShow(title,text)
            if True == ret:
                self.smtp_host = newHostAddress
            else:
                x = 1
                return
        #�����������Ϣ
        ReadWriteFile.writeFileHandle(addr)

        #print addr,pwd,filename
        flag,info = self.EmailHandle.isExisting(self.smtp_host,addr,pwd)
        if flag:

            self.EmailHandle.initValue(self.smtp_host,addr,pwd,filename)
            self.sendEmail()

            title = u"��ʾ��Ϣ"
            text = u"�������"
            dialog = wx.MessageDialog(self,text,title, wx.OK)

            if wx.ID_OK == dialog.ShowModal():
                # self.combo.SetValue('')
                self.combo.SetValue('')
                dialog.Destroy()
                return

        else:
            #print '��ַ�޷�ʶ����֤ʧ�ܣ�����ȷ�����ַ��������'

            title = u"��ʾ��Ϣ"
            text = u""+info+\
                   u"\n��ַ�޷�ʶ��,�����ʼ�������,��֤ʧ��,����ȷ�����ַ��������!!  " \
                   u"\n�������������Ƿ���ȷ�������ļ�" \
                   u"\n Yes ����ʼ���ַ�������ı������� \n No ���ֲ��� "
            ret = self.messageDialogShow(title,text)
            if True == ret:
                self.addressText.SetValue('')
                self.pwdText.SetValue('')
                return
            else:
                x = 1
                return


if __name__ == '__main__':
    app = wx.App()
    frame = MainClass(None,title=u"�����ʼ�����")
    frame.Show(True)
    app.MainLoop()
