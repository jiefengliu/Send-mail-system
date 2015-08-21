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
        headinfo,deductinfo,alldata,title= self.EmailHandle.readExcelFile(self.EmailHandle.excelFilename)    #读excel文件
        stateinfo=[]
        for i in range(len(alldata)):

            rowdata = alldata[i]
            content = self.EmailHandle.produceHtml(headinfo,deductinfo,rowdata)        #html格式文件
            receiveraddress = rowdata[len(rowdata)-1]
            # print receiveraddress
            if receiveraddress != '':
                 flag,info = self.EmailHandle.send_mail(self.smtp_host,self.mail_user,self.user_password,receiveraddress,title,content)
                 if flag:
                    #print  receiveadd+"    "+"发送成功"+"    "+ time.ctime(time.time())
                     sendinfo = u"向 "+str(rowdata[0])+u"  发送成功     "+receiveraddress+u"    "+ time.ctime(time.time())
                 else:
                     sendinfo = info +u"  所以向"+ str(rowdata[0])+u"   发送失败      "+receiveraddress+u"    "+ time.ctime(time.time())
                     # print info[1],len(info),info[2]
            else:
                #print ('此联系人地址不存在！！')
                sendinfo = str(rowdata[0])+u' 没有邮件地址没有发送！！'
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
            title = u"提示信息"
            text = u"某个文本框为空或者输入格式不正确,请正确输入!!  \n Yes 清空文本框 \n No 保持不变 "
            ret = self.messageDialogShow(title,text)
            if True == ret:
                self.combo.SetValue('')
                return
            else:
                x = 1
                return

        if self.EmailHandle.readFileFormal(filename):
            title = u"提示信息"
            text = u"请检查您的文件格式是否符合标准格式  \n Yes 清空文件路径文本框内容 \n No 保持不变 "
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
            title = u"提示信息"
            text = u"输入的邮件地址中smtp服务器地址与配置文件中的不一致" \
                   u"是否继续?  \n Yes 继续 \n No 请检测输入是否正确"
            ret = self.messageDialogShow(title,text)
            if True == ret:
                self.smtp_host = newHostAddress
            else:
                x = 1
                return
        #在这里接入信息
        ReadWriteFile.writeFileHandle(addr)

        #print addr,pwd,filename
        flag,info = self.EmailHandle.isExisting(self.smtp_host,addr,pwd)
        if flag:

            self.EmailHandle.initValue(self.smtp_host,addr,pwd,filename)
            self.sendEmail()

            title = u"提示信息"
            text = u"发送完成"
            dialog = wx.MessageDialog(self,text,title, wx.OK)

            if wx.ID_OK == dialog.ShowModal():
                # self.combo.SetValue('')
                self.combo.SetValue('')
                dialog.Destroy()
                return

        else:
            #print '地址无法识别，验证失败，请正确输入地址或者密码'

            title = u"提示信息"
            text = u""+info+\
                   u"\n地址无法识别,登入邮件服务器,验证失败,请正确输入地址或者密码!!  " \
                   u"\n或者请检测输入是否正确或配置文件" \
                   u"\n Yes 清空邮件地址和密码文本框内容 \n No 保持不变 "
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
    frame = MainClass(None,title=u"发送邮件工具")
    frame.Show(True)
    app.MainLoop()
