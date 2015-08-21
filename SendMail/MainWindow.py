#encoding: gbk

import os
import wx
import EmailTest

class MainFrame(wx.Frame):
    def __init__(self,parent,title):
        super(MainFrame,self).__init__(parent,title=title,size=(800,550))
        self.initLayout()
        self.Centre()


    def initLayout(self):
        panel = wx.Panel(self)
        sizer = wx.GridBagSizer(5,5)  #5 row 5 column

        infoLable = wx.StaticText(panel,label=u"首次操作可以阅读下帮助信息")
        sizer.Add(infoLable, pos=(0,1),span=(1,2),flag=wx.TOP|wx.RIGHT|wx.BOTTOM|wx.LEFT|wx.EXPAND,border=10)
        self.helpBtn = wx.Button(panel, label=u"帮助信息")
        sizer.Add(self.helpBtn, pos=(0, 3),flag=wx.TOP|wx.RIGHT|wx.BOTTOM|wx.LEFT|wx.EXPAND,border=10)
        line = wx.StaticLine(panel)
        sizer.Add(line,pos=(1,0),span=(1,5),flag=wx.EXPAND|wx.BOTTOM,border=10)


        sendAddressLable = wx.StaticText(panel,label=u"邮件地址")
        sizer.Add(sendAddressLable,pos=(2,0),flag=wx.LEFT,border=40)
        self.addressText = wx.TextCtrl(panel)
        sizer.Add(self.addressText,pos=(2,1),span=(1,3),flag=wx.EXPAND)

        pwdLable = wx.StaticText(panel, label=u"邮件密码")
        sizer.Add(pwdLable, pos=(3, 0), flag=wx.LEFT, border=40)
        self.pwdText = wx.TextCtrl(panel,-1,style=wx.TE_PASSWORD)
        sizer.Add(self.pwdText, pos=(3, 1), span=(1, 3), flag=wx.EXPAND,border=5)


        filePathLable = wx.StaticText(panel, label=u"文件路径")
        sizer.Add(filePathLable, pos=(4, 0), flag=wx.LEFT, border=40)
        self.combo = wx.TextCtrl(panel)
        sizer.Add(self.combo, pos=(4, 1), span=(1, 3),flag=wx.EXPAND, border=5)


        self.chooseBtn = wx.Button(panel, label=u"选择文件夹")
        sizer.Add(self.chooseBtn, pos=(4, 4), flag=wx.RIGHT, border=20)


        self.sendBtn = wx.Button(panel, label=u"发送")
        sizer.Add(self.sendBtn, pos=(5, 1),flag=wx.TOP,border=5)
        self.exitBtn = wx.Button(panel, label=u"退出")
        sizer.Add(self.exitBtn, pos=(5, 3),flag=wx.TOP|wx.RIGHT, border=5)



        line1 = wx.StaticLine(panel)
        sizer.Add(line1,pos=(6,0),span=(1,5),flag=wx.EXPAND|wx.TOP,border=10)


        self.reviewBtn = wx.Button(panel, label=u'文件预览')
        sizer.Add(self.reviewBtn, pos=(7, 0), flag=wx.LEFT, border=20)
        self.clearBtn = wx.Button(panel, label=u'清空内容')
        sizer.Add(self.clearBtn, pos=(7, 1), flag=wx.LEFT, border=5)
        self.reviewText = wx.TextCtrl(panel,style=wx.TE_MULTILINE|wx.TE_READONLY,size=(600,400))
        sizer.Add(self.reviewText,pos=(8,0),span=(7,5),flag=wx.RIGHT|wx.BOTTOM|wx.LEFT|wx.EXPAND,border=20)
        sizer.AddGrowableCol(2)
        panel.SetSizer(sizer)


        #绑定事件
        self.Bind(wx.EVT_BUTTON,self.helpOnClick,self.helpBtn)
        self.Bind(wx.EVT_BUTTON,self.chooseOnClick,self.chooseBtn)
        self.Bind(wx.EVT_BUTTON,self.sendOnClick,self.sendBtn)
        self.Bind(wx.EVT_BUTTON,self.exitOnClick,self.exitBtn)
        self.Bind(wx.EVT_BUTTON,self.reViewOnClick,self.reviewBtn)
        self.Bind(wx.EVT_BUTTON,self.clearOnClick,self.clearBtn)


    def defaultFileDialogOptions(self):
        ''' Return a dictionary with file dialog options that can be
            used in both the save file dialog as well as in the open
            file dialog. '''
        return dict(message=u'Choose a file', defaultDir=self.dirname,
                    wildcard=u'*.*')

    def askUserForFilename(self, **dialogOptions):
        dialog = wx.FileDialog(self, **dialogOptions)
        if dialog.ShowModal() == wx.ID_OK:
            userProvidedFilename = True
            self.filename = dialog.GetFilename()
            self.dirname = dialog.GetDirectory()
            # print self.filename
            # self.SetTitle() # Update the window title with the new filename
        else:
            userProvidedFilename = False
        dialog.Destroy()
        return userProvidedFilename

    def chooseOnClick(self,e):
        self.filename = ''
        self.dirname = '.'
        if self.askUserForFilename(style=wx.OPEN,**self.defaultFileDialogOptions()):
            compeleteFilePath = os.path.join(self.dirname, self.filename)
            # print compeleteFilePath
            self.combo.SetValue(compeleteFilePath)
            return  compeleteFilePath
        else:
            self.combo.SetValue('')
            return self.filename



    def helpOnClick(self, event):
        message = u"1. 在发送前,请确认文件内容格式是否为标准格式;\n " \
                  u"2. 同时,确保配置文件中的邮件地址,密码或SMTP地址正确;" \
                  u"\n3. 程序启动时会根据配置文件自动将基本信息显示到输入框中;" \
                  u"\n4. 如果您感觉输入框中的信息不正确,可以自动输入正确的信息;"
        title = u"提示信息"
        dialog = wx.MessageDialog(self, message, title, wx.OK)
        dialog.ShowModal()
        dialog.Destroy()

    def sendOnClick(self,e):
        x = 1

    def exitOnClick(self,e):
        self.Close()

    def reViewOnClick(self,e):

        self.filename = ''
        self.dirname = '.'
        if self.askUserForFilename(style=wx.OPEN,**self.defaultFileDialogOptions()):
            compeleteFile = os.path.join(self.dirname, self.filename)
            if -1 != compeleteFile.find(u'.xls'):
                fileClass = EmailTest.EmailHandle()
                headinfo,deductinfo,alldata,title= fileClass.readExcelFile(compeleteFile)    #读excel文件

                self.reviewText.AppendText(u'%s\n'%str(title))
                for j in range(len(headinfo)):
                    if j == 8 :
                        x = 1
                    else:
                        self.reviewText.AppendText(unicode(str(headinfo[j])+"    "))
                for i in range(len(alldata)):
                    temp = alldata[i]
                    self.reviewText.AppendText(u'\n')
                    for j in range(len(temp)):
                        if j == 8:
                            x=1
                        elif j==0:
                            self.reviewText.AppendText(unicode(str(temp[j])+"    "))
                        elif j==2:
                            self.reviewText.AppendText(unicode(str(temp[j])+"       "))
                        elif j==3:
                            self.reviewText.AppendText(unicode(str(temp[j])+"            "))
                        elif j==6:
                            self.reviewText.AppendText(unicode(str(temp[j])+"   "))
                        else:
                            self.reviewText.AppendText(unicode(str(temp[j])+"      "))
            else:
                textfile = open(os.path.join(self.dirname, self.filename), 'r')
                self.reviewText.SetValue(unicode(textfile.read()))
                textfile.close()

    def clearOnClick(self,e):
        self.reviewText.SetValue('')




'''
if __name__ == '__main__':
    app = wx.App()
    frame = MainFrame(None,title=u"发送邮件")
    frame.Show(True)
    app.MainLoop()

'''
