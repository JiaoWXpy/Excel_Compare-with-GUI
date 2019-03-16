'''
@Author  :   {RockyChen}

@Contact :   {rocky_chen777@163.com}

@Script:   ${Excel Compare Tool}

@File    :   ${login}.py

@Time    :   ${2019} ${3}${14}

---------------------
'''

import wx
from homepage import MyForm


class LoginFrame(wx.Frame):

    def __init__(self,parent):

        wx.Frame.__init__(self,parent,id=-1,title='Beyond Compare 登录',size=(300,300))
        panel=wx.Panel(self)

        self.ct = 3

        #back_image = 'C:\\Users\\rocky\\Desktop\\background.jpg'
       # to_bmp_image = wx.Image(back_image, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
       # self.bitmap = wx.StaticBitmap(self, -1, to_bmp_image, (0, 0))

        #居中
        self.Center()

        # 左上角图片
        icon_1 = wx.Icon(name='C:\\Users\\Administrator\\Desktop\\rocky_exe\\bp.png', type=wx.BITMAP_TYPE_PNG)
        self.SetIcon(icon_1)

        #登录/取消按钮
        button1=wx.Button(parent=panel,id=-1,label=u'登录',pos=(30,150))
        button2=wx.Button(parent=panel,id=-1,label=u'取消',pos=(180,150))

        #Textcrl 文本框
        wx.StaticText(panel,-1,"ID",(30,35))
        wx.StaticText(panel,-1,"Password",(30,95))
        self.Username=wx.TextCtrl(panel,-1,u"请输入ID",(100,30),(175,-1))
        self.Password=wx.TextCtrl(panel,-1,"...",(100,90),(175,-1),wx.TE_PASSWORD)
        self.Bind(wx.EVT_BUTTON,self.Login_in,button1)
        self.Bind(wx.EVT_BUTTON,self.Cancle,button2)

    # -------------------------------------------------------------------
    # def [Login_in] 登录至主界面
    def Login_in(self,evt):
        """
        输入密码正确则登录至主界面，否则还有三次机会输入错误则退出
        :param evt:单击确定
        """

        UserName=self.Username.GetValue()
        PassWord=self.Password.GetValue()

        if self.ct == 0:
            wx.Exit()
        elif (UserName == 'admin') and (PassWord =='rocky'):
            self.Destroy()
            dlg = wx.MessageDialog(None, "登陆成功！" , "成功",
                                   wx.YES_DEFAULT | wx.ICON_INFORMATION)
            if dlg.ShowModal() == wx.ID_YES:
                dlg.Destroy()
            frame2=MyForm()
            frame2.Show(True)
        elif (UserName != 'admin') or (PassWord != "rocky"):
            dlg = wx.MessageDialog(None, "ID或Password输入错误,您还有可以尝试%d次！"%(self.ct), "错误", wx.YES_DEFAULT | wx.ICON_WARNING)
            self.ct -= 1
            if dlg.ShowModal() == wx.ID_YES:
                dlg.Destroy()


    def Cancle(self,evt):
        wx.Exit()

class TwoFrame(wx.Frame):
    def __init__(self,parent=None):
        wx.Frame.__init__(self,parent,id=-1,title='TWO',size=(500,500))
class MyApp(wx.App):
    def OnInit(self):
        self.frame=LoginFrame(parent=None)
        self.SetTopWindow(self.frame)
        self.frame.Show(True)
        return True
if __name__ == '__main__':
    app=MyApp()
    app.MainLoop()