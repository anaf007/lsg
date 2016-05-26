#coding=utf-8


#!/bin/env python
import wx,time
from Controller import *
wildcard = u"表格2007(*.xlsx)|*.xlsx|表格2003(*.xls)|*.xls"
class MainPage(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, -1,size=(500,500),title=u'乐思购红草导入软件')
        self.Center()
        self.pal=wx.Panel(self,-1)
        self.pal.SetBackgroundColour('white')
        btnBox = wx.BoxSizer(wx.HORIZONTAL)
        self.preBtn = wx.Button(self.pal, -1, u"入库单据导入")
        self.skuBtn = wx.Button(self.pal, -1, u"更新基础数据")
        self.ExcelBtn = wx.Button(self.pal, -1, u"转换表格")
        self.orderBtn = wx.Button(self.pal, -1, u"出库单据导入")
        btnBox.Add(self.preBtn,0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)
        btnBox.Add(self.skuBtn,0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)
        btnBox.Add(self.ExcelBtn,0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)
        btnBox.Add(self.orderBtn,0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)

        main = wx.BoxSizer(wx.VERTICAL)
        main.Add(btnBox,0,wx.EXPAND,5)
        self.logText = wx.TextCtrl(self.pal,style=wx.TE_MULTILINE|wx.TE_RICH2|wx.HSCROLL)
        main.Add(self.logText, 1, flag=wx.EXPAND, border=5)

        self.pal.SetSizer(main)
        self.logText.SetValue(u'程序初始化完毕.\n')
        self.Bind(wx.EVT_BUTTON,lambda evt,mark='pre': self.OnBtn(evt,mark),self.preBtn)
        self.Bind(wx.EVT_BUTTON,lambda evt,mark='sku': self.OnBtn(evt,mark),self.skuBtn)
        self.Bind(wx.EVT_BUTTON,lambda evt,mark='excel': self.OnBtn(evt,mark),self.ExcelBtn)

         
    def OnBtn(self,evt,text=''):
        self.logText.SetValue(self.logText.GetValue()+u'选择文件.\n')
        dlg = wx.FileDialog(self.pal, message=u"选择文件",wildcard=wildcard,style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            self.logText.SetValue(self.logText.GetValue()+u'打开文件:'+dlg.GetPath()+".\n") 
            try:
            	if text=='pre':
            		pre_thread(self,dlg.GetPath()).start()
            	elif text=='sku':
            		sku_thread(self,dlg.GetPath()).start()
            	elif text=='excel':
            		excel_thread(self,dlg.GetPath()).start()
            except Exception, e:
                self.logText.SetValue(self.logText.GetValue()+u'处理表单错误:%s'%e+".\n") 
                wx.MessageBox(u'处理表单错误:%s\n'%e,u'提示',wx.ICON_ERROR)
        else:
            self.logText.SetValue(self.logText.GetValue()+u'打开文件文件失败.\n') 
            
        dlg.Destroy()
        self.logText.SetFocus()

    def SetLog(self,msg):
        self.logText.SetValue(self.logText.GetValue()+msg) 





if __name__ == '__main__':
    try:
        app = wx.PySimpleApp()
        # from SQLet import Connect
        MainPage().Show()
        app.MainLoop()
    except Exception, e:
        wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR)

