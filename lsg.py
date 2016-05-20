#coding=utf-8


#!/bin/env python
import wx

class MainPage(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, -1,size=(500,500),title=u'乐思购红草导入软件')
        self.Center()
        self.pal=wx.Panel(self,-1)
        self.pal.SetBackgroundColour('white')
        btnBox = wx.BoxSizer(wx.HORIZONTAL)
        btnBox.Add(wx.Button(self.pal, -1, u"入库单据导入"),0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)
        btnBox.Add(wx.Button(self.pal, -1, u"更新基础数据"),0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)
        btnBox.Add(wx.Button(self.pal, -1, u"转换表格"),0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)
        btnBox.Add(wx.Button(self.pal, -1, u"出库单据导入"),0,wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL,5)

        main = wx.BoxSizer(wx.VERTICAL)
        main.Add(btnBox,0,wx.EXPAND,5)
        logText = wx.TextCtrl(self.pal,style=wx.TE_MULTILINE|wx.TE_RICH2|wx.HSCROLL)
        main.Add(logText, 1, flag=wx.EXPAND, border=5)

        self.pal.SetSizer(main)
        logText.SetValue(u'程序初始化完毕.')


if __name__ == '__main__':
    try:
        app = wx.PySimpleApp()
        # from SQLet import Connect
        MainPage().Show()
        app.MainLoop()
    except Exception, e:
        wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR)

