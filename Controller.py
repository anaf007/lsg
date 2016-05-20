#coding=utf-8
import threading,sys,xlrd,wx
from xml.dom.minidom import Document
reload(sys)
sys.setdefaultencoding('utf-8')

class pre_thread(threading.Thread):
	def __init__(self,windows,path):
		threading.Thread.__init__(self)
		threading.Event().clear()
		self.win = windows
		self.path = path
	def run(self):
		wx.CallAfter(self.win.SetLog,u'正在读取表格数据.\n')
		table = xlrd.open_workbook(self.path,encoding_override='utf-8').sheets()[0]
		tableData = []
		for r in range(1,table.nrows):
			if table.row_values(r):
				tableData.append(table.row_values(r))
		wx.CallAfter(self.win.SetLog,u'正在转换XML文件.\n')
		try:
			doc = Document()  #创建DOM文档对象
			dcsmergedata = doc.createElement('dcsmergedata') #创建根元素
			dcsmergedata.setAttribute('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")#设置命名空间
			dcsmergedata.setAttribute('xsi:noNamespaceSchemaLocation','../lib/interface_pre_advice_header.xsd')#引用本地XML Schema
			doc.appendChild(dcsmergedata)
			arr = []
			for i,x in enumerate(tableData):
				if x[1] in arr:
					pass
				else:
					arr.appen(x[1]) #新订单号
					
				print x
		except Exception, e:
			raise u'转换XML错误%s\n'%e