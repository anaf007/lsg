#coding=utf-8
import threading,sys,xlrd,wx,os,time,ftplib
from xml.dom.minidom import Document
import SQLet
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
			dataheaders = doc.createElement('dataheaders')
			dcsmergedata.appendChild(dataheaders)

			arr = []
			index = 1

			for i,x in enumerate(tableData):
				
				if not x[1] in arr :
					arr.append(str(x[1])) #新订单号
					index = 1

					dataheader = doc.createElement('dataheader')
					dataheader.setAttribute('transaction','add')
					datalines = doc.createElement('datalines')

					#头 创建节点
					client_id = doc.createElement('client_id')
					notes = doc.createElement('notes')
					owner_id = doc.createElement('owner_id')
					pre_advice_id = doc.createElement('pre_advice_id')
					pre_advice_type = doc.createElement('pre_advice_type')
					site_id = doc.createElement('site_id')
					status = doc.createElement('status')
					supplier_id = doc.createElement('supplier_id')

					#头 节点设置值
					notes_t = doc.createTextNode(str(x[12]))
					notes.appendChild(notes_t)
					owner_id_t = doc.createTextNode('LSG')
					owner_id.appendChild(owner_id_t)
					pre_advice_id_t = doc.createTextNode(str(x[1]))
					pre_advice_id.appendChild(pre_advice_id_t)
					pre_advice_type_t = doc.createTextNode(u'入库单')
					pre_advice_type.appendChild(pre_advice_type_t)
					site_id_t = doc.createTextNode('CPLAJ')
					site_id.appendChild(site_id_t)
					status_t = doc.createTextNode('Released')
					status.appendChild(status_t)
				
					dataheaders.appendChild(dataheader)

					dataheader.appendChild(client_id)
					dataheader.appendChild(notes)
					dataheader.appendChild(owner_id)
					dataheader.appendChild(pre_advice_id)
					dataheader.appendChild(pre_advice_type)
					dataheader.appendChild(site_id)
					dataheader.appendChild(status)
					dataheader.appendChild(supplier_id)
					dataheader.appendChild(datalines)

				#行
				dataline = doc.createElement('dataline')
				dataline.setAttribute('transaction','add')
				line_client_id = doc.createElement('client_id')
				line_condition_id = doc.createElement('condition_id')
				line_line_id = doc.createElement('line_id')
				line_notes_id = doc.createElement('notes')
				line_owner_id = doc.createElement('owner_id')
				line_sku_id = doc.createElement('sku_id')
				line_pre_advice_id = doc.createElement('pre_advice_id')
				line_qty_due = doc.createElement('qty_due')

				dataline.appendChild(line_client_id)
				dataline.appendChild(line_condition_id)
				dataline.appendChild(line_line_id)
				dataline.appendChild(line_notes_id)
				dataline.appendChild(line_owner_id)
				dataline.appendChild(line_pre_advice_id)
				dataline.appendChild(line_qty_due)
				dataline.appendChild(line_sku_id)

				datalines.appendChild(dataline)

				line_client_id_t = doc.createTextNode('LSG')
				line_client_id.appendChild(line_client_id_t)
				line_condition_id_t = doc.createTextNode(str(x[2]))
				line_condition_id.appendChild(line_condition_id_t)
				line_line_id_t = doc.createTextNode(str(index))
				line_line_id.appendChild(line_line_id_t)
				line_notes_t = doc.createTextNode(str(x[12]))
				line_notes_id.appendChild(line_notes_t)
				line_owner_id_t = doc.createTextNode('LSG')
				line_owner_id.appendChild(line_owner_id_t)
				line_sku_id_t = doc.createTextNode(str(x[4]))
				line_sku_id.appendChild(line_sku_id_t)
				line_pre_advice_id_t = doc.createTextNode(str(x[1]))
				line_pre_advice_id.appendChild(line_pre_advice_id_t)
				line_qty_due_t = doc.createTextNode(str(int(float(str(x[10])))))
				line_qty_due.appendChild(line_qty_due_t)
				index = index+1
				

			filename = str(time.strftime(u"%Y%m%d%H%M%S",time.localtime()))
			path = sys.path[0]+'\\xml\\'+filename+'_LSG_cl_interface_pre.xml'
			f = open(path,'w')
			f.write(doc.toprettyxml(indent = '    '))
			f.close()
			wx.CallAfter(self.win.SetLog,u'转换完成,文件保存在:\n%s.\n'%path)

			#上传FTP
			ftp_server = '192.168.2.199'
			try:
				wx.CallAfter(self.win.SetLog,u'正在连接服务器%s...\n'%ftp_server)
				ftp = ftplib.FTP(ftp_server)
				ftp.login('tstdba','tstdba')
				ftp.storlines('STOR comms/intray/'+filename+'_LSG_cl_interface_pre.xml',open(path))
				wx.CallAfter(self.win.SetLog,u'XML文件上传完成.\n')
			except Exception, e:
				wx.CallAfter(self.win.SetLog,u'连接服务器%s失败...\n'%ftp_server)

		except Exception, e:
			wx.CallAfter(self.win.SetLog,u'转换XML错误,错误原因:%s\n'%e)
			 



class sku_thread(threading.Thread):
	def __init__(self,windows,path):
		threading.Thread.__init__(self)
		threading.Event().clear()
		self.win = windows
		self.path = path
	def run(self):
		wx.CallAfter(self.win.SetLog,u'正在检查表格数据.\n')
		table = xlrd.open_workbook(self.path,encoding_override='utf-8').sheets()[0]
		print table.col(0)[0].value.strip()
		try:
			if table.col(0)[0].value.strip() != u'仓库编码':
				message = u"第一行名称必须叫‘仓库编码’，请返回修改"
			if table.col(1)[0].value.strip() != u'仓库名称':
				message = u"第二行名称必须叫‘仓库名称’，请返回修改"
			if table.col(3)[0].value.strip() != u'库存类别':
				message = u"第三行名称必须叫‘库存类别’，请返回修改"
			if table.col(4)[0].value.strip() != u'储位':
				message = u"第四行名称必须叫‘储位’，请返回修改"
			if table.col(5)[0].value.strip() != u'商品编码':
				message = u"第五行名称必须叫‘商品编码’，请返回修改"
			if table.col(7)[0].value.strip() != u'简称':
				message = u"第六行名称必须叫‘简称’，请返回修改"
			if table.col(8)[0].value.strip() != u'颜色编码':
				message = u"第六行名称必须叫‘颜色编码’，请返回修改"
			if table.col(9)[0].value.strip() != u'颜色':
				message = u"第六行名称必须叫‘颜色’，请返回修改"
			if table.col(11)[0].value.strip() != u'款式编码':
				message = u"第六行名称必须叫‘款式编码’，请返回修改"
			if table.col(12)[0].value.strip() != u'款式':
				message = u"第六行名称必须叫‘款式’，请返回修改"
		except Exception,f:
			wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
		insert_data_list_marge = []
		
		try:
			for i,x in enumerate(insert_data):
				insert_data_list = {'ckbh':x[0],'ckmc':x[1],'kclb':x[3],'cw':x[4],\
				'spbm':x[5],'gdx':x[6],'jc':x[7],'ysbm':x[8],'ys':x[9],'ksbm':x[11],'ks':x[12],\
				'pc':'','yxrq':'','kcs':''}
				nsert_data_list_marge.append(insert_data_list)
			if insert_data_list_marge:
				# Connect().delete('sku','1=1')
				Connect().insert_many(insert_data_list_marge,'sku')
		except Exception, e:
			wx.CallAfter(self.win.SetLog,u'连接服务器%s失败.\n'%str('192.168.2.102'))


