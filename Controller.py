#coding=utf-8
import threading,sys,xlrd,wx,os,time,ftplib,xlwt
from xml.dom.minidom import Document
from SQLet import *
from collections import defaultdict
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
			dcsmergedata.setAttribute('xsi:noNamespaceSchemaLocation','../lib/interface_pre_advice_header.xsd')#引用本地XML Schema
			dcsmergedata.setAttribute('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")#设置命名空间
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
					client_id_t = doc.createTextNode('LSG')
					client_id.appendChild(client_id_t)
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
				

			filename = str(time.strftime(u"%Y_%m_%d_%H%M%S",time.localtime()))
			path = sys.path[0]+'\\xml\\'+filename+'_LSG_cl_interface_pre.xml'
			f = open(path,'w')
			f.write(doc.toprettyxml(indent = '    '))
			f.close()
			wx.CallAfter(self.win.SetLog,u'转换完成,文件保存在:\n%s.\n'%path)

			#上传FTP
			ftp_server = '192.168.2.100'
			try:
				wx.CallAfter(self.win.SetLog,u'正在连接服务器%s...\n'%ftp_server)
				ftp = ftplib.FTP(ftp_server)
				ftp.login('dcsdba','dcsabd')
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
		
		try:
			if table.col(0)[0].value.strip() != u'仓库编码':
				raise Exception(u"第一行名称必须叫‘仓库编码’，请返回修改")
			if table.col(1)[0].value.strip() != u'仓库名称':
				raise Exception(u"第二行名称必须叫‘仓库名称’，请返回修改")
			if table.col(3)[0].value.strip() != u'库存类别':
				raise Exception(u"第三行名称必须叫‘库存类别’，请返回修改")
			if table.col(4)[0].value.strip() != u'储位':
				raise Exception(u"第四行名称必须叫‘储位’，请返回修改")
			if table.col(5)[0].value.strip() != u'商品编码':
				raise Exception(u"第五行名称必须叫‘商品编码’，请返回修改")
			if table.col(7)[0].value.strip() != u'简称':
				raise Exception(u"第六行名称必须叫‘简称’，请返回修改")
			if table.col(8)[0].value.strip() != u'颜色编码':
				raise Exception(u"第六行名称必须叫‘颜色编码’，请返回修改")
			if table.col(9)[0].value.strip() != u'颜色':
				raise Exception(u"第六行名称必须叫‘颜色’，请返回修改")
			if table.col(11)[0].value.strip() != u'款式编码':
				raise Exception(u"第六行名称必须叫‘款式编码’，请返回修改")
			if table.col(12)[0].value.strip() != u'款式':
				raise Exception(u"第六行名称必须叫‘款式’，请返回修改")
		except Exception,f:
			wx.CallAfter(self.win.SetLog,u'表单列数读取错误，错误原因:%s\n'%f);return
		
		table_data_list= []
		for rownum in range(1,table.nrows):
			if table.row_values(rownum):
				table_data_list.append(table.row_values(rownum))

		insert_data_list_marge = []
		try:
			for i,x in enumerate(table_data_list):
				insert_data_list = {'ckbh':x[0],'ckmc':x[1],'kclb':x[3],'cw':x[4],\
				'spbm':x[5],'gdx':x[6],'jc':x[7],'ysbm':x[8],'ys':x[9],'ksbm':x[11],'ks':x[12],\
				'pc':'','yxrq':'','kcs':''}
				insert_data_list_marge.append(insert_data_list)
			wx.CallAfter(self.win.SetLog,u'正在保存数据，请勿关闭窗口.\n')
			if insert_data_list_marge:
				Connect().delete('lsg_sku','1=1')
				Connect().insert_many(insert_data_list_marge,'lsg_sku')
				wx.CallAfter(self.win.SetLog,u'更新基础数据完成.\n')
		except Exception, e:
			wx.CallAfter(self.win.SetLog,u'连接服务器%s失败.错误原因:%s\n'%(str('192.168.2.102'),str(e)))


class excel_thread(threading.Thread):
	def __init__(self,windows,path):
		threading.Thread.__init__(self)
		threading.Event().clear()
		self.win = windows
		self.path = path
	def run(self):
		wx.CallAfter(self.win.SetLog,u'正在检查表格数据.\n')
		table = xlrd.open_workbook(self.path,encoding_override='utf-8').sheets()[0]
		table_data_list= []
		for rownum in range(1,table.nrows):
			if table.row_values(rownum):
				table_data_list.append(table.row_values(rownum))

		#得到库存字典
		goods_list = []
		try:
			goods_tup = Connect().select('*','lsg_sku')
		except Exception, e:
			wx.CallAfter(self.win.SetLog,u'打开数据库失败%s.\n'%e);return
		
		ok_list = []
		no_list = []
		con_list = []
		order_lsg = 'lsgc'+str(time.strftime(u"%Y%m%d%H%M%S",time.localtime())) #单号
		wx.CallAfter(self.win.SetLog,u'正在转换表格数据.\n')
		try:
			for i in table_data_list:
				for goods in goods_tup:
					try:
						i[2] = str(int(float(str(i[2]))))
					except Exception, e:
						i[2] = str(i[2])
					try:
						i[3] = str(int(float(str(i[3]))))
					except Exception, e:
						i[3] = str(i[3])
					
					if str(int(float(str(i[0])))) == str(goods[5]) and str(i[2])==str(goods[9]) and str(i[3])==str(goods[11]):
						
						chbm = str(int(float(str(i[0]))))
						chmc = str(int(float(str(i[0]))))+"'"+str(i[1])+"'"
						if str(i[2])==str(goods[9]):
							chbm = chbm+str(goods[8])
							chmc = chmc+i[2]
						if str(i[3])==str(goods[11]):
							chbm = chbm+str(goods[10])
							chmc = chmc+"'"+str(i[3])
						ok_list.append([order_lsg,str(i[7]),'',chbm,chmc,str(int(float(str(i[5])))),str(i[8])])
						break
					else:
						continue
				else:
					no_list.append(i)
		except Exception, e:
			wx.CallAfter(self.win.SetLog,u'转换表格错误%s.\n'%e);return

		
		wx.CallAfter(self.win.SetLog,u'正在生成表格.\n')
		file = xlwt.Workbook(encoding='utf8')
		ok_title =['来源单据号','仓库','调出仓库','存货编码','存货名称','出库数量','备注']
		excel_ok = file.add_sheet(u'已转换行')
		for i,x in enumerate(ok_title):
			excel_ok.write(0,i,x)
		for i,x in enumerate(ok_list):
			for i_index,x_v in enumerate(x):
				excel_ok.write(i+1,i_index,x_v)

		excel_file = file.add_sheet(u'未转换行')
		title = [u'商品编号',u'商品名称',u'颜色',u'型号',u'赠品',u'捡货数量',u'确定√',u'条件',u'备件']
		for i,x in enumerate(title):
			excel_file.write(0,i,x)
		for i,x in enumerate(no_list):
			for i_index,x_v in enumerate(x):
				excel_file.write(i+1,i_index,x_v)
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			file.save(select_dialog.GetPath()+u"/LSGC"+time.strftime(u"%Y%m%d%H%M%S",time.localtime())+".xls")
		select_dialog.Destroy()
		wx.CallAfter(self.win.SetLog,u'转换操作完成.\n')

	

class order_thread(threading.Thread):
	def __init__(self,windows,path):
		threading.Thread.__init__(self)
		threading.Event().clear()
		self.win = windows
		self.path = path
	def run(self):
		wx.CallAfter(self.win.SetLog,u'正在检查表格数据.\n')
		table = xlrd.open_workbook(self.path,encoding_override='utf-8').sheets()[0]
		tableData= []
		for rownum in range(1,table.nrows):
			if table.row_values(rownum):
				tableData.append(table.row_values(rownum))

		wx.CallAfter(self.win.SetLog,u'正在转换XML文件.\n')
		try:
			doc = Document()  #创建DOM文档对象
			dcsmergedata = doc.createElement('dcsmergedata') #创建根元素
			dcsmergedata.setAttribute('xsi:noNamespaceSchemaLocation','../lib/interface_order_header.xsd')
			dcsmergedata.setAttribute('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")#设置命名空间
			doc.appendChild(dcsmergedata)
			dataheaders = doc.createElement('dataheaders')
			dcsmergedata.appendChild(dataheaders)

			arr = []
			index = 1
			for i,x in enumerate(tableData):
				
				if not x[0] in arr :
					arr.append(str(x[0])) #新订单号
					index = 1

					dataheader = doc.createElement('dataheader')
					dataheader.setAttribute('transaction','add')
					datalines = doc.createElement('datalines')

					#头 创建节点
					client_id = doc.createElement('client_id')
					customer_id = doc.createElement('customer_id')
					from_site_id = doc.createElement('from_site_id')
					instructions = doc.createElement('instructions')
					order_id = doc.createElement('order_id')
					order_type = doc.createElement('order_type')
					owner_id = doc.createElement('owner_id')
					ship_dock = doc.createElement('ship_dock')
					status = doc.createElement('status')



					#头 节点设置值
					client_id_t = doc.createTextNode('LSG')
					client_id.appendChild(client_id_t)
					customer_id_t = doc.createTextNode('LSG01')
					customer_id.appendChild(customer_id_t)
					from_site_id_t = doc.createTextNode('CPLAJ')
					from_site_id.appendChild(from_site_id_t)
					instructions_t = doc.createTextNode(x[6])
					instructions.appendChild(instructions_t)
					order_id_t = doc.createTextNode(x[0])
					order_id.appendChild(order_id_t)
					order_type_t = doc.createTextNode(u'出库单')
					order_type.appendChild(order_type_t)
					owner_id_t = doc.createTextNode('LSG')
					owner_id.appendChild(owner_id_t)
					ship_dock_t = doc.createTextNode('LFH01')
					ship_dock.appendChild(ship_dock_t)
					status_t = doc.createTextNode('Released')
					status.appendChild(status_t)

					dataheaders.appendChild(dataheader)

					dataheader.appendChild(client_id)
					dataheader.appendChild(customer_id)
					dataheader.appendChild(from_site_id)
					dataheader.appendChild(instructions)
					dataheader.appendChild(order_id)
					dataheader.appendChild(order_type)
					dataheader.appendChild(owner_id)
					dataheader.appendChild(ship_dock)
					dataheader.appendChild(status)

					dataheader.appendChild(datalines)

				#行
				dataline = doc.createElement('dataline')
				dataline.setAttribute('transaction','add')
				line_client_id = doc.createElement('client_id')
				line_condition_id = doc.createElement('condition_id')
				line_line_id = doc.createElement('line_id')
				line_notes_id = doc.createElement('notes')
				line_order_id = doc.createElement('order_id')
				line_owner_id = doc.createElement('owner_id')
				line_qty_ordered = doc.createElement('qty_ordered')
				line_sku_id = doc.createElement('sku_id')


				dataline.appendChild(line_client_id)
				dataline.appendChild(line_condition_id)
				dataline.appendChild(line_line_id)
				dataline.appendChild(line_notes_id)
				dataline.appendChild(line_order_id)
				dataline.appendChild(line_owner_id)
				dataline.appendChild(line_qty_ordered)
				dataline.appendChild(line_sku_id)

				datalines.appendChild(dataline)

				line_client_id_t = doc.createTextNode('LSG')
				line_client_id.appendChild(line_client_id_t)
				line_condition_id_t = doc.createTextNode(str(x[1]))
				line_condition_id.appendChild(line_condition_id_t)
				line_line_id_t = doc.createTextNode(str(index))
				line_line_id.appendChild(line_line_id_t)
				line_notes_t = doc.createTextNode(str(x[6]))
				line_notes_id.appendChild(line_notes_t)
				line_order_id_t = doc.createTextNode(str(x[0]))
				line_order_id.appendChild(line_order_id_t)
				line_owner_id_t = doc.createTextNode('LSG')
				line_owner_id.appendChild(line_owner_id_t)
				line_qty_ordered_t = doc.createTextNode(str(int(float(str(x[5])))))
				line_qty_ordered.appendChild(line_qty_ordered_t)
				line_sku_id_t = doc.createTextNode(str(int(float(str(x[3])))))
				line_sku_id.appendChild(line_sku_id_t)
				
				index = index+1
			filename = str(time.strftime(u"%Y_%m_%d_%H%M%S",time.localtime()))
			path = sys.path[0]+'\\xml\\'+filename+'_interface_order.xml'
			f = open(path,'w')
			f.write(doc.toprettyxml(indent = '    '))
			f.close()
			wx.CallAfter(self.win.SetLog,u'转换完成,文件保存在:\n%s.\n'%path)

			#上传FTP
			ftp_server = '192.168.2.100'
			try:
				wx.CallAfter(self.win.SetLog,u'正在连接服务器%s...\n'%ftp_server)
				ftp = ftplib.FTP(ftp_server)
				ftp.login('dcsdba','dcsabd')
				ftp.storlines('STOR comms/intray/'+filename+'_interface_order.xml',open(path))
				wx.CallAfter(self.win.SetLog,u'XML文件上传完成.\n')
			except Exception, e:
				wx.CallAfter(self.win.SetLog,u'连接服务器%s失败...\n'%ftp_server)

				
		except Exception, e:
			wx.CallAfter(self.win.SetLog,u'转换XML错误,错误原因:%s\n'%e)