#coding=utf-8
import xml
print(xml.__file__)
from xml.dom.minidom import Document

doc = Document()  #创建DOM文档对象

dcsmergedata = doc.createElement('dcsmergedata') #创建根元素
dcsmergedata.setAttribute('xmlns:xsi',"http://www.w3.org/2001/XMLSchema-instance")#设置命名空间
dcsmergedata.setAttribute('xsi:noNamespaceSchemaLocation','../lib/interface_pre_advice_header.xsd')#引用本地XML Schema
doc.appendChild(dcsmergedata)
############book:Python处理XML之Minidom################
dataheaders = doc.createElement('dataheaders')
dataheader = doc.createElement('dataheader')
dataheader.setAttribute('transaction','add')

client_id = doc.createElement('client_id')
client_id_t = doc.createTextNode('28')
client_id.appendChild(client_id_t)
dataheader.appendChild(client_id)

notes = doc.createElement('notes')
dataheader.appendChild(notes)
notes_t = doc.createTextNode('5-17销售退货：xsthd20160517001')
notes.appendChild(notes_t)

owner_id = doc.createElement('owner_id')
dataheader.appendChild(owner_id)
owner_id_t = doc.createTextNode('LSG')
owner_id.appendChild(owner_id_t)

pre_advice_id = doc.createElement('pre_advice_id')
dataheader.appendChild(pre_advice_id)
pre_advice_id_t = doc.createTextNode('LSGR20160518003')
pre_advice_id.appendChild(pre_advice_id_t)

pre_advice_type = doc.createElement('pre_advice_type')
dataheader.appendChild(pre_advice_type)
pre_advice_type_t = doc.createTextNode('入库单')
pre_advice_type.appendChild(pre_advice_type_t)

site_id = doc.createElement('site_id')
dataheader.appendChild(site_id)
site_id_t = doc.createTextNode('CPLAJ')
site_id.appendChild(site_id_t)

status = doc.createElement('status')
dataheader.appendChild(status)
status_t = doc.createTextNode('Released')
status.appendChild(status_t)

supplier_id = doc.createElement('supplier_id')
dataheader.appendChild(supplier_id)

datalines = doc.createElement('datalines')
dataheader.appendChild(datalines)
dataline = doc.createElement('dataline')
datalines.appendChild(dataline)
dataline.setAttribute('transaction','add')

line_client_id = doc.createElement('client_id')
dataline.appendChild(line_client_id)
line_client_id_t = doc.createTextNode('LSG')
line_client_id.appendChild(line_client_id_t)

line_condition_id = doc.createElement('condition_id')
dataline.appendChild(line_condition_id)
line_condition_id_t = doc.createTextNode('大库残损')
line_condition_id.appendChild(line_condition_id_t)

line_line_id = doc.createElement('line_id')
dataline.appendChild(line_line_id)
line_line_id_t = doc.createTextNode('1')
line_line_id.appendChild(line_line_id_t)

line_notes_id = doc.createElement('notes')
dataline.appendChild(line_notes_id)
line_notes_t = doc.createTextNode('1')
line_notes_id.appendChild(line_notes_t)

line_owner_id = doc.createElement('owner_id')
dataline.appendChild(line_owner_id)
line_owner_id_t = doc.createTextNode('LSG')
line_owner_id.appendChild(line_owner_id_t)

line_pre_advice_id = doc.createElement('pre_advice_id')
dataline.appendChild(line_pre_advice_id)
line_pre_advice_id_t = doc.createTextNode('LSGR20160518003')
line_pre_advice_id.appendChild(line_pre_advice_id_t)

line_qty_due = doc.createElement('qty_due')
dataline.appendChild(line_qty_due)
line_qty_due_t = doc.createTextNode('1')
line_qty_due.appendChild(line_qty_due_t)

line_sku_id = doc.createElement('sku_id')
dataline.appendChild(line_sku_id)
line_sku_id_t = doc.createTextNode('1128061373010')
line_sku_id.appendChild(line_sku_id_t)



dcsmergedata.appendChild(dataheaders)
dataheaders.appendChild(dataheader)




########### 将DOM对象doc写入文件
f = open('bookstore.xml','w')
f.write(doc.toprettyxml(indent = '    '))#对齐
f.close()