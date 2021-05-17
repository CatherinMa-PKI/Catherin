# 修改功能
### 方法一：只把修改的那一条回写进数据库。
适用范围:对于原表需要修改的行数不多的情况下。例如字段抓取结果大部分正确，只有少部分的需要修改的情况。
修改操作流程：在前台展示需要修改的原表格 → 用户选择需要修改的行 → 在相应的文本属性中填入或者选择修改的值 → 填入修改人名 → 点击‘确认修改’按钮调动脚本进行修改 → 前台展示修改后的数据＋生成修改log
修改设计流程：原表转为竖表，类为所有需要修改的列，值为对应类的值 → 在数据库中建好修改表格纵表,和原表的结构保持一样。写好修改和写入修改log的存储过程 → 写修改按钮的ironPython脚本 → 把数据库中的修改数据替换掉原表数据生成修改后数据表，该表需要链接到源 → 前台展示修改后数据表。
Tips：
1.最好原表为竖表进行修改，这样比较好留下修改的log，而且ironPython不用定义很多列，不容易出错。但是如果不想进行过多Pivot和Unpivot的操作的话，可以只把需要修改的列拿出来，用横表的方式写进脚本。
2.修改表替换到原表的方法:建议用修改表和原表匹配好之后写一个新的列‘修改’值为‘是’。然后筛选掉‘是’的行，再把修改表添加行进去。
代码演示1：
例子说明：原表为横表，需要确认修改的字段总共7个：肿瘤类型，良恶性，T，N，M，癌症分期，复发风险
```
from Spotfire.Dxp.Data import AddRowsSettings
import System
from System import Environment, Threading, DateTime
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from Spotfire.Dxp.Data import DataType, DataTableSaveSettings
from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings
from Spotfire.Dxp.Data import*
import time

import clr, datetime
clr.AddReference('System.Data')
clr.AddReference("System.Windows.Forms")
from System.Data import SqlClient
from System.Windows.Forms import MessageBox, MessageBoxButtons
from System.Windows.Forms import DialogResult

from System.Collections.Generic import List, Dictionary 
from Spotfire.Dxp.Data import DataTable, IndexSet, RowSelection
from System.Collections import ArrayList
from Spotfire.Dxp.Framework.ApplicationModel import NotificationService

def ClearInput():
	Document.Properties["肿瘤类型"] = ""
	Document.Properties["癌症分期"] = ""
	Document.Properties["复发风险"] = ""
	Document.Properties["T"] = ""
	Document.Properties["N"] = ""
	Document.Properties["M"] = ""
	Document.Properties["良恶性"] = ""
def CleanSpace(item):
	item=item.replace(' ','')
	return item 	
#获得修改人姓名
currentUser=Document.Properties['修改人']
print currentUser

dt = Document.Data.Tables["甲状腺相关病例_修改"]

Cursor1 = DataValueCursor.CreateFormatted(dt.Columns["EMPIID"])
Cursor2 = DataValueCursor.CreateFormatted(dt.Columns["患者姓名"])
Cursor3 = DataValueCursor.CreateFormatted(dt.Columns["就诊流水号"])
Cursor4 = DataValueCursor.CreateFormatted(dt.Columns["肿瘤类型"])
Cursor5 = DataValueCursor.CreateFormatted(dt.Columns["良恶性"])
Cursor6 = DataValueCursor.CreateFormatted(dt.Columns["T"])
Cursor7 = DataValueCursor.CreateFormatted(dt.Columns["N"])
Cursor8 = DataValueCursor.CreateFormatted(dt.Columns["M"])
Cursor9 = DataValueCursor.CreateFormatted(dt.Columns["癌症分期"])
Cursor10 = DataValueCursor.CreateFormatted(dt.Columns["复发风险"])

#获得标记的行
markings = Document.ActiveMarkingSelectionReference.GetSelection(dt)

sqlStr1=' '
sqlStr2=' '
comment1=' '
comment2=' '
comment3=' '
comment4=' '
comment5=' '
comment6=' '
comment7=' '

conn = SqlClient.SqlConnection("Server=172.30.100.170;Database=PKI_ZBK_DataMapping;UID=sa;PWD=Shdlrmyy@170#")
conn.Open()
print "Connect DB"

for row in dt.GetRows(markings.AsIndexSet(),Cursor1,Cursor2,Cursor3,Cursor4,Cursor5,Cursor6,Cursor7,Cursor8,Cursor9,Cursor10):
	column1 = Cursor1.CurrentValue
	column2 = Cursor2.CurrentValue
	column3 = Cursor3.CurrentValue
	column4 = Cursor4.CurrentValue
	column5 = Cursor5.CurrentValue
	column6 = Cursor6.CurrentValue
	column7 = Cursor7.CurrentValue
	column8 = Cursor8.CurrentValue
	column9 = Cursor9.CurrentValue
	column10 = Cursor10.CurrentValue
	if CleanSpace(Document.Properties["肿瘤类型"]) != "" and CleanSpace(Document.Properties["肿瘤类型"]) != column4:
		comment1='肿瘤类型:'+column4+'修改为'+CleanSpace(Document.Properties["肿瘤类型"])
		column4=CleanSpace(Document.Properties["肿瘤类型"])

	if Document.Properties["良恶性"] != "" and Document.Properties["良恶性"] != column5:
		comment2='良恶性:'+column5+'修改为'+Document.Properties["良恶性"]
		column5=Document.Properties["良恶性"] 

	if Document.Properties["T"] != "" and Document.Properties["T"] != column6:
		comment3='T:'+column6+'修改为'+Document.Properties["T"]
		column6=Document.Properties["T"] 

	if Document.Properties["N"] != "" and Document.Properties["N"] != column7:
		comment4='N:'+column7+'修改为'+Document.Properties["N"]
		column7=Document.Properties["N"] 

	if Document.Properties["M"] != "" and Document.Properties["M"] != column8:
		comment5='M:'+column8+'修改为'+Document.Properties["M"]
		column8=Document.Properties["M"] 

	if Document.Properties["癌症分期"] != "" and Document.Properties["癌症分期"] != column9:
		comment6='癌症分期:'+column9+'修改为'+Document.Properties["癌症分期"]
		column9=Document.Properties["癌症分期"]
 
	if Document.Properties["复发风险"] != "" and Document.Properties["复发风险"] != column10:
		comment7='复发风险:'+column10+'修改为'+Document.Properties["复发风险"]
		column10=Document.Properties["复发风险"]

	if column4 == "(空)" or column4 == "(Empty)":
		column4 = "" 
	if column5 == "(空)" or column5 == "(Empty)":
		column5 = "" 
	if column6 == "(空)" or column6 == "(Empty)":
		column6 = "" 
	if column7 == "(空)" or column7 == "(Empty)":
		column7 = "" 
	if column8 == "(空)" or column8 == "(Empty)":
		column8 = "" 
	if column9 == "(空)" or column9 == "(Empty)":
		column9 = "" 
	if column10 == "(空)" or column10 == "(Empty)":
		column10 = "" 
	sqlStr1 = "EXEC PKI_ZBK_DataMapping_InsertEditAZFQ '" + column1 + "','" + column2 + "','" + column3 + "','" + column4 + "','" + column5 + "','"+ column6 + "','"+ column7 + "','"+ column8 + "','"+ column9 + "','"+ column10 + "'"
	print sqlStr1
	comment=comment1+" "+comment2+' '+comment3+' '+comment4+' '+comment5+' '+comment6+' '+comment7
	comment=comment.strip()
	sqlStr2 = "EXEC PKI_ZBK_DataMapping_EditedAZFXLog '" + column1 + "','" + column2 + "','" + column3 + "','" + comment + "','" + currentUser+  "'"
	print sqlStr2

	cmd = SqlClient.SqlCommand(sqlStr1, conn)
	exe = cmd.ExecuteReader()
	exe.Close()
	print('修改表格回写成功')

	cmd = SqlClient.SqlCommand(sqlStr2, conn)
	exe = cmd.ExecuteReader()
	exe.Close()
	print('Log回写成功')

conn.Close()

dt1 = Document.Data.Tables["甲状腺修改Log"]
dt.Refresh()
dt1.Refresh()

print "修改完成！"

Document.Properties["修改结果"] = "修改成功! ^-^"
Document.Properties["修改内容"] = comment
ClearInput()
comment=''
```
代码演示2：修改原表为纵表，需要确认修改的字段总共7个：肿瘤类型，良恶性，T，N，M，癌症分期，复发风险，即有7个类。
```
from Spotfire.Dxp.Data import AddRowsSettings
import System
from System import Environment, Threading, DateTime
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from Spotfire.Dxp.Data import DataType, DataTableSaveSettings
from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings
from Spotfire.Dxp.Data import*
import time

import clr, datetime
clr.AddReference('System.Data')
clr.AddReference("System.Windows.Forms")
from System.Data import SqlClient
from System.Windows.Forms import MessageBox, MessageBoxButtons
from System.Windows.Forms import DialogResult

from System.Collections.Generic import List, Dictionary 
from Spotfire.Dxp.Data import DataTable, IndexSet, RowSelection
from System.Collections import ArrayList
from Spotfire.Dxp.Framework.ApplicationModel import NotificationService

def ClearInput():
	Document.Properties["肿瘤类型"] = ""
	Document.Properties["癌症分期"] = ""
	Document.Properties["复发风险"] = ""
	Document.Properties["T"] = ""
	Document.Properties["N"] = ""
	Document.Properties["M"] = ""
	Document.Properties["良恶性"] = ""
def CleanSpace(item):
	item=item.replace(' ','')
	return item 	
dataTable = Document.Data.Tables["甲状腺相关病例_修改"]

rowCount - dataTable.RowCount
rowsToInclude = IndexSet(rowCount,True)
print rowCount

cursor1 = DataValueCursor.CreateFormatted(dataTable.Columns["EMPIID"])
cursor2 = DataValueCursor.CreateFormatted(dataTable.Columns["患者姓名"])
cursor3 = DataValueCursor.CreateFormatted(dataTable.Columns["就诊流水号"])
cursor4 = DataValueCursor.CreateFormatted(dataTable.Columns["类"])
cursor5 = DataValueCursor.CreateFormatted(dataTable.Columns["值"])

#获取当前使用登录的UserName
currentUser = Threading.Thread.CurrentPrincipal.Identity.Name
print currentUser

sqlStr1 = ""
sqlStr2 = ""
Document.Properties["修改结果"] = ""


conn = SqlClient.SqlConnection("Server=172.30.100.170;Database=PKI_ZBK_DataMapping;UID=sa;PWD=Shdlrmyy@170#")
conn.Open()
print "Connect DB"

for row in dataTable.GetRows(rowsToInclude,cursor1,cursor2,cursor3,cursor4,cursor5):
	empi = cursor1.CurrentValue
	patientName = cursor2.CurrentValue
	serialNo = cursor3.CurrentValue
	itemName = cursor4.CurrentValue
	origValue = cursor5.CurrentValue
	try:
		newValue = Document.Properties[itemName]
	except KeyError as e:
		print itemName+" 不允许修改"
		newValue = ""

	newValue = newValue.replace(" ", "")
	if newValue != "":
		if newValue != origValue:
			sqlStr1 = "EXEC PKI_ZBK_DataMapping_ModifyAZFQ '" + empi+ "','"+ patientName + "','"+ serialNo+ "','" + itemName + "','" + newValue + "'"
			print sqlStr1
			sqlStr2 = "EXEC PKI_ZBK_DataMapping_InsertModifyAZFQLog '" + empi + "','" + patientName + "','" + serialNo + "','" + itemName + "','" + origValue + "','" + newValue + "','" + currentUser + "'"
			print sqlStr2
			
			comment=itemName+':'+origValue+'修改为'+newValue+',修改人为'+currentUser

			cmd = SqlClient.SqlCommand(sqlStr1, conn)
			exe = cmd.ExecuteReader()
			exe.Close()
			print "修改表格回写成功"

			cmd = SqlClient.SqlCommand(sqlStr2, conn)
			exe = cmd.ExecuteReader()
			exe.Close()
			print "Log表格回写成功"
conn.Close()
dt1 = Document.Data.Tables["甲状腺修改Log"]
dt.Refresh()
dt1.Refresh()

print "修改完成！"

Document.Properties["修改结果"] = "修改成功! ^-^"
Document.Properties["修改内容"] = comment
ClearInput()
```

### 方法二：把整张表都回写进数据库，之后找到那一条进行sql修改。
适用范围:对于原表不大，行数较少，并且需要修改确认较多条数，需要修改确认的字段比较多的情况。
修改操作流程：在前台展示需要修改的原表格 → 用户选择需要修改的行 → 在相应的文本属性中填入或者选择修改的值 → 点击‘确认修改’按钮调动脚本进行修改 → 前台展示修改后的数据＋生成修改log
修改设计流程：原表转为竖表，类为所有需要修改的列，值为对应类的值 → 将原表全量回写进数据库 → 写好修改和写入修改log的存储过程 → 写修改按钮的ironPython脚本 → 拽入修改过的原表 → 前台展示修改后数据表。
Tips: 建议分成另外一个Modify模板，修改完的数据再到处给应用模板用。
代码演示：
例子说明：原表为横表，需要确认修改的字段总共26个。修改的脚本和上方法一的脚本一致，下面只展示回写脚本。
```
# 将原表写进数据库
from Spotfire.Dxp.Data import AddRowsSettings
import System
from System import Environment, Threading, DateTime
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from Spotfire.Dxp.Data import DataType, DataTableSaveSettings
from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings
from Spotfire.Dxp.Data import*
import time
import datetime

import clr, datetime
clr.AddReference('System.Data')
clr.AddReference("System.Windows.Forms")
from System.Data import SqlClient
from System.Windows.Forms import MessageBox, MessageBoxButtons
from System.Windows.Forms import DialogResult

from System.Collections.Generic import List, Dictionary 
from Spotfire.Dxp.Data import DataTable, IndexSet, RowSelection
from System.Collections import ArrayList
from Spotfire.Dxp.Framework.ApplicationModel import NotificationService

dataTable = Document.Data.Tables["甲状腺相关病例_修改"]
rowCount = dataTable.RowCount
rowsToInclude = IndexSet(rowCount,True)
print rowCount

cursor1 = DataValueCursor.CreateFormatted(dataTable.Columns["EMPIID"])
cursor2 = DataValueCursor.CreateFormatted(dataTable.Columns["患者姓名"])
cursor3 = DataValueCursor.CreateFormatted(dataTable.Columns["就诊流水号"])
cursor4 = DataValueCursor.CreateFormatted(dataTable.Columns["类"])
cursor5 = DataValueCursor.CreateFormatted(dataTable.Columns["值"])
cursor6 = DataValueCursor.CreateFormatted(dataTable.Columns["是否新记录"])


conn = SqlClient.SqlConnection("Server=172.30.100.170;Database=PKI_ZBK_DataMapping;UID=sa;PWD=Shdlrmyy@170#")
conn.Open()
print "Connect DB"

sqlStr1 = ""

for row in dataTable.GetRows(rowsToInclude,cursor1,cursor2,cursor3,cursor4,cursor5,cursor6):
	empi = cursor1.CurrentValue
	patientName = cursor2.CurrentValue
	serialNo = cursor3.CurrentValue
	category = cursor4.CurrentValue
	value = cursor5.CurrentValue
	NewYN = cursor6.CurrentValue
	if NewYN == '是'：
		sqlStr1 = "EXEC PKI_ZBK_DataMapping_InsertAZFQ '" + empi + "','" + patientName + "','" + serialNo + "','"+ category + "','"+ value + "'"
		cmd = SqlClient.SqlCommand(sqlStr1, conn)
		exe = cmd.ExecuteReader()
		exe.Close()
		print "该条记录写入完成"
conn.Close()

#修改脚本
def ClearInput():
	Document.Properties["肿瘤类型"] = ""
	Document.Properties["癌症分期"] = ""
	Document.Properties["复发风险"] = ""
	Document.Properties["T"] = ""
	Document.Properties["N"] = ""
	Document.Properties["M"] = ""
	Document.Properties["良恶性"] = ""
def CleanSpace(item):
	item=item.replace(' ','')
	return item 	
dataTable = Document.Data.Tables["甲状腺相关病例_修改"]

rowCount - dataTable.RowCount
rowsToInclude = IndexSet(rowCount,True)
print rowCount

cursor1 = DataValueCursor.CreateFormatted(dataTable.Columns["EMPIID"])
cursor2 = DataValueCursor.CreateFormatted(dataTable.Columns["患者姓名"])
cursor3 = DataValueCursor.CreateFormatted(dataTable.Columns["就诊流水号"])
cursor4 = DataValueCursor.CreateFormatted(dataTable.Columns["类"])
cursor5 = DataValueCursor.CreateFormatted(dataTable.Columns["值"])

#获取当前使用登录的UserName
currentUser = Threading.Thread.CurrentPrincipal.Identity.Name
print currentUser

sqlStr1 = ""
sqlStr2 = ""
Document.Properties["修改结果"] = ""


conn = SqlClient.SqlConnection("Server=172.30.100.170;Database=PKI_ZBK_DataMapping;UID=sa;PWD=Shdlrmyy@170#")
conn.Open()
print "Connect DB"

for row in dataTable.GetRows(rowsToInclude,cursor1,cursor2,cursor3,cursor4,cursor5):
	empi = cursor1.CurrentValue
	patientName = cursor2.CurrentValue
	serialNo = cursor3.CurrentValue
	itemName = cursor4.CurrentValue
	origValue = cursor5.CurrentValue
	try:
		newValue = Document.Properties[itemName]
	except KeyError as e:
		print itemName+" 不允许修改"
		newValue = ""

	newValue = newValue.replace(" ", "")
	if newValue != "":
		if newValue != origValue:
			sqlStr1 = "EXEC PKI_ZBK_DataMapping_ModifyAZFQ '" + empi+ "','"+ patientName + "','"+ serialNo+ "','" + itemName + "','" + newValue + "'"
			print sqlStr1
			sqlStr2 = "EXEC PKI_ZBK_DataMapping_InsertModifyAZFQLog '" + empi + "','" + patientName + "','" + serialNo + "','" + itemName + "','" + origValue + "','" + newValue + "','" + currentUser + "'"
			print sqlStr2
			
			comment=itemName+':'+origValue+'修改为'+newValue+',修改人为'+currentUser

			cmd = SqlClient.SqlCommand(sqlStr1, conn)
			exe = cmd.ExecuteReader()
			exe.Close()
			print "修改表格回写成功"

			cmd = SqlClient.SqlCommand(sqlStr2, conn)
			exe = cmd.ExecuteReader()
			exe.Close()
			print "Log表格回写成功"
conn.Close()
# Empty list to hold DataTables
Tbls = List[DataTable]()
Tbls.Add(Document.Data.Tables["修改记录"]) #dt1 a DataTable string parameter
Tbls.Add(Document.Data.Tables["癌症分期修改_纵表"]) #dt1 a DataTable string parameter

# Notification service
notify = Application.GetService[NotificationService]();

# Execute something after tables are loaded
def afterLoad(exception, Document=Document, notify=notify):
	if not exception:
		Document.Properties["修改结果"] = "OK"
	else:
		notify.AddErrorNotification("Error refreshing table(s)","Error details",str(exception))
		Document.Properties["修改结果"] = "Error,pass"
		pass
	Document.Properties["修改结果"] = "修改完成！"

# Refresh table(s)
Document.Data.Tables.RefreshAsync(Tbls, afterLoad)
dataTable.Refresh()
print "修改完成！"
ClearInput()
```
### 方法三：通过写stream的方式写修改后的表格
建议比较小的表格使用这个方法，因为stream是写全张表。首先需要有一张原始表，然后复制原始表为新表dt2用来前台展示和mark，mark的数据修改完之后存在dt3，然后把dt3的数据给替换到dt2。

```
from Spotfire.Dxp.Data import *
import clr
clr.AddReference('System')
from System.IO import FileStream, FileMode, File, MemoryStream, SeekOrigin, StreamWriter
import System.String
from Spotfire.Dxp.Data.Import import TextDataReaderSettings
from Spotfire.Dxp.Data.Import import TextFileDataSource
from Spotfire.Dxp.Data import *
from datetime import datetime

dt = Document.Data.Tables["数据表"]
column1Cursor = DataValueCursor.CreateFormatted(dt.Columns["列1"])
column2Cursor = DataValueCursor.CreateFormatted(dt.Columns["列2"])
column3Cursor = DataValueCursor.CreateFormatted(dt.Columns["列3"])
column4Cursor = DataValueCursor.CreateFormatted(dt.Columns["列4"])
column5Cursor = DataValueCursor.CreateNumeric(dt.Columns["列5"])

#获取标记的行
markings = Document.ActiveMarkingSelectionReference.GetSelection(dt)
markedata = list();

for row in dt.GetRows(markings.AsIndexSet(),column3Cursor):
	value = column3Cursor.CurrentValue
	if value <> str.Empty:
		markedata.Add(value)
print(markedata)


#修改表中标记数据的值
stream = MemoryStream();
csvWriter = StreamWriter(stream)#, Encoding.UTF8)
csvWriter.WriteLine("列1,列2,列3,列4,列5\r\n")

for row in dt.GetRows(column1Cursor, column2Cursor, column3Cursor, column4Cursor, column5Cursor):
	column1 = column1Cursor.CurrentValue
	column2 = column2Cursor.CurrentValue
	column3 = column3Cursor.CurrentValue
	column4 = datetime.strptime(column4Cursor.CurrentValue, "%Y-%m-%d").date()
	column5 = column5Cursor.CurrentValue
	if column3 in markedata:
		Document.Properties["Tip1"] = "更改" + column2 + "为："+ ChangeValue
		column2 = ChangeValue

	csvWriter.WriteLine(System.String.Format("{0},{1},{2},{3},{4}", column1, column2, column3, column4, column5))

settings = TextDataReaderSettings()
settings.Separator = ","
settings.AddColumnNameRow(0)
settings.ClearDataTypes(False)
settings.SetDataType(0,DataType.String)
settings.SetDataType(1,DataType.String)
settings.SetDataType(2,DataType.String)
settings.SetDataType(3,DataType.Date)
settings.SetDataType(4,DataType.Integer)

csvWriter.Flush()
stream.Seek(0, SeekOrigin.Begin)
fs = TextFileDataSource(stream, settings)
dt = Document.Data.Tables["数据表"].ReplaceData(fs)

#重置
markedata = list()
```



