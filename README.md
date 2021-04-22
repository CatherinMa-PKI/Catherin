# 修改功能
### 方法一：只把修改的那一条回写进数据库。
最好也是竖表进行修改，这样的话，比较好留下修改的log，而且ironPython不用定义很多列，不容易出错。但是如果不想进行过多Pivot和Unpivot的操作的话，可以只把需要修改的列拿出来，用横表的方式写进脚本。

### 方法二：把整张表都回写进数据库，之后找到那一条进行sql修改。
首先把整张表以纵表的形式写进数据库，然后定义主键之后，在库里面定位，修改。这种方法可能修改一次需要的时间比较长，建议分成另外一个modify模板，修改完的数据再导出给应用模板用。

### 方法三：通过写stream的方式写修改后的表格
建议比较小的表格使用这个方法，因为stream是写全张表。首先需要有一张原始表，然后复制原始表为新表dt2用来前台展示和mark，mark的数据修改完之后存在dt3，然后把dt3的数据给替换到dt2。

///
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

///




