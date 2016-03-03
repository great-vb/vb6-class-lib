# vb6-class-lib
a lib with classes from project "vb6-classes"

```vb
'连接到数据库
Dim db As New AdodbHelper
db.SetConnToFile "数据库文件路径"
'执行查询语句得到记录集
Dim res As Adodb.Recordset
Set res = db.ExecQuery("Select * From `students`")
```
