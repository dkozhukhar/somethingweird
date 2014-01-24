Option Explicit

Option Compare Database

Sub Main()

Dim db As DAO.Database
'Set db = CurrentDb
Set db = OpenDatabase("C:\Users\dmytro\Documents\Database2.accdb")

Dim sql As String
'sql = "SELECT * FROM [Table1]"

Dim qd As DAO.QueryDef
Set qd = db.QueryDefs("Table1 Query")

Dim rs As DAO.Recordset
'Set rs = db.OpenRecordset(sql)
Set rs = qd.OpenRecordset()

Dim dt As Date
Dim bn As String


Do Until rs.EOF         'loop to cycle through the table
  'subroutine(rs![datestamp], rs![bank])
  dt = rs![datestamp]
  bn = rs![bank]
  subroutine dt, bn
  rs.MoveNext
Loop

rs.Close
db.Close

Debug.Print vbCrLf & "closed connection at " & Now()

End Sub


Public Sub subroutine(ByVal datestamp As Date, ByVal bankname As String)

Dim db As DAO.Database
Set db = CurrentDb

Dim sql As String
sql = "SELECT * FROM [Table1] WHERE [datestamp]=#" & Format(datestamp, "yyyy-mm-dd") & "#" & _
    " AND [bank] = '" & bankname & "'"

Dim datestamp1 As Date
Dim bankname1 As String

Dim rs As DAO.Recordset
Set rs = db.OpenRecordset(sql)
datestamp1 = rs![datestamp]
bankname1 = rs![bank]
rs.Close
Set rs = Nothing

Dim rst As DAO.Recordset
Set rst = db.OpenRecordset("wording", dbOpenTable)
rst.AddNew
rst![Field1] = "x"
rst![Field2] = "#" & Format(datestamp, "yyyy-mm-dd") & "#" & _
    " AND [bank] = '" & bankname & "'"
rst![Field3] = datestamp
rst.Update
rst.Close
Set rst = Nothing

db.Close

Debug.Print "complited" & bankname1 & " " & Format(datestamp1, "yyyy-mm-dd") & vbCrLf & _
        "closed connection at " & Now()

End Sub
