Dim cn ' As ADODB.Connection
Set cn = CreateObject ("ADODB.Connection")
'Set cn = New ADODB.Connection
With cn
    '.Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "DSN=t4;" 'thats my dsn for sure
    .Open
End With
Debug.Print vbCrLf & "established connection at " & Now()

Dim rs ' As ADODB.Recordset
Set rs = CreateObject("ADODB.Recordset")
Set rs = cn.Execute("select * from wording")

Debug.Print vbCrLf & "executed query at " & Now()

rs.MoveFirst

Do Until rs.EOF         'loop to cycle through the table
  Debug.Print rs.Fields.Item("id") & _
    vbTab & rs.Fields.Item("Field1") & _
        vbTab & rs.Fields.Item("Field2")
  rs.MoveNext
Loop

Debug.Print vbCrLf & "printed output at " & Now()

rs.Close
cn.Close

Debug.Print vbCrLf & "closed connection at " & Now()

