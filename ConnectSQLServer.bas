Attribute VB_Name = "ConnectSQLServer"
Option Explicit

Sub ConnectSQLServer()

    Dim cn As ADODB.connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.connection
    Set rs = New ADODB.Recordset
            
    cn.connectionstring = _
    "Provider=MSOLEDBSQL;" & _
    "Server=server,port;" & _
    "Database=database;" & _
    "UID=user;" & _
    "PWD=password;"
    
    cn.Open
    
    If cn.State = 1 Then
        Debug.Print "Connected!"
    End If
    
    Set rs = cn.Execute("select * from usr;")
    
    Range("A1").CopyFromRecordset rs
    
    rs.Close
    cn.Close

End Sub
