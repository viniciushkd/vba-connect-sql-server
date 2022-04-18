Attribute VB_Name = "ConnectSQLServer"
Option Explicit

Sub ConnectSQLServer()

    Dim cn As ADODB.connection
    Set cn = New ADODB.connection
            
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
    
    cn.Close

End Sub
