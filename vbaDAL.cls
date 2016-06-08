VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbaDAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private objConnection As ADODB.Connection
Public objCommand As ADODB.Command
Public rowsAffected As Integer

Public Function Execute(qrySQL As String) As ADODB.RecordSet
    Dim rs As ADODB.RecordSet
    Dim connStr As String
    rowsAffected = 0
    objConnection.Open
    Set objCommand = New ADODB.Command
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = qrySQL
    objCommand.CommandType = adCmdText
    Set rs = objCommand.Execute(rowsAffected)
    Set Execute = rs
End Function

Public Function CloseConnection()
    objConnection.Close
End Function

Public Function ExecuteSP(procName As String) As Integer
    rowsAffected = 0
    objConnection.Open
    objCommand.CommandType = adCmdStoredProc
    objCommand.CommandText = procName
    objCommand.ActiveConnection = objConnection
    objCommand.Execute (rowsAffected)
    objConnection.Close
    ExecuteSP = rowsAffected
End Function

Private Sub Class_Initialize()
    Dim connStr As String
    Set objConnection = New ADODB.Connection
    connStr = "Provider=SQLOLEDB;Data Source=[ServerName];Initial Catalog=[DatabaseName];Trusted_Connection=yes;"
    objConnection.ConnectionString = connStr
End Sub
