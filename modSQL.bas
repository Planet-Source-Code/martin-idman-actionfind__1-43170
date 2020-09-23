Attribute VB_Name = "modSQL"
' Force variable declaration
Option Explicit

' Sub to open connection to database
Public Sub OpenConnection(ByVal strOpenConnectionDSN, ByVal lngOpenConnectionTimeout)
   With cn1
      .CursorLocation = adUseServer
      .Mode = adModeReadWrite
      .ConnectionTimeout = lngOpenConnectionTimeout
      .CommandTimeout = lngOpenConnectionTimeout
   End With
   cn1.Open ("DSN=" & strOpenConnectionDSN)
End Sub

' Sub to close connection
Public Sub CloseConnection()
   cn1.Close
   Set cn1 = Nothing
End Sub

' Sub to run SQL commands with open recordset
Public Sub ExecSQL(ByRef RS As ADODB.Recordset, ByVal strExecSQL As String)
   ' Debug mode for internal debugging
   cn1.CursorLocation = adUseClient
   Set RS = cn1.Execute(strExecSQL)
End Sub

