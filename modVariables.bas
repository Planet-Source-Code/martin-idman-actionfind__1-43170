Attribute VB_Name = "modVariables"
' Force declaration
Option Explicit

' Dimesion public variables
Public cn1 As New ADODB.Connection
Public RS1 As ADODB.Recordset
Public strSQL As String
Public intTemp As Integer
Public strDSNOld As String
Public strDSNNew As String
Public bolDatabaseOpen As Boolean
Public strSearchFor As String
Public fldField As Field
Public strField As String


