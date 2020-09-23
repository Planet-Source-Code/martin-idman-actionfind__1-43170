VERSION 5.00
Begin VB.Form frmForm 
   Caption         =   "Action Find"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   Icon            =   "frmActionFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFrame 
      Caption         =   "Start and progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   780
      Index           =   3
      Left            =   60
      TabIndex        =   14
      Top             =   6150
      Width           =   5340
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   420
         Left            =   4335
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblLabel 
         Caption         =   "Searching table"
         Height          =   315
         Index           =   5
         Left            =   105
         TabIndex        =   17
         Top             =   300
         Width           =   1980
      End
      Begin VB.Label lblTable 
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2085
         TabIndex        =   16
         Top             =   285
         Width           =   2175
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1740
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   4335
      Width           =   5340
      Begin VB.OptionButton optShow 
         Caption         =   "All"
         Height          =   195
         Index           =   1
         Left            =   4125
         TabIndex        =   21
         Top             =   285
         Width           =   810
      End
      Begin VB.OptionButton optShow 
         Caption         =   "One per table and field"
         Height          =   195
         Index           =   0
         Left            =   2130
         TabIndex        =   20
         Top             =   270
         Value           =   -1  'True
         Width           =   1980
      End
      Begin VB.ListBox lstResult 
         BackColor       =   &H8000000B&
         Height          =   1035
         ItemData        =   "frmActionFind.frx":030A
         Left            =   2130
         List            =   "frmActionFind.frx":030C
         TabIndex        =   13
         Top             =   570
         Width           =   3075
      End
      Begin VB.Label lblLabel 
         Caption         =   "(Table - Field)"
         Height          =   315
         Index           =   8
         Left            =   255
         TabIndex        =   22
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Caption         =   "Found occurances"
         Height          =   315
         Index           =   7
         Left            =   135
         TabIndex        =   19
         Top             =   585
         Width           =   1935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Show result rows"
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   18
         Top             =   285
         Width           =   1935
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Search for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1395
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   2835
      Width           =   5340
      Begin VB.OptionButton optOperator 
         Caption         =   "Exact match (=)"
         Height          =   195
         Index           =   0
         Left            =   2070
         TabIndex        =   26
         Top             =   660
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton optOperator 
         Caption         =   "Contains (%)"
         Height          =   195
         Index           =   1
         Left            =   3585
         TabIndex        =   25
         Top             =   675
         Width           =   1530
      End
      Begin VB.TextBox txtSearchFor 
         Height          =   315
         Left            =   2055
         TabIndex        =   9
         Top             =   240
         Width           =   3090
      End
      Begin VB.ComboBox cboFieldType 
         Height          =   315
         ItemData        =   "frmActionFind.frx":030E
         Left            =   2070
         List            =   "frmActionFind.frx":031E
         TabIndex        =   8
         Text            =   "Text"
         Top             =   960
         Width           =   3090
      End
      Begin VB.Label lblLabel 
         Caption         =   "Search for (no ' or "")"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Search fieldtype"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   990
         Width           =   1935
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Datasource and tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2700
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   5340
      Begin VB.CheckBox chkFilter 
         Height          =   315
         Left            =   2070
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.ListBox lstTables 
         Height          =   1620
         ItemData        =   "frmActionFind.frx":033C
         Left            =   2070
         List            =   "frmActionFind.frx":033E
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   960
         Width           =   3105
      End
      Begin prjActionFind.EBDSNCombo cboDSN 
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         Top             =   240
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   556
      End
      Begin VB.Label lblLabel 
         Caption         =   "select/deselect all rows)"
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblLabel 
         Caption         =   "(Doubleclick row 2 to"
         Height          =   315
         Index           =   9
         Left            =   270
         TabIndex        =   23
         Top             =   1215
         Width           =   1935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Data Source Name (DSN)"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Show system tables"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   630
         Width           =   1935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Select tables to search"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   990
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +—————————————————————————————————————————————————————————————————+
' | Program  : Action Find                                          |
' | Copyright: Martin Idman                                         |
' | Specifics: Works on SQL Server only. Via ODBC                   |
' | Usage    : Search selected tables in a database for a value     |
' |            For example: You know that somewhere there's a field |
' |            containing the text "Hello" but you don't know where |
' |            Now you can find it easily.                          |
' +—————————————————————————————————————————————————————————————————+

' Force variable declaration
Option Explicit

' Sub when form loads
Private Sub Form_Load()
   ' Set DSN filter to SQL Server for DSN combobox
   cboDSN.DriverFilter = "SQL Server"
End Sub

' Sub to handle click on combobox
Private Sub cboDSN_LostFocus()
   ' Load available tables when control lost focus
   If cboDSN.DSN <> "" Then LoadTables
End Sub

' Sub to load available tables
Private Sub LoadTables()
   ' Set newDSN to selected DSN
   strDSNNew = cboDSN.DSN
   ' Check if a new DSN is selected
   If strDSNOld <> strDSNNew Then
      ' Clear combobox
      lstTables.Clear
      ' Connect to database if it's not open
      If bolDatabaseOpen = True Then
         CloseConnection
      End If
      OpenConnection cboDSN.DSN, 10
      bolDatabaseOpen = True
      ' Get all tables
      strSQL = ""
      strSQL = strSQL + "SELECT [name] FROM dbo.sysobjects WHERE "
      strSQL = strSQL + "[type] ='U' "
      If chkFilter.Value = vbChecked Then
         strSQL = strSQL & "OR [type] = 'S' "
      End If
      strSQL = strSQL & "ORDER BY [name]"
      ' Run query
      ExecSQL RS1, strSQL
      ' Loop tables and add to combobox
      While Not RS1.EOF
         lstTables.AddItem RS1![Name]
         RS1.MoveNext
      Wend
      ' Close recordset
      RS1.Close
      ' Set oldDSN to newDSN
      strDSNOld = strDSNNew
   ' No change made
   End If
End Sub

' Handle search button
Private Sub cmdSearch_Click()
   ' Change text on button
   If cmdSearch.Caption = "&Search" Then
      cmdSearch.Caption = "&Stop"
      Search
   Else
      cmdSearch.Caption = "&Search"
      Exit Sub
   End If
End Sub

' Sub to run the search
Private Sub Search()
   ' Dimension variables
   Dim intSearch As Integer
   Dim bolSearch() As Boolean
   ' Clear resultlist
   lstResult.Clear
   ' Remove forbidden "'" in search-text
   txtSearchFor.Text = Replace(txtSearchFor.Text, "'", "")
   ' Loop all itemps in tables listbox
   For intTemp = 0 To lstTables.ListCount - 1
      ' Let system have resources
      Sleep 0.1
      ' Change caption on button
      If cmdSearch.Caption = "&Search" Then GoTo SearchStopped
      ' Check if this table is selected
      If lstTables.Selected(intTemp) Then
         ' Store table name to label
         lblTable.Caption = lstTables.List(intTemp)
         ' Refresh label
         lblTable.Refresh
         ' Move to row in list
         lstTables.ListIndex = intTemp
         ' Make query with all corresponding fiels in this table
         strSQL = MakeQuery(lstTables.List(intTemp), cboFieldType.Text, txtSearchFor.Text)
         ' Function returned no fields found, goto next table
         If strSQL = "" Then GoTo NextTable
         ' Run query
         ExecSQL RS1, strSQL
         ' Redim variable to hold all hits on fieldname (0-n)
         ReDim bolSearch(RS1.Fields.Count - 1)
         ' Loop and add results
         While Not RS1.EOF
            ' Reset counter
            intSearch = 0
            ' Get which field/s contains searched string
            For Each fldField In RS1.Fields
               ' Check if field contains value
               Select Case cboFieldType.Text
                  Case "Text" ' Text type fields
                     If optOperator(0).Value = True Then
                        ' Get filedname that matches search exact (=)
                        strField = IIf(fldField.Value = Replace(strSearchFor, "'", ""), fldField.Name, "")
                     Else
                        ' Get filedname that matches search like (%)
                        strField = IIf(InStr(1, UCase(fldField.Value), UCase(Replace(Replace(strSearchFor, "'", ""), "%", ""))) > 0, fldField.Name, "")
                     End If
                  Case "Number" ' Numeric type fields
                     ' Get filedname that matches search exact (=)
                     strField = IIf(fldField.Value = strSearchFor, fldField.Name, "")
                  Case "Date" ' Date type fields
                     ' Get filedname that matches search exact (=)
                     strField = IIf(Mid(fldField.Value, 1, 10) = Replace(strSearchFor, "'", ""), fldField.Name, "")
                  Case "Time" ' Time type fields
                     ' Get filedname that matches search exact (=)
                     strField = IIf(Mid(fldField.Value, 12, 8) = Replace(strSearchFor, "'", ""), fldField.Name, "")
               End Select
               ' If value was found
               If strField <> "" Then
                  ' Check if it's one match per table or not
                  If optShow(0).Value = True And bolSearch(intSearch) = True Then
                     ' This table and field already displayed, do nothing
                  Else
                     ' Add taken to variable
                     bolSearch(intSearch) = True
                     ' Att to result list
                     lstResult.AddItem lblTable.Caption & " - " & fldField.Name
                     ' Move to last row in result list
                     lstResult.ListIndex = lstResult.ListCount - 1
                     ' Hide highlight
                     lstResult.ListIndex = -1
                     ' Refresh result list
                     lstResult.Refresh
                  End If
               End If
               ' Add 1 to counter
               intSearch = intSearch + 1
            Next
            ' Move to next record
            RS1.MoveNext
         Wend
         ' Close recordset
         RS1.Close
      End If
NextTable:
   Next intTemp
   ' Change button caption
   cmdSearch.Caption = "&Search"
SearchStopped:
   ' Message ready
   MsgBox "Search finished/stopped, returned " & CStr(lstResult.ListCount) & " row/s!", vbInformation, "Search results"
End Sub

' Function to make query
Private Function MakeQuery(ByVal strMakeQueryTable As String, _
                           ByVal strMakeQueryFieldType As String, _
                           ByVal strMakeQuerySearchFor As String) As String
   ' Dimension variables
   Dim strMakeQuerySQL As String
   Dim strMakeQueryIn As String
   ' Convert field-types to xtype in syscolumns table
   Select Case strMakeQueryFieldType
      Case "Text" ' All text-based field-types (except ntext and text which don't work)
         strMakeQueryIn = "167,175,231,239"
         ' Store search for text and add "'" because this is text we're searching
         If optOperator(0).Value = True Then
            strSearchFor = "'" & strMakeQuerySearchFor & "'"
         Else
            strSearchFor = "'%" & strMakeQuerySearchFor & "%'"
         End If
      Case "Number" ' All numeric-based field-types
         strMakeQueryIn = "48,52,56,59,60,62,106,108,122"
         ' Store search for text
         strSearchFor = strMakeQuerySearchFor
      Case "Date", "Time" ' All date/time-based field-types
         strMakeQueryIn = "58,61"
         ' Store search for text and add "'" because this is text we're searching
         strSearchFor = "'" & strMakeQuerySearchFor & "'"
   End Select
   ' Make query to get all fields with corresponding type
   strMakeQuerySQL = ""
   strMakeQuerySQL = strMakeQuerySQL & "SELECT b.name FROM "
   strMakeQuerySQL = strMakeQuerySQL & "sysobjects a LEFT JOIN syscolumns b ON b.id = a.id "
   strMakeQuerySQL = strMakeQuerySQL & "WHERE a.name ='" & strMakeQueryTable & "' AND "
   strMakeQuerySQL = strMakeQuerySQL & "b.xtype IN (" & strMakeQueryIn & ")"
   ' Run query
   ExecSQL RS1, strMakeQuerySQL
   ' If no corresponding fields was found
   If RS1.EOF Then
      ' Return empty string
      MakeQuery = ""
      ' Close recordset
      RS1.Close
      ' Exit function to calling sub
      Exit Function
   End If
   ' Default values
   strMakeQuerySQL = "SELECT * FROM " & strMakeQueryTable & " WHERE "
   ' Loop tables and add to string
   While Not RS1.EOF
      ' Add field to query
      Select Case strMakeQueryFieldType
         Case "Text", "Number" ' Text and numbers, use "="
            strMakeQuerySQL = strMakeQuerySQL & RS1![Name] & IIf(optOperator(0).Value = True, " = ", " LIKE ") & strSearchFor & " OR "
         Case "Date" ' Convert to date(10)
            strMakeQuerySQL = strMakeQuerySQL & "CONVERT(CHAR(10), " & RS1![Name] & ", 120) = " & strSearchFor & " OR "
         Case "Time" ' Convert to time(8)
            strMakeQuerySQL = strMakeQuerySQL & "CONVERT(CHAR(8), " & RS1![Name] & ", 108) = " & strSearchFor & " OR "
      End Select
      ' Move to next record
      RS1.MoveNext
   Wend
   ' Close open recordset
   RS1.Close
   ' Remove last "OR"
   strMakeQuerySQL = Left(strMakeQuerySQL, Len(strMakeQuerySQL) - 4)
   ' Return value to calling sub
   MakeQuery = strMakeQuerySQL
End Function

' Sub to enable / disable all tables
Private Sub lstTables_DblClick()
   ' Check if first is selected
   If lstTables.Selected(0) Then
      ' Deselect
      For intTemp = 0 To lstTables.ListCount - 1
         lstTables.Selected(intTemp) = False
      Next intTemp
   Else
      ' Select
      For intTemp = 0 To lstTables.ListCount - 1
         lstTables.Selected(intTemp) = True
      Next intTemp
   End If
   ' Move to top of list
   lstTables.ListIndex = 0
End Sub

' Sub to enable / disable showing of system tables
Private Sub chkFilter_Click()
   strDSNOld = ""
   LoadTables
End Sub

' Sub to disable typing in combobox
Private Sub cboFieldType_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

