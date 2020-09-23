VERSION 5.00
Begin VB.UserControl EBDSNCombo 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   ScaleHeight     =   315
   ScaleWidth      =   2430
   Begin VB.ComboBox cboDSN 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   2235
   End
End
Attribute VB_Name = "EBDSNCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[Description]
'   This user control offers a pick list of User and System ODBC DSNs.

'[Author]
'   Richard Allsebrook  <RA>    RichardAllsebrook@earlybirdmarketing.com

'[History]
'   Version 1.0.0   06 Jun 2001
'   Initial Release

'[Declarations]
Option Explicit

'Property Storage
Private strDriverFilter     As String           'DriverFilter

'Mapped Properties
'DSN => cboDSN.Text

'[ODBC API Declarations]
Private Declare Function SQLGetDiagRec Lib "odbc32" ( _
  ByVal HandleType As Integer, _
  ByVal Handle As Long, _
  ByVal RecNumber As Integer, _
  ByVal SQLState As String, _
  ByRef NativeErrorPtr As Long, _
  ByVal MessageText As String, _
  ByVal BufferLength As Integer, _
  ByRef TextLengthPtr As Integer) _
    As Integer
    
Private Declare Function SQLAllocHandle Lib "odbc32" ( _
  ByVal HandleType As Integer, _
  ByVal InputHandle As Long, _
  ByRef OutputHandle As Long _
    ) As Integer
  
Private Declare Function SQLFreeHandle Lib "odbc32" ( _
  ByRef HandleType As Integer, _
  ByRef Handle As Long _
    ) As Integer
    
Private Declare Function SQLSetEnvAttrInteger Lib "odbc32" Alias "SQLSetEnvAttr" ( _
  ByVal EnvironmentHandle As Long, _
  ByVal Attr As Integer, _
  ByVal Value As Long, _
  ByVal StringLength As Integer) _
    As Integer

Private Declare Function SQLDataSources Lib "odbc32" ( _
  ByVal EnvironmentHandle As Long, _
  ByVal Direction As Integer, _
  ByVal ServerName As String, _
  ByVal BufferLength1 As Integer, _
  ByRef NameLength1 As Integer, _
  ByVal Description As String, _
  ByVal BufferLength2 As Integer, _
  ByRef NameLength2 As Integer _
    ) As Integer

Private Const SQL_SUCCESS = 0
Private Const SQL_ERROR = -1

Private Const SQL_HANDLE_ENV = 1

Private Const SQL_ATTR_ODBC_VERSION = 200
Private Const SQL_OV_ODBC2 = 2

Private Const SQL_FETCH_NEXT = 1
Private Const SQL_FETCH_FIRST = 2
    
Private Sub UserControl_InitProperties()

    cboDSN.ListIndex = -1
    strDriverFilter = ""
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        DSN = PropBag.ReadProperty("DSN", "")
        strDriverFilter = .ReadProperty("DriverFilter", "")
    End With
    
    Refresh
    
End Sub

Private Sub UserControl_Resize()

'[Description]
'   Resize the constituent controls to fit the new control size

'[Code]

    With UserControl
        .Height = 315
        cboDSN.Width = .Width
    End With
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "DSN", cboDSN.Text, ""
        .WriteProperty "DriverFilter", strDriverFilter, ""
    End With
    
End Sub

Public Property Get DSN() As String

    DSN = cboDSN.Text
    
End Property

Public Property Let DSN(NewValue As String)

'[Description]
'   Attempt to change the DSN
'   If the DSN does not appear in the list do not set and raise an error

'[Code]

    On Error GoTo ErrorTrap
    
    If Len(NewValue) = 0 Then
        cboDSN.ListIndex = -1
    Else
        cboDSN.Text = NewValue
    End If
    
    PropertyChanged "DSN"
    
    Exit Property
    
ErrorTrap:

    'item not found in collection
    cboDSN.ListIndex = -1
    
End Property

Public Function Refresh()

'[Declarations]
Dim strCurrentDSN           As String   'Currently selected DSN
Dim hEnv                    As Long     'ODBC Environment Handle
Dim intSQLReturn            As Integer

Dim strServerName           As String * 255
Dim intServerNameLen        As Integer
Dim strDescription          As String * 255
Dim intDescriptionLen       As Integer

'[Code]

    'Store the currently selected DSN
    strCurrentDSN = cboDSN.Text
    
    'Build a new list of available DSN
    cboDSN.Clear
    
    If SQLAllocHandle(SQL_HANDLE_ENV, 0, hEnv) = SQL_ERROR Then
        'Failed to allocate Environment Handle
        Err.Raise vbObjectError + 1, "EBDSNCombo_Refresh", "Unable to allocate an ODBC Environment Handle"
        
    Else
        'We have an Environment handle
        '- Inform the Driver Manager that we need ODBC2 conformance
        
        If SQLSetEnvAttrInteger(hEnv, SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC2, 0) = -1 Then
            'Failed to set conformance level
            Err.Raise vbObjectError + 2, "EBDSNCombo_Refresh", "Unable to set ODBC2 conformance"
            
        Else
        
            'We have set the conformance level
            '- Fetch a list of ODBC data sources
            
            'Attempt to fetch first DSN
            intSQLReturn = SQLDataSources(hEnv, SQL_FETCH_FIRST, strServerName, Len(strServerName), intServerNameLen, strDescription, Len(strDescription), intDescriptionLen)
            
            Do While intSQLReturn = SQL_SUCCESS
            
                If Len(strDriverFilter) = 0 _
                  Or Left(strDescription, intDescriptionLen) = strDriverFilter Then
                    'This data source matches the DriverFilter property (or
                    'DriverFilter not set)
                    '- Add it to the list
                    cboDSN.AddItem Left(strServerName, intServerNameLen)
                End If
            
                'Attempt to fetch the next DSN (if any)
                intSQLReturn = SQLDataSources(hEnv, SQL_FETCH_NEXT, strServerName, Len(strServerName), intServerNameLen, strDescription, Len(strDescription), intDescriptionLen)
            Loop
            
        End If
            
        'Free the environment handle
        SQLFreeHandle SQL_HANDLE_ENV, hEnv
    End If
    
    'Attempt to reselect the current DSN
    '(it may not be in the list any more)
    DSN = strCurrentDSN
    
End Function

Private Function RaiseODBCError(hEnv As Long)

'[Description]
'   Displays the first ODBC error (if any)

'[Notes]
'   Used only for debugging purposes (not exposed)
'   As the ODBC API can result in more than one error being generated,
'   it is usual to keep polling the stack to retrieve all the errors.
'   As this function is used purely for debugging purposes, we are only
'   interested in the first error generated.

'[Declarations]
Dim strSQLState             As String * 5       'SQLState at time of error
Dim lngErrorNo              As Long             'ODBC Error No
Dim strMessage              As String * 255     'Error message text
Dim intMessageLen           As Integer          'Length of error message
Dim intSQLReturn            As Integer          'Return state of API call

'[Code]

    'Fetch first Error
    intSQLReturn = SQLGetDiagRec(SQL_HANDLE_ENV, hEnv, 1, strSQLState, lngErrorNo, strMessage, Len(strMessage), intMessageLen)

    If intSQLReturn = SQL_SUCCESS Then
        'Display error
        MsgBox Left("ODBC Error " & lngErrorNo & vbCrLf _
        & strSQLState & " : " & strMessage, intMessageLen)
    End If
    
End Function

Public Property Get DriverFilter() As String

'[Description]
'   Return the DriverFilter Property

'[Code]

    DriverFilter = strDriverFilter
    
End Property

Public Property Let DriverFilter(NewValue As String)

'[Description]
'   Set the DriverFilter property and refresh the list of available DSN

'[Code]

    strDriverFilter = NewValue
    PropertyChanged "DriverFilter"
    
    Refresh
    
End Property
