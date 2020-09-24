VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CadoGeneric"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'******************************************************************************
'** Module.........: CAdoGeneric
'** Description....: Dbase Methodes and properties
'**                  Reference to ADO 2.5
'**
'** Author, date...: Yves Lessard , 20-Aug-1999.
'**
'** Version........: 1.0.0.A
'** Modifications..: Y Lessard
'** 1.0.1 - 27-Feb-2001
'**                  Add error handling on all Ado stuff
'**                  Fix Constructor and Destructor
'**
'** 1.0.2 - 19-July-2001
'**                  All public methodes return a Boolean Value
'**
'**
'** Property             Data Type     Description
'** ------------------   ---------     --------------------------------------
'** Dbname               String        DataBase Name
'** ServerName           String        Server Name (SQL-Server)
'** Password             String        Password for ODBC and SQl-Server
'** USerID               String        The User for ODBC & SQL-Server
'** ErrorInfo            String        Return Error Info
'** ErrorNumber          Long          Return Error Number
'**
'** Method(Public)       Description
'** ------------------   --------------------------------------
'** RunSQL               Execute SQL without recordset return
'** GetBatchRecord       Get a recordset for Later Read/Write (Batch)
'** SaveBatchRecord      Update recordset from (GetBatchRecord) in BatchMode
'** ReadOnlyRecord       Get a recordset for Read Only (Fast)
'** GetRecord            Get a Recordset for Read/Write
'** SaveRecord           Save a recordset
'** SetConnectType       The Conection Type (Access, ODBC, SQL-Server)
'******************************************************************************
Private Const m_ClassName = "CadoGeneric"
Private Const vbObjectError = 513

Private Const ERR_UPDATE_COLLISION = -2147217864
Private Const ERR_CONFLICTS_FOUND = vbObjectError + 1

Private m_sDbname As String
Private m_sConnect As String
Private AdoConnect As ADODB.Connection
Private LocRecset As ADODB.Recordset
Attribute LocRecset.VB_VarHelpID = -1
Private m_CType As Integer
Private m_ErrorNumber As Long
Private m_szErrorInfo As String

Enum CType
    [ACCESS] = 0
    [SQLServer] = 1
    [ODBC] = 2
End Enum

Private m_szConnection As String
Private m_bChange As Boolean
Private mvarPassword As String
Private mvarUserID As String
Private mvarServerName As String


'**************************
'****    Properties    ****
'**************************

Public Property Get ErrorInfo() As String
'******************************************************************************
'** Description....: Return error info as text
'** Author, date...: Yves Lessard  19-Jul-2001.
'******************************************************************************
    ErrorInfo = m_szErrorInfo
End Property

Public Property Get ErrorNumber() As Long
'******************************************************************************
'** Description....: Return Error Number (Like HRESULT)
'** Author, date...: Yves Lessard  19-Jul-2001.
'******************************************************************************
    ErrorNumber = m_ErrorNumber
End Property

Public Property Let ServerName(ByVal vData As String)
'******************************************************************************
'** Description....: The Server Name SQL-Server
'** Author, date...: Yves Lessard  27-Feb-2001.
'******************************************************************************
    mvarServerName = vData
    m_bChange = True
End Property

Public Property Get ServerName() As String
'******************************************************************************
'** Description....: Get the Server Name
'** Author, date...: Yves Lessard  21-Jun-2000.
'******************************************************************************
    ServerName = mvarServerName
End Property

Public Property Let PASSWORD(ByVal vData As String)
'******************************************************************************
'** Description....: Assign Password for ODBC ou SQL
'** Author, date...: Yves Lessard  27-Feb-2001.
'******************************************************************************
    mvarPassword = vData
    m_bChange = True
End Property

Public Property Get PASSWORD() As String
'******************************************************************************
'** Description....: Get the Password
'** Author, date...: Yves Lessard  19-Jul-2001.
'******************************************************************************
    PASSWORD = mvarPassword
End Property

Public Property Let UserID(ByVal vData As String)
'******************************************************************************
'** Description....: Assign UserID for ODBC ou SQL
'** Author, date...: Yves Lessard  27-Feb-2001.
'******************************************************************************
    mvarUserID = vData
    m_bChange = True
End Property

Public Property Get UserID() As String
'******************************************************************************
'** Description....: Get the UserID
'** Author, date...: Yves Lessard  17-Jun-2000.
'******************************************************************************
    UserID = mvarUserID
End Property

Public Property Let Dbname(ByVal sData As String)
'******************************************************************************
'** Description....: Assign Dbase name
'** Author, date...: Yves Lessard , 27-Feb-2001.
'******************************************************************************
    m_sDbname = sData
    m_bChange = True
End Property

Public Property Get Dbname() As String
'******************************************************************************
'** Description....: Get the Dbase name
'** Author, date...: Yves Lessard , 12-Aug-1999.
'******************************************************************************
Dbname = m_sDbname
End Property



'******************************
'****    Methodes Public   ****
'******************************


Public Sub SetConnectType(ByVal vData As CType)
'******************************************************************************
'** SubRoutine.....: SetConnectType
'**
'** Description....: Choose the Connection Type
'**
'** Author, date...: Yves Lessard , 21-Jun-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** vData               Enum       R     Access, ODBC, or SQL-Server connection
'******************************************************************************
    m_CType = vData
    '** Change the Connection Type
    m_bChange = True
End Sub

Public Function ReadOnlyRecord(ByVal strQuery As String, _
                                ByRef RecSet As ADODB.Recordset) As Boolean
'******************************************************************************
'** Function.......: ReadOnlyRecord
'**
'** Description....: Get a recordset for read only
'**
'** Author, date...: Yves Lessard, 17-nov-1999.
'**
'** Modifications..: 19-July-2001
'**                  Return Boolean result
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** strQuery            String     R     The query
'** RecSet              Recordset  W     The Client Recordset
'** readOnlyRecord      Boolean    W     TRUE -> Success else Error
'******************************************************************************
On Error GoTo ErrorSection

Dim bResult As Boolean
'** Default No Error
bResult = True
ClearError
ReConnect

'**YL Open Recordset
With LocRecset
    .ActiveConnection = AdoConnect
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Source = strQuery
    .Open
End With

Set LocRecset.ActiveConnection = Nothing
'** Transfert The Recordset
Set RecSet = LocRecset.Clone
LocRecset.Close

'********************
'Exit Point
'********************
ExitPoint:
ReadOnlyRecord = bResult
Exit Function
'********************
'Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "ReadOnlyRecord", m_ClassName, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "ReadOnlyRecord", m_ClassName, vbLogEventTypeError
    End Select
End If
bResult = False
Resume ExitPoint
End Function

Public Function RunSQL(ByVal strSQL As String) As Boolean
'******************************************************************************
'** SubRoutine.....: RunSQL
'**
'** Description....: Execute a query only
'**
'**
'** Author, date...: Yves Lessard , 17-Aug-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** strSQL              String     R     The query
'** RunSQL              Boolean    W     TRUE-> No error
'******************************************************************************
On Error GoTo ErrorSection
Dim bResult As Boolean
'** Default No Error
bResult = True
ClearError
ReConnect

AdoConnect.BeginTrans
AdoConnect.Execute (strSQL)
AdoConnect.CommitTrans

'********************
' B) Exit Point
'********************
ExitPoint:
RunSQL = bResult
Exit Function
'********************
' C) Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "RunSQL", m_ClassName, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "RunSQL", m_ClassName, vbLogEventTypeError
    End Select
End If
bResult = False
AdoConnect.RollbackTrans
Resume ExitPoint

End Function


Public Function GetBatchRecord(ByVal strQuery As String, _
                                ByRef RecSet As ADODB.Recordset) As Boolean
'******************************************************************************
'** Function.......: GetBatchRecord
'**
'** Description....: Get a recordset for
'**                  Batchmode Read/Write operation
'**                  Use SaveBatchRecord for Saving
'**
'** Author, date...: Yves Lessard, 16-nov-1999.
'**
'** Modifications..: 27-Feb-2001
'**                  Fix New recordset stuff & ExitPoint
'**                  19-July-2001
'**                  Use Recordset on Client Side & return Boolean
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**  strQuery           String     R      The Query
'**  RecSet             Recordset  W      The destination Recordset
'**  GetBatchRecord     Boolean    W      TRUE -> No error else FALSE
'******************************************************************************
On Error GoTo ErrorSection

Dim bResult As Boolean
'** Default No Error
bResult = True
ClearError
ReConnect

'**YL Open Recordset for batch mode
With LocRecset
    .ActiveConnection = AdoConnect
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset      'Need the RecordCount Property
    .LockType = adLockBatchOptimistic
    .Source = strQuery
    .Open
End With

Set LocRecset.ActiveConnection = Nothing
'** Transfert The Recordset
Set RecSet = LocRecset.Clone
LocRecset.Close

'********************
'Exit Point
'********************
ExitPoint:
GetBatchRecord = bResult
Exit Function
'********************
'Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "GetBatchRecord", m_ClassName, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "GetBatchRecord", m_ClassName, vbLogEventTypeError
    End Select
End If
bResult = False
Resume ExitPoint
End Function

Public Function GetRecord(ByVal strQuery As String, _
                          ByRef RecSet As ADODB.Recordset) As Boolean
'******************************************************************************
'** Function.......: GetRecord
'**
'** Description....: Get a recordset for Read/Write operation
'**                  Use SaveRecord for saving
'**
'** Author, date...: Yves Lessard, 16-nov-1999.
'**
'** Modifications..: 27-Feb-2001
'**                  Remove new recordset stuff & Fix exitpoint
'**                  19-July-2001
'**                  Return Boolean result
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**  strQuery           String     R      The Query
'**  RecSet             Recordset  W      The Client Recordset
'**  GetRecord          Boolean    W      TRUE- > OK else Error
'******************************************************************************
On Error GoTo ErrorSection

Dim bResult As Boolean
'** Default No Error
bResult = True
ClearError
ReConnect

'**YL Open Recordset
With LocRecset
    .ActiveConnection = AdoConnect
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = strQuery
    .Open
End With

Set LocRecset.ActiveConnection = Nothing
'** Transfert The Recordset
Set RecSet = LocRecset.Clone
LocRecset.Close

'********************
'Exit Point
'********************
ExitPoint:
GetRecord = bResult
Exit Function
'********************
'Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "GetRecord", m_ClassName, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "GetRecord", m_ClassName, vbLogEventTypeError
    End Select
End If
bResult = False
Resume ExitPoint
End Function

Public Function SaveBatchRecord(ByVal AdoRS As ADODB.Recordset) As Boolean
'******************************************************************************
'** SubRoutine.....: SaveBatchRecord
'**
'** Description....: Save Recordset in Batch Mode
'**                  If no Error then Return TRUE
'**
'** Author, date...: Yves Lessard , 17-nov-1999.
'**
'** Modifications..: 27-Feb-2001
'**                  Fix recordset and ExitPoint
'**                  19-July-2001
'**                  Use Tip from Alan Gordon Book
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** AdoRS               Recordset   R    The Recordset to save
'** SaveBatchRecord     Boolean     W    TRUE = No Error
'******************************************************************************
On Error GoTo ErrorSection

Dim bResult As Boolean
'** Default No Error
bResult = True
ClearError
ReConnect

With AdoRS
    .MarshalOptions = adMarshalModifiedOnly
    .ActiveConnection = AdoConnect
End With

'** We start transaction
AdoConnect.BeginTrans
AdoRS.UpdateBatch adAffectAll

'** Thanks to Alan Gordon
AdoRS.Filter = adFilterConflictingRecords
'** If All records updated then recordcount = 0
If AdoRS.RecordCount > 0 Then
    Err.Raise ERR_CONFLICTS_FOUND, , AdoRS.RecordCount & " could not be saved"
End If
AdoConnect.CommitTrans

'********************
'Exit Point
'********************
ExitPoint:
Set AdoRS.ActiveConnection = Nothing
SaveBatchRecord = bResult
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case ERR_UPDATE_COLLISION
        '**YL Resync problem so try again
        Err.Clear
        With AdoRS
            .Resync adAffectAllChapters, adResyncUnderlyingValues
            .UpdateBatch
        End With
        Resume ExitPoint
    Case Else
    ShowError Err.Number, Err.Description, "SaveBatchRecord", m_ClassName, vbLogEventTypeError
End Select
bResult = False
'** Error so Rollback
AdoConnect.RollbackTrans
Resume ExitPoint

End Function

Public Function SaveRecord(ByVal AdoRS As ADODB.Recordset) As Boolean
'******************************************************************************
'** SubRoutine.....: SaveRecord
'**
'** Description....: Save Recordset
'**                  Save Record from Function GetRecord
'**
'** Author, date...: Yves Lessard, 17-nov-1999.
'**
'** Modifications..: 27-Feb-2001
'**                  Forgot Resume ExitPoint in error handling
'**                  19-July-2001
'**                  Better error handling
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** AdoRS               Recordset   R    The Recordset to save
'** SaveRecord          Boolean     W    TRUE -> Ok else error
'******************************************************************************
On Error GoTo ErrorSection

Dim bResult As Boolean
'** Default No Error
bResult = True
ClearError
ReConnect

Set AdoRS.ActiveConnection = AdoConnect

'** We start transaction
AdoConnect.BeginTrans
AdoRS.Update

'** Commit the transaction
AdoConnect.CommitTrans
Set AdoRS.ActiveConnection = Nothing

'********************
'Exit Point
'********************
ExitPoint:
SaveRecord = bResult
Exit Function
'********************
'Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "SaveRecord", m_ClassName, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "SaveRecord", m_ClassName, vbLogEventTypeError
    End Select
End If
bResult = False
AdoConnect.RollbackTrans
Resume ExitPoint

End Function



'******************************
'****    Methodes Private  ****
'******************************

Private Sub Class_Initialize()

    Set AdoConnect = New ADODB.Connection
    Set LocRecset = New ADODB.Recordset
    '** Fisrt the connection is to build again
    m_bChange = True
    '** No error
    m_ErrorNumber = 0
End Sub

Private Sub Class_Terminate()
'******************************************************************************
'** SubRoutine.....: Class_Terminate()
'**
'** Description....: Destructor class
'**
'** Author, date...: Yves Lessard , 27-Feb-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** None
'******************************************************************************
On Error GoTo ErrorSection

'** YL Do we need to close recordset ?
If LocRecset.State = adStateOpen Then
    LocRecset.Close
End If
Set LocRecset = Nothing

'*YL Do we need to close the connection ?
If Not AdoConnect Is Nothing Then
    If AdoConnect.State = adStateOpen Then
        AdoConnect.Close
    End If
    Set AdoConnect = Nothing
End If

'********************
'Exit Point
'********************
ExitPoint:
Exit Sub
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "Class_Terminate", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Private Sub ClearError()
'******************************************************************************
'** SubRoutine.....: ClearError
'**
'** Description....: Clear all error variables
'**
'** Author, date...: Yves Lessard , 20-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** None
'******************************************************************************
    m_ErrorNumber = 0
    m_szErrorInfo = ""
End Sub

Private Sub ReConnect()
'******************************************************************************
'** SubRoutine.....: ReConnect
'**
'** Description....: Check if Connection is Open
'**                  or if Connection has Changed
'**
'** Author, date...: Yves Lessard , 19-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** None
'******************************************************************************
On Error GoTo ErrorSection

'** Rebuild connection if necessary
If m_bChange Then
    m_szConnection = RebuildConnection
    '** The Conection Change so we must Close it
    If AdoConnect.State = adStateOpen Then
        AdoConnect.Close
    End If
End If
    
'** If the connection is already Open we skip This
If Not (AdoConnect.State = adStateOpen) Then
    '** Use Connection for rich Info on Error
    With AdoConnect
        .ConnectionString = m_szConnection
        .Open
    End With
End If

'********************
'Exit Point
'********************
ExitPoint:
Exit Sub
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "ReConnect", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Private Function RebuildConnection() As String
'******************************************************************************
'** Function.......: RebuildConnection
'**
'** Description....: Base on Connection Type
'**                  we build the Connect String
'**
'** Author, date...: Yves Lessard , 17-Jun-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** RebuildConnection   String     W     The Connection String
'******************************************************************************
On Error GoTo ErrorSection

Dim szLocalString As String

Select Case m_CType
    Case CType.ACCESS
        '** Now using MDAC 2.5
        If Len(Me.PASSWORD) = 0 Then
            '** No Password Protection
            szLocalString = "Provider=Microsoft.jet.OLEDB.4.0;Data source=" & Dbname
        Else
            '** We have a Password
            szLocalString = "Provider=Microsoft.jet.OLEDB.4.0;" & _
                        "Data Source=" & Dbname & ";" & _
                        "Jet OLEDB:Database Password=" & PASSWORD & ";" & _
                        "admin"
        End If
    Case CType.ODBC
        szLocalString = "Provider=MSDASQL.1;Password=" & PASSWORD & ";Persist Security Info=True;User ID=" _
                        & UserID & ";Data Source=" & Dbname
    Case CType.SQLServer
        szLocalString = "Provider=sqloledb;Data Server=" & ServerName & ";User Id=" & UserID _
                        & ";Password=" & PASSWORD & ";database=" & Dbname
End Select

RebuildConnection = szLocalString
m_bChange = False
'********************
'Exit Point
'********************
ExitPoint:
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "RebuildConnection", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Function


'*********************************
'****    Error(s) Handling    ****
'*********************************

Private Sub ShowError(ErrorNumber As Long, ErrorMsg As String _
                      , ErrorModule As String, ErrorForm As String _
                     , LogEventType As Long, Optional ErrorInfo As Variant)
'******************************************************************************
'** Module.........: ShowError
'** Description....: This routine is used to show the current
'**                  error Message and LOG the error to a file.
'**
'** Author, date...: Yves Lessard , 19-Jul-2001.
'**
'** Name                Type     Acces   Description
'** ------------------  -------  ------  --------------------------------------
'**  ErrorNumber         Long      R      Error Number
'**  ErrorMsg            String    R      Error Message
'**  ErrorModule         String    R      Module name where the error occured
'**  ErrorForm           String    R      Form Name where the error occured
'**  LogEventType        Long      R      Log event type (vbLogEventTypeError ,
'**                                       vbLogEventTypeWarning , vbLogEventTypeInformation)
'**  ErrorInfo           Variant   R      Additional error Information to Display
'**
'******************************************************************************
On Error GoTo ErrorSection
Dim ErrorTitle As String
Dim ErrorMessage As String

ErrorTitle = "ERROR - " & ErrorNumber & " - " & ErrorModule & " - " & ErrorForm
ErrorMessage = "ERROR  " & ErrorNumber & " - " & ErrorMsg

If Not IsMissing(ErrorInfo) Then
    ErrorMessage = ErrorMessage & vbCrLf & ErrorInfo
End If

m_szErrorInfo = ErrorMessage
m_ErrorNumber = ErrorNumber
App.LogEvent ErrorTitle & ": " & ErrorMessage, LogEventType

ExitPoint:
Exit Sub

ErrorSection:
Resume ExitPoint

End Sub

