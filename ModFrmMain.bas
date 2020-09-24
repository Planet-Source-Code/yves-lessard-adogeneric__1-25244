Attribute VB_Name = "ModFrmMain"
Option Explicit
'******************************************************************************
'** Module.........: ModForm
'**
'** Description....: Routines for FrmMain
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 19-Jul-2001.
'**
'** Modifications..:
'** Version........: 1.0.0.A
'**
'******************************************************************************
Private Const m_ClassName = "ModFrmMain"

Private MonAdo As CadoGeneric
Private TempRS As ADODB.Recordset
Private cHourglass As CmouseGlass

Public Sub FillCitiesCombo(ByVal TheCombo As ComboBox)
'******************************************************************************
'** SubRoutine.....: FillCitiesCombo
'**
'** Description....: Fill a combo with city name
'**                  the fatest way possible
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 19-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** TheCombo            ComboBox   W     The Combo to fill
'******************************************************************************
On Error GoTo ErrorSection
'** Glass Cursor

Set cHourglass = New CmouseGlass
Set MonAdo = New CadoGeneric
Set TempRS = New ADODB.Recordset
Dim szQuery As String
szQuery = "SELECT * FROM TblCity"

TheCombo.Clear
With MonAdo
    .Dbname = App.Path & "\Client.MDB"
    .SetConnectType ACCESS
End With

If MonAdo.ReadOnlyRecord(szQuery, TempRS) Then
    While Not TempRS.EOF
        '** We Can also use this one
        'TheCombo.AddItem TempRS.Fields("CityName").Value
        '** Collect is the fastest in the west
        TheCombo.AddItem TempRS.Collect("CityName")
        TheCombo.ItemData(TheCombo.ListCount - 1) = TempRS.Collect(0)
        TempRS.MoveNext
    Wend
End If

TempRS.Close

'********************
'Exit Point
'********************
ExitPoint:
'** Free up memory
Set cHourglass = Nothing
Set MonAdo = Nothing
Set TempRS = Nothing
Exit Sub
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "FillCitiesCombo", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub


'*********************************
'****    Error(s) Handling    ****
'*********************************

Public Sub ShowError(ErrorNumber As Long, ErrorMsg As String _
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

MsgBox ErrorMessage, vbOKOnly + vbExclamation, ErrorTitle
App.LogEvent ErrorTitle & ": " & ErrorMessage, LogEventType

ExitPoint:
Exit Sub

ErrorSection:
Resume ExitPoint

End Sub

