VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CmouseGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'******************************************************************************
'** Module.........: CmouseGlass
'** Description....: Une fonction pour effet Glass et Normal
'**                  On utilise comme ceci
'**                  Private Sub YourFunction()
' *
'**                 Dim cHourglass    As cMouseGlass
'**                 Set cHourglass = New cMouseGlass
'**
'**                 ' .... Execution de codes routines
'**                 ' .... A la fin de la routine le curseur Souris
'**                 ' .... revient automatiquement a ce qu'il etait
'**
'**                 End Sub
'**
'**
'** Author, date...: Yves Lessard , 31-Aug-1999.
'**
'** Version........: 1.0.0.A
'** Modifications..:
'** 1.0.1 - xxx (), dd-mmm-yyyy
'**
'**
'******************************************************************************

Private nOldMousPointer    As Long

Private Sub Class_Initialize()
'******************************************************************************
'** SubRoutine.....: Class_Initialize
'**
'** Description....: Sauve le curseur et active le Sablier
'**
'** Main Blocks....:
'**
'** Author, date...: Yves Lessard  (WAI), 31-Aug-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**
'******************************************************************************
   nOldMousPointer = Screen.MousePointer
   Screen.MousePointer = vbHourglass

End Sub

Private Sub Class_Terminate()
'******************************************************************************
'** SubRoutine.....: Class_Terminate
'**
'** Description....: Replace le curseur a ce qu'il etait
'**
'** Main Blocks....:
'**
'** Author, date...: Yves Lessard  (WAI), 31-Aug-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**
'******************************************************************************
   
   Screen.MousePointer = nOldMousPointer

End Sub

