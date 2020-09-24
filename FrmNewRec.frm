VERSION 5.00
Begin VB.Form FrmNewRec 
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPhone 
      Height          =   285
      Left            =   1200
      MaxLength       =   13
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox TxtLname 
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox TxtFname 
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Phone"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Last Name"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fisrt Name"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   750
   End
End
Attribute VB_Name = "FrmNewRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MaClass As CadoGeneric
Private RecSet As ADODB.Recordset

Private Sub Command1_Click()
Set MaClass = New CadoGeneric
Set RecSet = New ADODB.Recordset
Dim szQuery As String

'** First must perform a Dummy query
szQuery = "SELECT FName, LName, Phone From TblClient" & _
          " WHERE ClientID = -1"
With MaClass
    .Dbname = App.Path & "\Client.MDB"
    .SetConnectType ACCESS
End With

If MaClass.GetRecord(szQuery, RecSet) Then
    '** We can now save the record
    RecSet.AddNew
    RecSet.Fields("FName").Value = TxtFname.Text
    RecSet.Fields("LName").Value = TxtLname.Text
    RecSet.Fields("Phone").Value = TxtPhone.Text
    If MaClass.SaveRecord(RecSet) Then
        MsgBox "Record Saved"
        RecSet.Close
    Else
        MsgBox MaClass.ErrorInfo
    End If
Else
    MsgBox "error"
End If

Set RecSet = Nothing
Set MaClass = Nothing

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmNewRec = Nothing
End Sub
