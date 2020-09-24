VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Add New Record"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fill Combo"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox CboCities 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save as Batch"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill the Grid (Batch)"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   4080
      Width           =   495
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MaClass As CadoGeneric
Private MyRecset As ADODB.Recordset

Private Sub CboCities_Click()
    Label1 = CboCities.ItemData(CboCities.ListIndex)
End Sub

Private Sub Command1_Click()
Dim szQuery As String

szQuery = "SELECT FName, LName, Phone FROM TblClient"
With MaClass
    If .GetBatchRecord(szQuery, MyRecset) Then
        Set DataGrid1.DataSource = MyRecset
        DataGrid1.Columns(0).Width = 1100
        DataGrid1.Columns(1).Width = 1500
        DataGrid1.Columns(2).Width = 1350
    Else
        MsgBox .ErrorInfo & "  " & .ErrorNumber
    End If
End With


End Sub

Private Sub Command2_Click()
    If MaClass.SaveBatchRecord(MyRecset) Then
        '** Success
    Else
        MsgBox "Error save Batch"
    End If
    
End Sub

Private Sub Command3_Click()
    FillCitiesCombo CboCities
    
End Sub

Private Sub Command4_Click()
    FrmNewRec.Show vbModal
    
End Sub

Private Sub Form_Load()
Set MaClass = New CadoGeneric
Set MyRecset = New ADODB.Recordset
With MaClass
    .Dbname = App.Path & "\Client.MDB"
    .SetConnectType ACCESS
End With

End Sub
