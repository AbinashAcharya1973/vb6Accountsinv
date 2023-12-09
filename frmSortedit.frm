VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSortedit 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort No Edit"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   10620
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   10335
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmSortedit.frx":0000
         Height          =   4095
         Left            =   120
         OleObjectBlob   =   "frmSortedit.frx":0014
         TabIndex        =   4
         Top             =   120
         Width           =   10095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.TextBox txtsortno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Sort No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SortMaster"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmSortedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As DAO.Recordset, REC2 As Recordset
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Data1.databasename = App.Path & "\cuts.mdb"
End Sub

Private Sub txtsortno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

End If
End Sub
