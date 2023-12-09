VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmUnitcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Code Entry"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5775
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "UnitCode"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5535
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Form1.frx":0000
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "Form1.frx":0014
         TabIndex        =   8
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox cboUnit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtUnitQty 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtUnitName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Unit Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmUnitcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Db.Execute ("insert into UnitCode(UnitName,Qty,UnitType) values('" & Me.txtUnitName.Text & "'," & Me.txtUnitQty.Text & ",'" & Me.cboUnit.Text & "')")
Data1.Refresh
Me.txtUnitName.Text = ""
Me.txtUnitQty.Text = ""
Me.txtUnitName.SetFocus

End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.cboUnit.AddItem ("METER")
Me.cboUnit.AddItem ("C.M")
Me.cboUnit.AddItem ("KG")
Me.cboUnit.AddItem ("GM")
Me.cboUnit.AddItem ("Number")
Me.cboUnit.ListIndex = 0
Data1.databasename = App.Path & "\cuts.mdb"
End Sub

Private Sub txtUnitName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.txtUnitName.Text <> "" Then
Me.txtUnitQty.SetFocus
End If
End If
End Sub

Private Sub txtUnitQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.txtUnitQty.Text <> "" Then
Me.cboUnit.SetFocus
End If
End If
End Sub
