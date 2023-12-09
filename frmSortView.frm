VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSortView 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort Edit / View"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6135
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmSortView.frx":0000
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "frmSortView.frx":0014
         TabIndex        =   6
         Top             =   120
         Width           =   5895
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtSortNo 
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
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cboItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Sort No"
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Item"
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
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sort_Quality"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmSortView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As DAO.Recordset
Private Sub cboItem_Change()
Data1.RecordSource = "select * from Sort_Quality where ItemCode=" & Me.cboItem.ItemData(Me.cboItem.ListIndex) & ""
Data1.Refresh
Me.cboItem.ToolTipText = Me.cboItem.ItemData(Me.cboItem.ListIndex)
End Sub
Private Sub cboItem_Click()
cboItem_Change
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Data1.databasename = App.Path & "\cuts.mdb"
Set REC1 = Db.OpenRecordset("select * from ItemMaster")
While Not REC1.EOF
Me.cboItem.AddItem (REC1("Items"))
Me.cboItem.ItemData(Me.cboItem.NewIndex) = REC1("SlNo")
REC1.MoveNext
Wend
If Me.cboItem.ListCount > 0 Then
Me.cboItem.ListIndex = 0
End If
End Sub

Private Sub txtsortno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Data1.RecordSource = "select * from Sort_Quality where SortNo='" & Me.txtSortNo.Text & "'"
Data1.Refresh
End If
End Sub
