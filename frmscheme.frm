VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmscheme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheme"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8880
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "scheme"
      Top             =   1350
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmscheme.frx":0000
      Height          =   2040
      Left            =   150
      OleObjectBlob   =   "frmscheme.frx":0014
      TabIndex        =   7
      Top             =   600
      Width           =   8565
   End
   Begin VB.TextBox txtleakage 
      Height          =   315
      Left            =   7725
      TabIndex        =   3
      Text            =   "0"
      Top             =   150
      Width           =   1065
   End
   Begin VB.TextBox txtscheme 
      Height          =   315
      Left            =   5850
      TabIndex        =   2
      Text            =   "0"
      Top             =   150
      Width           =   915
   End
   Begin VB.TextBox txtqty 
      Height          =   315
      Left            =   3975
      TabIndex        =   1
      Text            =   "0"
      Top             =   150
      Width           =   915
   End
   Begin VB.ComboBox cboitemname 
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
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   3240
   End
   Begin VB.Label Label3 
      Caption         =   "Leakage"
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
      Left            =   6900
      TabIndex        =   6
      Top             =   150
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Scheme"
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
      Left            =   5025
      TabIndex        =   5
      Top             =   150
      Width           =   840
   End
   Begin VB.Label Label1 
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
      Height          =   315
      Left            =   3450
      TabIndex        =   4
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "frmscheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As DAO.Recordset

Private Sub cboitemname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtqty.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Data1.databasename = db.Name
Set rec = db.OpenRecordset("select * from stock")
    While Not rec.EOF
        Me.cboitemname.AddItem rec("itemname")
        Me.cboitemname.ItemData(Me.cboitemname.NewIndex) = rec("productcode")
        rec.MoveNext
    Wend
    If Me.cboitemname.ListCount > 0 Then
        Me.cboitemname.ListIndex = 0
    End If
End Sub

Private Sub txtleakage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    db.Execute "insert into scheme values('" & Me.cboitemname.Text & "'," & Me.txtqty.Text & "," & Me.txtscheme.Text & "," & Me.txtleakage.Text & "," & Me.cboitemname.ItemData(Me.cboitemname.ListIndex) & ")"
    Me.txtqty.Text = "0"
    Me.txtscheme.Text = "0"
    Me.txtleakage.Text = "0"
    Me.cboitemname.SetFocus
    Me.Data1.Refresh
    Me.DBGrid1.Refresh
End If
End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtscheme.SetFocus
End If
End Sub

Private Sub txtscheme_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtleakage.SetFocus
End If
End Sub
