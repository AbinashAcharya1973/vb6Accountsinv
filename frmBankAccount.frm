VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmBankAccount 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Account"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   13275
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BankAccount"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmBankAccount.frx":0000
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "frmBankAccount.frx":0014
         TabIndex        =   5
         Top             =   120
         Width           =   12855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox cboBank 
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
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Bank "
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
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBankAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset

Private Sub cboBank_Change()
Data1.RecordSource = "select * from BankAccount where BankCode='" & Me.cboBank.ItemData(Me.cboBank.ListIndex) & "'"
Data1.Refresh
Set rec1 = Db.OpenRecordset("select * from BankMaster where Slno=" & Me.cboBank.ItemData(Me.cboBank.ListIndex))
If Not rec1.EOF Then
Me.txtAddress.Text = rec1("Address")
End If
End Sub

Private Sub cboBank_Click()
cboBank_Change
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Set rec1 = Db.OpenRecordset("select * from BankMaster")
While Not rec1.EOF
Me.cboBank.AddItem (rec1("Bank"))
Me.cboBank.ItemData(Me.cboBank.NewIndex) = rec1("Slno")
rec1.MoveNext
Wend
If Me.cboBank.ListCount > 0 Then
Me.cboBank.ListIndex = 0
End If
End Sub
