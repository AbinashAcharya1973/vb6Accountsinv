VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPartyDrLedger 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sundry Debtor Ledger"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13305
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PartyDrLedger"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmPartyDrLedger.frx":0000
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "frmPartyDrLedger.frx":0014
         TabIndex        =   6
         Top             =   120
         Width           =   12855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin VB.TextBox txtPcode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8400
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   480
         Width           =   4935
      End
      Begin VB.ComboBox cboParty 
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
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Party"
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
         Left            =   2760
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPartyDrLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset

Private Sub cboParty_Change()
Set REC1 = Db.OpenRecordset("select * from PartyDr where Slno=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
If Not REC1.EOF Then
Me.txtAddress.Text = REC1("Address")
Me.txtPcode.Text = Me.cboParty.ItemData(Me.cboParty.ListIndex)
End If
Data1.RecordSource = "select * from PartyDrLedger where Pcode='" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & "'"
Data1.Refresh
End Sub

Private Sub cboparty_Click()
cboParty_Change
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Set REC1 = Db.OpenRecordset("select * from PartyDr")
While Not REC1.EOF
Me.cboParty.AddItem (REC1("Party"))
Me.cboParty.ItemData(Me.cboParty.NewIndex) = REC1("Slno")
REC1.MoveNext
Wend
If Me.cboParty.ListCount > 0 Then
Me.cboParty.ListIndex = 0
End If

End Sub
