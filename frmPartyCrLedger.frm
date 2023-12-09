VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPartyCrLedger 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sundry Crediter Ledger"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   13305
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PartyCrLedger"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmPartyCrLedger.frx":0000
         Height          =   5175
         Left            =   120
         OleObjectBlob   =   "frmPartyCrLedger.frx":0014
         TabIndex        =   8
         Top             =   240
         Width           =   12855
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
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
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   3975
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
         Left            =   4320
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtPartyCode 
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
         Left            =   9000
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
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
         Left            =   3480
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Code"
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
         Left            =   8400
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPartyCrLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As DAO.Recordset
Private Sub cboParty_Change()
Set REC1 = Db.OpenRecordset("select * from Partycr where slno=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
If Not REC1.EOF Then
Me.txtAddress.Text = REC1("Address")
Me.txtPartyCode.Text = Me.cboParty.ItemData(Me.cboParty.ListIndex)
End If
Data1.RecordSource = "select * from PartyCrLedger where Pcode='" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & "'"
Data1.Refresh
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Set REC1 = Db.OpenRecordset("select * from PartyCr")
While Not REC1.EOF
Me.cboParty.AddItem (REC1("party"))
Me.cboParty.ItemData(Me.cboParty.NewIndex) = REC1("Slno")
REC1.MoveNext
Wend
If Me.cboParty.ListCount > 0 Then
Me.cboParty.ListIndex = 0
End If
End Sub

