VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPartyDrEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Party Dr Edit"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmPartyDrEdit.frx":0000
   ScaleHeight     =   5520
   ScaleWidth      =   11400
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PartyDr"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11175
      Begin VB.TextBox TxtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   120
         Width           =   4935
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmPartyDrEdit.frx":D4E3
         Height          =   4575
         Left            =   120
         OleObjectBlob   =   "FrmPartyDrEdit.frx":D4F7
         TabIndex        =   2
         Top             =   600
         Width           =   10935
      End
      Begin VB.Label Label1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmPartyDrEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp_accid
Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    ans = MsgBox("Update This?", vbYesNo)
    If ans = 6 Then
        If ColIndex = 0 Then
            db.Execute ("update ledgermaster set AccName='" & Trim(Me.DBGrid1.Columns(0)) & "' where AccId=" & Me.DBGrid1.Columns(19))
        End If
        If ColIndex = 1 Or ColIndex = 2 Then
            db.Execute ("update ledgermaster set ADDRESS1='" & Trim(Me.DBGrid1.Columns(1)) & "',ADDRESS2='" & Trim(Me.DBGrid1.Columns(18)) & "' where AccId=" & Me.DBGrid1.Columns(19))
        End If
    End If
End Sub

Private Sub DBGrid1_AfterDelete()
db.Execute ("delete * from ledgermaster where accid=" & temp_accid)
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
temp_accid = DBGrid1.Columns(19)
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        Me.DBGrid1.AllowUpdate = True
    End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtsearch_Change()
    Data1.RecordSource = "select * from partyDr where party like '" & Me.TxtSearch.Text & "*'"
    Data1.Refresh
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

    End If
End Sub
