VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPartyCrEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Party Creditor Edit"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   13350
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
      Left            =   1320
      TabIndex        =   2
      Top             =   180
      Width           =   4935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PartyCr"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   13155
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmPartyCrEdit.frx":0000
         Height          =   4635
         Left            =   120
         OleObjectBlob   =   "frmPartyCrEdit.frx":0014
         TabIndex        =   1
         Top             =   120
         Width           =   12915
      End
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
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmPartyCrEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp_accid

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 0 Then
    db.Execute "Update ledgermaster set Accname='" & Me.DBGrid1.Columns(ColIndex) & "' where accid=" & Me.DBGrid1.Columns(6)
End If
If ColIndex = 1 Then
    db.Execute "Update ledgermaster set Address1='" & Me.DBGrid1.Columns(ColIndex) & "' where accid=" & Me.DBGrid1.Columns(6)
End If
If ColIndex = 2 Then
    db.Execute "Update ledgermaster set Address2='" & Me.DBGrid1.Columns(ColIndex) & "' where accid=" & Me.DBGrid1.Columns(6)
End If
If ColIndex = 7 Then
    db.Execute "Update ledgermaster set statecode='" & Me.DBGrid1.Columns(ColIndex) & "' where accid=" & Me.DBGrid1.Columns(6)
End If

End Sub

Private Sub DBGrid1_AfterDelete()
db.Execute ("delete * from ledgermaster where accid=" & temp_accid)
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
ans = MsgBox("Do you want to Delete the Record?", vbYesNo)
If ans = 6 Then
    temp_accid = DBGrid1.Columns(6)
Else
    Cancel = False
End If
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
    Data1.RecordSource = "select * from partyCr where party like '" & Me.txtsearch.Text & "*'"
    Data1.Refresh
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
