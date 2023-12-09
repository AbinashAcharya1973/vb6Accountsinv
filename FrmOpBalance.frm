VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmOpBalance 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening Balance Entry"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmOpBalance.frx":0000
   ScaleHeight     =   8205
   ScaleWidth      =   11460
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
      RecordSource    =   "LedgerMaster"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   11175
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11175
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmOpBalance.frx":D4E3
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "FrmOpBalance.frx":D4F7
         TabIndex        =   8
         Top             =   120
         Width           =   10935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.TextBox TxtSearch 
         Height          =   285
         Left            =   7080
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox CboGroup 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Groups"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmOpBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset, temp_accid
Private Sub cboGroup_Change()
    Data1.RecordSource = "select * from ledgermaster where GroupId=" & Me.cbogroup.ItemData(Me.cbogroup.ListIndex) & " order by AccName Asc"
    Data1.Refresh
    Set rec1 = db.OpenRecordset("select sum(obalance) as totalamount from ledgermaster where groupid=" & Me.cbogroup.ItemData(Me.cbogroup.ListIndex))
    If Not IsNull(rec1!totalamount) Then
        Me.lbltotal.Caption = Format(rec1!totalamount, "###########0.00")
    Else
        Me.lbltotal.Caption = "0.00"
    End If
End Sub
Private Sub cboGroup_Click()
    cboGroup_Change
End Sub

Private Sub DBGrid1_AfterDelete()
db.Execute ("delete * from partydr where accid=" & temp_accid)
db.Execute ("delete * from partycr where accid=" & temp_accid)
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    db.Execute ("update ledgertran set groupid=" & Me.DBGrid1.Columns(2) & " where accid=" & Me.DBGrid1.Columns(0))
    Set rec1 = db.OpenRecordset("select * from groups where groupid=" & Me.DBGrid1.Columns(2))
    If Not rec1.EOF Then
        db.Execute ("update ledgermaster set transactiontype='" & rec1("Groupnature") & "',groupname='" & rec1("groupname") & "' where accid=" & Me.DBGrid1.Columns(0))
    End If
    'Data1.Refresh
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
temp_accid = DBGrid1.Columns(0)
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Data1.databasename = dbname
    Set rec1 = db.OpenRecordset("select * from Groups")
    While Not rec1.EOF
        Me.cbogroup.AddItem (rec1("GroupName"))
        Me.cbogroup.ItemData(Me.cbogroup.NewIndex) = rec1("GroupId")
        rec1.MoveNext
    Wend
    If Me.cbogroup.ListCount > 0 Then
        Me.cbogroup.ListIndex = 0
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Label3_Click()

End Sub

Private Sub txtsearch_Change()
Me.Data1.RecordSource = "select * from LedgerMAster where AccName like '" & Me.TxtSearch.Text & "*'"
Me.Data1.Refresh
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Data1.RecordSource = "select * from ledgermaster where accname like '" & Me.TxtSearch.Text & "*'"
        Data1.Refresh
    End If
End Sub
