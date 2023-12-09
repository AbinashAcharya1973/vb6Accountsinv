VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewJournal 
   BackColor       =   &H00E3E3E3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Journal"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   8610
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
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
      Left            =   5678
      TabIndex        =   30
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "DELETE"
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
      Left            =   4438
      TabIndex        =   23
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "EDIT"
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
      Left            =   3198
      TabIndex        =   22
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
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
      Left            =   1958
      TabIndex        =   5
      Top             =   7200
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E3E3E3&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   8415
      Begin VB.TextBox txtNarration 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Width           =   8175
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmNewJournal.frx":0000
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmNewJournal.frx":0014
         TabIndex        =   20
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   8415
      Begin VB.ComboBox CboChGroup 
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
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtChildAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6720
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cboChildAccount 
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
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtChildTran 
         Appearance      =   0  'Flat
         BackColor       =   &H00E3E3E3&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label LblChAddress 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   1440
         Width           =   6735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Child Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Child Tran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox CboMainGroup 
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
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   480
         Width           =   3855
      End
      Begin MSMask.MaskEdBox txtJDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtSlNo 
         Appearance      =   0  'Flat
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
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtMainAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox cboMainTran 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox cboMainAccount 
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
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label LblMainAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label Label10 
         Caption         =   "Group"
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
         TabIndex        =   27
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
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
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Sl No."
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
         Left            =   5760
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
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
         Left            =   5760
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Main Tran "
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
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Main Account"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1815
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
      RecordSource    =   "Temp_Journal"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmNewJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset, rec As DAO.Recordset, rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset, BEFORE_AMOUNT, JDELETE, rec5 As Recordset
Attribute rec.VB_VarUserMemId = 1073938432
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432
Attribute BEFORE_AMOUNT.VB_VarUserMemId = 1073938432
Attribute JDELETE.VB_VarUserMemId = 1073938432
Attribute rec5.VB_VarUserMemId = 1073938432
Private Sub Slider1_Click()

End Sub

Private Sub CboChGroup_Change()
    Set rec1 = db.OpenRecordset("select * from ledgermaster where GroupID=" & Me.CboChGroup.ItemData(Me.CboChGroup.ListIndex))
    Me.cboChildAccount.Clear
    While Not rec1.EOF
        Me.cboChildAccount.AddItem (rec1("AccName"))
        Me.cboChildAccount.ItemData(Me.cboChildAccount.NewIndex) = rec1("AccID")
        rec1.MoveNext
    Wend
    If Me.cboChildAccount.ListCount > 0 Then
        Me.cboChildAccount.ListIndex = 0
    End If
End Sub
Private Sub CboChGroup_Click()
    CboChGroup_Change
End Sub
Private Sub CboChGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Me.txtNarration.SetFocus
    End If
    If KeyCode = 13 Then
        Me.cboChildAccount.SetFocus
    End If
End Sub

Private Sub cboChildAccount_Change()
    Set rec1 = db.OpenRecordset("select * from ledgermaster where AccId=" & Me.cboChildAccount.ItemData(Me.cboChildAccount.ListIndex))
    If Not IsNull(rec1!Address1) Then
        Me.LblChAddress.Caption = rec1("Address1")
    Else
        Me.LblChAddress.Caption = ""
    End If

End Sub
Private Sub cboChildAccount_Click()
    cboChildAccount_Change
End Sub
Private Sub cboChildAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtChildAmount.SetFocus
    End If
End Sub

Private Sub cboMainAccount_Change()
'Set rec1 = Db.OpenRecordset("select * from LedgerMaster where AccName not like '" & Me.cboMainAccount.Text & "'")
'Me.cboChildAccount.Clear
'While Not rec1.EOF
'Me.cboChildAccount.AddItem (rec1("AccName"))
'Me.cboChildAccount.ItemData(Me.cboChildAccount.NewIndex) = rec1("AccId")
'rec1.MoveNext
'Wend
'If Me.cboChildAccount.ListCount > 0 Then
'Me.cboChildAccount.ListIndex = 0
'End If
    Set rec1 = db.OpenRecordset("select * from ledgermaster where AccId=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex))
    If Not IsNull(rec1!Address1) Then
        Me.LblMainAddress.Caption = rec1("Address1")
    Else
        Me.LblMainAddress.Caption = ""
    End If

End Sub
Private Sub cboMainAccount_Click()
    cboMainAccount_Change
End Sub
Private Sub cboMainAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboMainTran.SetFocus
    End If
End Sub

Private Sub CboMainGroup_Change()
    Set rec1 = db.OpenRecordset("select * from ledgermaster where GroupID=" & Me.CboMainGroup.ItemData(Me.CboMainGroup.ListIndex))
    Me.cboMainAccount.Clear
    While Not rec1.EOF
        Me.cboMainAccount.AddItem (rec1("AccName"))
        Me.cboMainAccount.ItemData(Me.cboMainAccount.NewIndex) = rec1("AccID")
        rec1.MoveNext
    Wend
    If Me.cboMainAccount.ListCount > 0 Then
        Me.cboMainAccount.ListIndex = 0
    End If

End Sub
Private Sub CboMainGroup_Click()
    CboMainGroup_Change
End Sub
Private Sub CboMainGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboMainAccount.SetFocus
    End If
End Sub

Private Sub cboMainTran_Change()
    If Me.cboMainTran.Text = "DEBIT" Then
        Set rec4 = db.OpenRecordset("select * from temp_journal")
        If Not rec4.EOF Then
            db.Execute ("update temp_journal set TranType='CREDIT'")
            Data1.Refresh
        End If
        Me.txtChildTran.Text = "CREDIT"
    Else
        Set rec4 = db.OpenRecordset("select * from temp_journal")
        If Not rec4.EOF Then
            db.Execute ("update temp_journal set TranType='DEBIT'")
            Data1.Refresh
        End If
        Me.txtChildTran.Text = "DEBIT"
    End If
End Sub
Private Sub cboMainTran_Click()
    cboMainTran_Change
End Sub
Private Sub cboMainTran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboMainTran.Locked = True
        Me.cboMainAccount.Locked = True
        Me.CboChGroup.SetFocus
    End If
End Sub

Private Sub cmddelete_Click()
    db.Execute ("delete * from Temp_Journal")
    Data1.Refresh
    Me.txtMainAmount.Text = "0.00"
    Me.txtSlNo.Locked = False
    JDELETE = "y"
    Me.txtJDate.SetFocus
End Sub
Private Sub cmdedit_Click()
    db.Execute ("delete * from tempjournal")
    Data1.Refresh
    Me.txtMainAmount.Text = "0.00"
    Me.txtSlNo.Locked = False
    Me.txtJDate.SetFocus
End Sub

Private Sub cmdprint_Click()
frmprintjournal.Show 0
End Sub

Private Sub cmdsave_Click()
'temp_day = Left((Me.txtJDate.Text), 2)
'temp_month = Mid((Me.txtJDate.Text), 4, 2)
'temp_year = Right((Me.txtJDate.Text), 4)
'
'Accperiod_day = Left(AccountingPeriod, 2)
'Accperiod_month = Mid(AccountingPeriod, 4, 2)
'Accperiod_year = Right(AccountingPeriod, 4)
'
'ans = MsgBox("Save This?", vbYesNo)
'If ans = 6 Then
'Set rec = Db.OpenRecordset("select * from Temp_Journal")
'If Not rec.EOF Then
''------Check And Delete Previous Entry-------------
'Set rec1 = Db.OpenRecordset("select * from JournalHead where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'If Not rec1.EOF Then
'    Set rec3 = Db.OpenRecordset("select * from LedgerMAster where AccId=" & rec1("AccId"))
'    If Not rec3.EOF Then
'        Set rec4 = Db.OpenRecordset("select * from ledgertran where AccId=" & rec1("AccId") & " and VoucherType='Journal' and VoucherSlno=" & Me.txtSlNo.Text & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'        If Not rec4.EOF Then
'            If rec1("ParentTrn") = "CREDIT" Then
'            tempsign = rec3("Dr")
'            Else
'            tempsign = rec3("Cr")
'            End If
'            Db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance" & tempsign & rec1("Amount") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec4("SlNo"))
'            Db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Journal' and VoucherSlno=" & Me.txtSlNo.Text & " and TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'        End If
'    End If
'Db.Execute ("delete * from JournalHead where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'End If
'    '=============New Entry===========
'    Set rec3 = Db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex))
'    If Not rec3.EOF Then
'        Set rec4 = Db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & " and SlNo=(select max(slno) from LedgerTran where AccID=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ")")
'            If Not rec4.EOF Then
'            mainLedger_Balance = rec4("Balance")
'            mainledgerslno = rec4("SlNo") + 1
'            Else
'            mainLedger_Balance = 0
'            mainledgerslno = 1
'            End If
'            If Me.cboMainTran.Text = "CREDIT" Then
'            tempsign = rec3("Cr")
'            Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & mainledgerslno & ",'" & Me.txtJDate.Text & "','By Sundries',0," & Me.txtMainAmount.Text & "," & mainLedger_Balance & tempsign & Val(Me.txtMainAmount.Text) & "," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ",'" & Trim(Me.TxtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec3("GroupId") & ")")
'            Else
'            tempsign = rec3("Dr")
'            Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & mainledgerslno & ",'" & Me.txtJDate.Text & "','By Sundries'," & Me.txtMainAmount.Text & ",0," & mainLedger_Balance & tempsign & Me.txtMainAmount.Text & "," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ",'" & Trim(Me.TxtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec3("GroupId") & ")")
'            End If
'    End If
'    Db.Execute ("insert into JournalHead (SlNo,JDate,AccId,AccName,ParentTrn,Amount,ChildTran,Narration) values(" & Me.txtSlNo.Text & ",'" & Me.txtJDate.Text & "'," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ",'" & Me.cboMainAccount.Text & "','" & Me.cboMainTran.Text & "'," & Me.txtMainAmount.Text & ",'" & Me.txtChildTran.Text & "','" & Trim(Me.TxtNarration.Text) & "')")
'
'
''------------Check And Delete Child Transaction----------------
'Set rec1 = Db.OpenRecordset("select * from JournalDetails where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'If Not rec1.EOF Then
'While Not rec1.EOF
'    Set rec2 = Db.OpenRecordset("select * from LedgerMaster where AccId=" & rec1("AccId"))
'    If Not rec2.EOF Then
'        Set rec4 = Db.OpenRecordset("select * from LedgerTran where AccId=" & rec1("AccId") & " and VoucherType='Journal' and VoucherSlno=" & Me.txtSlNo.Text & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'        If Not rec4.EOF Then
'            If rec1("TranType") = "CREDIT" Then
'            tempsign = rec2("Dr")
'            Else
'            tempsign = rec2("Cr")
'            End If
'            Db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance" & tempsign & rec1("Amount") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec4("SlNo"))
'            Db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Journal'  and VoucherSlno=" & Me.txtSlNo.Text & " and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'        End If
'    End If
'rec1.MoveNext
'Wend
'Db.Execute ("delete * from JournalDetails where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
'End If
''---------------New Child Transaction Entry-----------------
'Set rec1 = Db.OpenRecordset("select * from Temp_Journal")
'If Not rec1.EOF Then
'While Not rec1.EOF
'    Db.Execute ("insert into JournalDetails (SlNo,JDate,AccId,AccName,Amount,TranType) values(" & Me.txtSlNo.Text & ",'" & Me.txtJDate.Text & "'," & rec1("AccId") & ",'" & rec1("AccName") & "'," & rec1("Amount") & ",'" & rec1("TranType") & "')")
'    Set rec2 = Db.OpenRecordset("select * from LedgerMaster where AccId=" & rec1("AccId"))
'    If Not rec2.EOF Then
'        Set rec3 = Db.OpenRecordset("select * from LedgerTran where AccId=" & rec1("AccId") & " and SlNo=(select Max(SlNo) from LedgerTran where AccId=" & rec1("AccId") & ")")
'        If Not rec3.EOF Then
'        Ledger_Balance = rec3("Balance")
'        Ledger_slno = rec3("SlNo") + 1
'        Else
'        Ledger_Balance = 0
'        Ledger_slno = 1
'        End If
'            If rec1("TranType") = "CREDIT" Then
'            tempsign = rec2("Cr")
'            Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & Ledger_slno & ",'" & Me.txtJDate.Text & "','By " & Me.cboMainAccount.Text & "',0," & rec1("Amount") & "," & Ledger_Balance & tempsign & rec1("Amount") & "," & rec1("AccId") & ",'" & Trim(Me.TxtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec2("GroupId") & ")")
'            End If
'            If rec1("TranType") = "DEBIT" Then
'            tempsign = rec2("Dr")
'            Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & Ledger_slno & ",'" & Me.txtJDate.Text & "','By " & Me.cboMainAccount.Text & "'," & rec1("Amount") & ",0," & Ledger_Balance & tempsign & rec1("Amount") & "," & rec1("AccId") & ",'" & Trim(Me.TxtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec2("GroupId") & ")")
'            End If
'    End If
'rec1.MoveNext
'Wend
'End If
'
'Db.Execute ("delete * from temp_journal")
'Data1.Refresh
'Me.cboMainAccount.Locked = False
'Me.cboMainTran.Locked = False
'Me.txtSlNo.Locked = True
'Me.TxtNarration.Text = ""
'mainLedger_Balance = 0
'mainledgerslno = 0
'Ledger_Balance = 0
'Ledger_slno = 0
'Me.txtMainAmount.Text = "0.00"
'Me.txtJDate.SetFocus
'End If
'End If
'-----Change------------

    temp_day = Left((Me.txtJDate.Text), 2)
    temp_month = Mid((Me.txtJDate.Text), 4, 2)
    temp_year = Right((Me.txtJDate.Text), 4)
    temp_from = temp_month & "/" & temp_day & "/" & temp_year

    Accperiod_day = Left(AccountingPeriod, 2)
    Accperiod_month = Mid(AccountingPeriod, 4, 2)
    Accperiod_year = Right(AccountingPeriod, 4)
    AccPeriod = Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year
    ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then
        Set rs = db.OpenRecordset("select * from Temp_Journal")
        If Not rs.EOF Then
            Set rec = db.OpenRecordset("select * from JournalDetails where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_from & "#")
            If Not rec.EOF Then
                Set rec5 = db.OpenRecordset("select * from JournalHead where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_from & "#")
                MAccId = rec5("AccId")
                If Not rec5.EOF Then
                End If
                While Not rec.EOF
                    Set rec1 = db.OpenRecordset("select * from LedgertRan where AccId=" & MAccId & " and VoucherType='Journal' and Tdate=#" & temp_from & "# and TranAccId=" & rec("AccId") & " and VoucherSlno=" & Me.txtSlNo.Text)
                    If Not rec1.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec1("AccId") & " and SlNo>=" & rec1("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Journal' and VoucherSlno=" & Me.txtSlNo.Text & " and TDate =#" & temp_from & "# and TranAccId=" & rec("AccId"))
                    End If
                    Set rec1 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec("AccId") & " and Tdate=#" & temp_from & "# and VoucherType='Journal' and VoucherslNo=" & Me.txtSlNo.Text & " and TranAccId=" & MAccId)
                    If Not rec1.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec1("AccId") & " and SlNo>=" & rec1("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Journal' and VoucherSlno=" & Me.txtSlNo.Text & " and TDate =#" & temp_from & "# and TranAccId=" & MAccId)
                    End If
                    rec.MoveNext
                Wend
                db.Execute ("delete * from JournalHead where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_from & "#")
                db.Execute ("delete * from JournalDetails where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_from & "#")
            End If

            '=============New Entry===========
            Set rec = db.OpenRecordset("select * from Temp_Journal")
            If Not rec.EOF Then
                db.Execute ("insert into JournalHead (SlNo,JDate,AccId,AccName,ParentTrn,Amount,ChildTran,Narration) values(" & Me.txtSlNo.Text & ",'" & Me.txtJDate.Text & "'," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ",'" & Me.cboMainAccount.Text & "','" & Me.cboMainTran.Text & "'," & Me.txtMainAmount.Text & ",'" & Me.txtChildTran.Text & "','" & Trim(Me.txtNarration.Text) & "')")
                While Not rec.EOF
                    db.Execute ("insert into JournalDetails (SlNo,JDate,AccId,AccName,Amount,TranType) values(" & Me.txtSlNo.Text & ",'" & Me.txtJDate.Text & "'," & rec("AccId") & ",'" & rec("AccName") & "'," & rec("Amount") & ",'" & rec("TranType") & "')")
                    Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex))
                    If Not rec3.EOF Then
                        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & " and SlNo=(select max(slno) from LedgerTran where AccID=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ")")
                        If Not rec4.EOF Then
                            mainledgerslno = rec4("SlNo") + 1
                        Else
                            mainledgerslno = 1
                        End If
                        If Me.cboMainTran.Text = "CREDIT" Then
                            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & mainledgerslno & ",'" & Me.txtJDate.Text & "','By " & rec("AccName") & "',0," & rec("Amount") & ",0," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ",'" & Trim(Me.txtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec3("GroupId") & "," & rec("AccId") & ")")
                        Else
                            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & mainledgerslno & ",'" & Me.txtJDate.Text & "','To " & rec("AccName") & "'," & rec("Amount") & ",0,0," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ",'" & Trim(Me.txtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec3("GroupId") & "," & rec("AccId") & ")")
                        End If
                    End If
                    Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccId=" & rec("AccId"))
                    If Not rec3.EOF Then
                        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec("AccId") & " and SlNo=(select max(slno) from LedgerTran where AccID=" & rec("AccId") & ")")
                        If Not rec4.EOF Then
                            Ledger_slno = rec4("SlNo") + 1
                        Else
                            Ledger_slno = 1
                        End If
                        If rec("TranType") = "CREDIT" Then
                            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Ledger_slno & ",'" & Me.txtJDate.Text & "','By " & Me.cboMainAccount.Text & "',0," & rec("Amount") & ",0," & rec("AccId") & ",'" & Trim(Me.txtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec3("GroupId") & "," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ")")
                        End If
                        If rec("TranType") = "DEBIT" Then
                            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Ledger_slno & ",'" & Me.txtJDate.Text & "','By " & Me.cboMainAccount.Text & "'," & rec("Amount") & ",0,0," & rec("AccId") & ",'" & Trim(Me.txtNarration.Text) & "','Journal'," & Me.txtSlNo.Text & "," & rec3("GroupId") & "," & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & ")")
                        End If
                    End If
                    rec.MoveNext
                Wend
            End If

            db.Execute ("delete * from temp_journal")
            Data1.Refresh
            Me.cboMainAccount.Locked = False
            Me.cboMainTran.Locked = False
            Me.txtSlNo.Locked = True
            Me.txtNarration.Text = ""
            mainLedger_Balance = 0
            mainledgerslno = 0
            Ledger_Balance = 0
            Ledger_slno = 0
            Me.txtMainAmount.Text = "0.00"
            Me.txtJDate.SetFocus
        End If
    End If

End Sub

Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    diffamount = BEFORE_AMOUNT - Val(Me.DBGrid1.Columns(2))
    Me.txtMainAmount.Text = Format(Val(Me.txtMainAmount.Text) - diffamount, "###############0.00")
End Sub
Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    BEFORE_AMOUNT = Val(Me.DBGrid1.Columns(2))
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Me.txtMainAmount.Text = Format(Me.txtMainAmount.Text - Val(Me.DBGrid1.Columns(2)), "#############0.00")
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Me.txtJDate.Text = Format(Date, "DD/MM/YYYY")
    Set rec1 = db.OpenRecordset("select * from Groups")
    While Not rec1.EOF
        Me.CboMainGroup.AddItem (rec1("GroupName"))
        Me.CboChGroup.AddItem (rec1("GroupName"))
        Me.CboMainGroup.ItemData(Me.CboMainGroup.NewIndex) = rec1("GroupID")
        Me.CboChGroup.ItemData(Me.CboChGroup.NewIndex) = rec1("GroupID")
        rec1.MoveNext
    Wend
    If Me.CboMainGroup.ListCount > 0 Then
        Me.CboMainGroup.ListIndex = 0
        Me.CboChGroup.ListIndex = 0
    End If
    Set rec1 = db.OpenRecordset("select max(SlNo) as max_slno from JournalHead where Jdate=#" & Format(Date, "MM") & "/" & Format(Date, "DD") & "/" & Format(Date, "YYyy") & "#")
    If Not IsNull(rec1!max_slno) Then
        temp_slno = rec1!max_slno + 1
    Else
        temp_slno = 1
    End If
    Me.txtSlNo.Text = temp_slno
    'Set rec1 = Db.OpenRecordset("select *from LedgerMaster")
    'While Not rec1.EOF
    'Me.cboMainAccount.AddItem (rec1("accname"))
    'Me.cboMainAccount.ItemData(Me.cboMainAccount.NewIndex) = rec1("AccId")
    'rec1.MoveNext
    'Wend
    'If Me.cboMainAccount.ListCount > 0 Then
    'Me.cboMainAccount.ListIndex = 0
    'End If
    Me.cboMainTran.AddItem ("DEBIT")
    Me.cboMainTran.AddItem ("CREDIT")
    Me.cboMainTran.ListIndex = 0
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("delete * from Temp_journal")
End Sub

Private Sub txtChildAmount_GotFocus()
    Me.txtChildAmount.SelStart = 0
    Me.txtChildAmount.SelLength = Len(Me.txtChildAmount.Text)
End Sub
Private Sub txtChildAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Val(Me.txtChildAmount.Text) > 0 Then
            If Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) <> Me.cboChildAccount.ItemData(Me.cboChildAccount.ListIndex) Then
                Set rec1 = db.OpenRecordset("select * from Temp_journal where AccId=" & Me.cboMainAccount.ItemData(Me.cboMainAccount.ListIndex) & " or AccId=" & Me.cboChildAccount.ItemData(Me.cboChildAccount.ListIndex))
                If rec1.EOF Then
                    Me.txtMainAmount.Text = Format(Val(Me.txtMainAmount.Text) + Val(Me.txtChildAmount.Text), "############0.00")
                    db.Execute ("insert into temp_journal (AccId,AccName,Amount,TranType) values(" & Me.cboChildAccount.ItemData(Me.cboChildAccount.ListIndex) & ",'" & Me.cboChildAccount.Text & "'," & Me.txtChildAmount.Text & ",'" & Me.txtChildTran.Text & "')")
                    Data1.Refresh
                    Me.txtChildAmount.Text = "0.00"
                    Me.cboChildAccount.ListIndex = 0
                    Me.cboChildAccount.SetFocus
                Else
                    MsgBox "Allready Exists", vbCritical
                End If
            End If
        Else
            MsgBox "Zero Not Enter", vbCritical
        End If
        Me.CboChGroup.SetFocus
    End If

End Sub

Private Sub txtJDate_GotFocus()
    Me.txtJDate.SelStart = 0
    Me.txtJDate.SelLength = Len(Me.txtJDate.Text)
End Sub
Private Sub txtJDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.txtJDate.Text), 2)
        temp_month = Mid((Me.txtJDate.Text), 4, 2)
        temp_year = Right((Me.txtJDate.Text), 4)

        Accperiod_day = Left(AccountingPeriod, 2)
        Accperiod_month = Mid(AccountingPeriod, 4, 2)
        Accperiod_year = Right(AccountingPeriod, 4)
        Set rec1 = db.OpenRecordset("select max(SlNo) as max_slno from JournalHead where JDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
        If Not IsNull(rec1!max_slno) Then
            Me.txtSlNo.Text = rec1!max_slno + 1
        Else
            Me.txtSlNo.Text = 1
        End If
        If Me.txtSlNo.Locked = True Then
            Me.CboMainGroup.SetFocus
        Else
            Me.txtSlNo.SetFocus
        End If
    End If
End Sub
Private Sub txtNarration_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub
Private Sub txtNarration_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtslno_GotFocus()
    Me.txtSlNo.SelStart = 0
    Me.txtSlNo.SelLength = Len(Me.txtSlNo.Text)
End Sub
Private Sub txtslno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.txtJDate.Text), 2)
        temp_month = Mid((Me.txtJDate.Text), 4, 2)
        temp_year = Right((Me.txtJDate.Text), 4)
        db.Execute ("delete * from temp_journal")
        Set rec = db.OpenRecordset("select * from JournalHead where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
        If Not rec.EOF Then
            Set rec2 = db.OpenRecordset("select * from LedgerMAster where AccId=" & rec("AccId"))
            If Not rec2.EOF Then
                Me.CboMainGroup.Text = rec2("Groupname")
                Me.cboMainAccount.Text = Trim(rec2("AccName"))
                Me.txtMainAmount.Text = Format(rec("Amount"), "#######0.00")
                Me.cboMainTran.Text = rec("ParentTrn")
                Me.txtChildTran.Text = rec("ChildTran")
                Me.txtNarration.Text = rec("NARRATION")
            End If
            Set rec3 = db.OpenRecordset("select * from Journaldetails where SlNo=" & Me.txtSlNo.Text & " and JDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
            If Not rec3.EOF Then
                While Not rec3.EOF
                    db.Execute ("insert into Temp_Journal (AccId,AccName,Amount,TranType) values(" & rec3("AccId") & ",'" & rec3("AccName") & "'," & rec3("Amount") & ",'" & rec3("TranType") & "')")
                    rec3.MoveNext
                Wend
            End If
            Data1.Refresh
            Me.cboMainAccount.SetFocus

            '-----------Journal Transaction Delete------------
            If JDELETE = "y" Then
                ans = MsgBox("Confirm Delete?", vbYesNo)
                If ans = 6 Then
                    Set rec1 = db.OpenRecordset("select * from JournalHead where JDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
                    If Not rec1.EOF Then
                        Set rec2 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Journal' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                        If Not rec2.EOF Then
                            db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec1("AccId") & " and SlNo>=" & rec2("SlNo"))
                            db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Journal' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                        End If
                    End If
                    '--------------JournalDetails delete------------
                    Set rec1 = db.OpenRecordset("select * from JournalDetails where JDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
                    If Not rec1.EOF Then
                        While Not rec1.EOF
                            Set rec2 = db.OpenRecordset("select * from LedgerMaster Where AccId=" & rec1("AccId"))
                            If Not rec2.EOF Then
                                Set rec3 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Journal' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                                If Not rec3.EOF Then
                                    db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec1("AccId") & " and SlNo>=" & rec3("SlNo"))
                                    db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Journal' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                                End If
                            End If
                            rec1.MoveNext
                        Wend
                    End If
                    db.Execute ("delete * from JournalHead where JDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
                    db.Execute ("delete * from JournalDetails where JDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
                End If
                db.Execute ("delete * from Temp_Journal")
                Data1.Refresh
                Me.txtMainAmount.Text = "0.00"
                Me.txtJDate.SetFocus
                JDELETE = "n"
                Me.txtSlNo.Locked = True
            End If
        End If
    End If
End Sub
