VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_receive 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Collection Voucher"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2400
      TabIndex        =   20
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   5415
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
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
         Left            =   3840
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNarration 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   2040
         Width           =   3975
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtamount 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txt_ch_dd_date 
         Height          =   315
         Left            =   3840
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.TextBox txtch_dd_no 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cbopayment 
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
         ItemData        =   "frm_receive.frx":0000
         Left            =   1320
         List            =   "frm_receive.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Discount"
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
         Left            =   2760
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Narration"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Banker"
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
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ch/DD Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ch./DD No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment by"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cboparty 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin MSMask.MaskEdBox txt_receipt_date 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.TextBox txtreceipt_no 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblGroup 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1320
         TabIndex        =   24
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1320
         TabIndex        =   23
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No."
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
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_receive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As Recordset, temp_pcode, temp_ledger_balance, temp_ledger_slno, bank_slno, bank_balance, Party_GroupId, tempsign
Private Sub cboBank_Change()
temp_day = Left((Me.txt_receipt_date.Text), 2)
temp_month = Mid((Me.txt_receipt_date.Text), 4, 2)
temp_year = Right((Me.txt_receipt_date.Text), 4)

Accperiod_day = Left(AccountingPeriod, 2)
Accperiod_month = Mid(AccountingPeriod, 4, 2)
Accperiod_year = Right(AccountingPeriod, 4)
Set rec1 = Db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboBank.ItemData(Me.cboBank.ListIndex) & " and Slno=(select Max(SlNo) from LedgerTran where AccId=" & Me.cboBank.ItemData(Me.cboBank.ListIndex) & " and TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "#)")
If Not rec1.EOF Then
bank_slno = rec1("Slno") + 1
bank_balance = rec1("Balance")
Else
bank_slno = 1
bank_balance = 0
End If
End Sub

Private Sub cboBank_Click()
cboBank_Change
End Sub

Private Sub cboBank_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtNarration.SetFocus
End If
End Sub
Private Sub cboParty_Change()
temp_day = Left((Me.txt_receipt_date.Text), 2)
temp_month = Mid((Me.txt_receipt_date.Text), 4, 2)
temp_year = Right((Me.txt_receipt_date.Text), 4)

Accperiod_day = Left(AccountingPeriod, 2)
Accperiod_month = Mid(AccountingPeriod, 4, 2)
Accperiod_year = Right(AccountingPeriod, 4)
Set rec1 = Db.OpenRecordset("select * from PartyDr where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
If Not rec1.EOF Then
Me.lblAddress.Caption = rec1("Address")
End If
Set rec2 = Db.OpenRecordset("select * from LedgerMaster Where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
If Not rec2.EOF Then
tempsign = rec2("Cr")
Party_GroupId = rec2("GroupId")
Me.lblGroup.Caption = rec2("GroupName")
End If
Set rec1 = Db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & " and SlNo=(select max(slno) from LedgerTran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
If Not rec1.EOF Then
Me.txtAmount.Text = rec1("Balance")
End If
cboParty.ToolTipText = cboParty.ItemData(cboParty.ListIndex)
End Sub

Private Sub cboParty_Click()
cboParty_Change
End Sub

Private Sub cboparty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cbopayment.SetFocus
End If
End Sub

Private Sub cbopayment_Change()
If cbopayment.Text = "Cash" Then
    Me.txt_ch_dd_date.Enabled = False
    Me.txtch_dd_no.Enabled = False
    Me.cboBank.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    Label8.Enabled = False
   
End If
If cbopayment.Text = "DD" Then
    Me.txt_ch_dd_date.Enabled = True
    Me.txtch_dd_no.Enabled = True
    Me.cboBank.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = True
    Label8.Enabled = True
    Me.cboBank.Clear
    Set rec1 = Db.OpenRecordset("select * from LedgerMaster where Groupname like 'Bank Accounts'")
    While Not rec1.EOF
    Me.cboBank.AddItem (rec1("AccName"))
    Me.cboBank.ItemData(Me.cboBank.NewIndex) = rec1("AccID")
    rec1.MoveNext
    Wend
    If Me.cboBank.ListCount > 0 Then
    Me.cboBank.ListIndex = 0
    End If
End If
If cbopayment.Text = "Cheque" Then
    Me.txt_ch_dd_date.Enabled = True
    Me.txtch_dd_no.Enabled = True
    Me.cboBank.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = True
    Label8.Enabled = True
    Me.cboBank.Clear
    Set rec1 = Db.OpenRecordset("select * from LedgerMaster where Groupname like 'Bank Accounts'")
    While Not rec1.EOF
    Me.cboBank.AddItem (rec1("AccName"))
    Me.cboBank.ItemData(Me.cboBank.NewIndex) = rec1("AccID")
    rec1.MoveNext
    Wend
    If Me.cboBank.ListCount > 0 Then
    Me.cboBank.ListIndex = 0
    End If
End If

End Sub

Private Sub cbopayment_Click()
cbopayment_Change
End Sub

Private Sub cbopayment_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtAmount.SetFocus
End If
End Sub
Private Sub cmdSave_Click()
ans = MsgBox("Save the Receipt?", vbYesNo)
If ans = 6 Then
temp_day = Left((Me.txt_receipt_date.Text), 2)
temp_month = Mid((Me.txt_receipt_date.Text), 4, 2)
temp_year = Right((Me.txt_receipt_date.Text), 4)

Accperiod_day = Left(AccountingPeriod, 2)
Accperiod_month = Mid(AccountingPeriod, 4, 2)
Accperiod_year = Right(AccountingPeriod, 4)
    
    If cbopayment.Text = "Cash" Then
        Set rec1 = Db.OpenRecordset("select * from LedgerMaster where AccName like 'Cash*'")
        If Not rec1.EOF Then
        temp_cashId = rec1("AccId")
            Set rec2 = Db.OpenRecordset("select * from LedgerTran where AccId=" & rec1("AccId") & " and Slno=(select Max(SlNo) from LedgerTran where AccId=" & rec1("AccId") & " and TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "#)")
            If Not rec2.EOF Then
            CashBookNo = rec2("slno") + 1
            cashbook_balance = rec2("Balance")
            End If
            Db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance+" & Val(Me.txtAmount.Text) & " where AccId=" & rec1("AccId") & " and SlNo > " & rec2("SlNo"))
            Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & CashBookNo & ",'" & Me.txt_receipt_date.Text & "','To " & Me.cboParty.Text & "'," & Me.txtAmount.Text & ",0," & cashbook_balance + Val(Me.txtAmount.Text) & "," & rec1("AccId") & ",'" & Me.txtNarration.Text & "','Receipt'," & Me.txtreceipt_no.Text & "," & rec1("GroupId") & ")")
        End If
    Else
        Db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance+" & Val(Me.txtAmount.Text) & " where AccID=" & Me.cboBank.ItemData(Me.cboBank.ListIndex) & " and SlNo>" & bank_slno - 1)
        Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & bank_slno & ",'" & Me.txt_receipt_date.Text & "','To " & Me.cboParty.Text & "'," & Me.txtAmount.Text & ",0," & bank_balance + Val(Me.txtAmount.Text) & "," & rec1("AccId") & ",'" & Me.txtNarration.Text & "','Receipt'," & Me.txtreceipt_no.Text & "," & rec1("GroupId") & ")")
        bank_balance = bank_balance + Val(Me.txtAmount.Text)
        bank_slno = bank_slno + 1
    End If
        If Val(Me.txtDiscount.Text) > 0 Then
        Set rec1 = Db.OpenRecordset("select * from LedgerMaster where AccName like 'Discount Allowed'")
        If Not rec1.EOF Then
            Set rec2 = Db.OpenRecordset("select * from LedgerTran where AccId=" & rec1("AccId") & " and Slno=(select Max(SlNo) from LedgerTran where AccId=" & rec1("AccId") & " and TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "#)")
            If Not rec2.EOF Then
            Discount_ledger_balance = rec2("Balance")
            Discount_ledger_SlNo = rec2("SlNo") + 1
            Db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance+" & Val(Me.txtDiscount.Text) & " where AccId=" & rec1("AccId") & " and SlNo > " & rec2("SlNo"))
            Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & Discount_ledger_SlNo & ",'" & Me.txt_receipt_date.Text & "','To " & Me.cboParty.Text & "'," & Me.txtDiscount.Text & ",0," & Discount_ledger_balance + Val(Me.txtDiscount.Text) & "," & rec1("AccId") & ",'" & Me.txtNarration.Text & "','Receipt'," & Me.txtreceipt_no.Text & "," & rec1("GroupId") & ")")
            End If
        End If
        End If
    temp_Amount = Val(Me.txtAmount.Text) + Val(Me.txtDiscount.Text)
    Set rec1 = Db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & " and SlNo=(select max(slno) from LedgerTran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & " and  TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "#)")
    If Not rec1.EOF Then
    temp_ledger_slno = rec1("Slno") + 1
    temp_ledger_balance = rec1("balance")
    Else
    temp_ledger_slno = 1
    temp_ledger_slno = 0
    End If

    If Me.cbopayment.Text = "Cash" Then
        Db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance " & tempsign & temp_Amount & "  where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & " and Slno > " & temp_ledger_slno - 1)
        Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & temp_ledger_slno & ",'" & Me.txt_receipt_date.Text & "','By Sundries',0," & temp_Amount & "," & temp_ledger_balance & tempsign & temp_Amount & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ",'" & Me.txtNarration.Text & "','Receipt'," & Me.txtreceipt_no.Text & "," & Party_GroupId & ")")
        Db.Execute ("insert into partycollection (Receiptno,ReceiptDate,Party,Paidby,Amount,AccID,Discount,DrAccId) values(" & Me.txtreceipt_no.Text & ",'" & Me.txt_receipt_date.Text & "','" & Me.cboParty.Text & "','" & Me.cbopayment.Text & "'," & Me.txtAmount.Text & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & "," & Me.txtDiscount.Text & "," & temp_cashId & ")")
    Else
        Db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance " & tempsign & temp_Amount & "  where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & " and Slno > " & temp_ledger_slno - 1)
        Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & temp_ledger_slno & ",'" & Me.txt_receipt_date.Text & "','By Bank',0," & temp_Amount & "," & temp_ledger_balance & tempsign & temp_Amount & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ",'" & Me.txtNarration.Text & "','Receipt'," & Me.txtreceipt_no.Text & "," & Party_GroupId & ")")
        Db.Execute ("insert into partycollection (Receiptno,ReceiptDate,Party,Paidby,Ch_DD_No,Ch_DD_Date,Banker,Amount,AccId,Discount,DrAccid) values(" & Me.txtreceipt_no.Text & ",'" & Me.txt_receipt_date.Text & "','" & Me.cboParty.Text & "','" & Me.cbopayment.Text & "','" & Me.txtch_dd_no.Text & "','" & Me.txt_ch_dd_date.Text & "','" & Me.cboBank.Text & "'," & Me.txtAmount.Text & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & "," & Me.txtDiscount.Text & "," & Me.cboBank.ItemData(Me.cboBank.ListIndex) & ")")
    End If
'    temp_ledger_balance = temp_ledger_balance - temp_amount
'    temp_ledger_slno = temp_ledger_slno + 1
    
'    Set rec1 = Db.OpenRecordset("select Tin from PartyDr where AccId=" & Me.cboparty.ItemData(Me.cboparty.ListIndex))
'        If Not rec1.EOF Then
'        TEMP_SUBFIX = Left(rec1("TIN"), 3)
'        If TEMP_SUBFIX = "SRI" Or TEMP_SUBFIX = "" Then
'        invtype = "RETAIL"
'        Else
'        invtype = "TAX"
'        End If
'        End If
'    diff_amount = Val(txtamount.Text)
'    Set rec1 = Db.OpenRecordset("select * from Invoicehead where invtype='" & invtype & "' and PCode='" & Me.cboparty.ItemData(Me.cboparty.ListIndex) & "' and Bill_Type='Credit' and Balance>0")
'    While Not rec1.EOF
'    temp_balance = rec1("Balance")
'        If diff_amount > 0 Then
'            If rec1("Balance") < diff_amount Then
'            Db.Execute ("update InvoiceHead set paid=paid+" & rec1("Balance") & ",Balance=Balance-" & rec1("Balance") & " where invtype='" & invtype & "' and pcode='" & Me.cboparty.ItemData(Me.cboparty.ListIndex) & "' and Bill_Type='Credit' and InvNo=" & rec1("InvNo"))
'            End If
'            If rec1("Balance") > diff_amount Then
'            Db.Execute ("update invoicehead set Paid=Paid+" & diff_amount & ",Balance=Balance-" & diff_amount & " where invtype='" & invtype & "' and pcode='" & Me.cboparty.ItemData(Me.cboparty.ListIndex) & "' and Bill_Type='Credit' and invno=" & rec1("InvNo"))
'            End If
'        End If
'        Db.Execute ("update invoicehead set Recipt_No='" & Me.txtreceipt_no.Text & "' where invno=" & rec1("InvNo") & " and invtype='" & invtype & "'")
'    rec1.MoveNext
'    diff_amount = Val(diff_amount) - Val(temp_balance)
'    Wend
'    diff_amount = 0
    Me.txtreceipt_no.Text = Val(txtreceipt_no.Text) + 1
    Me.txt_ch_dd_date.Text = "__/__/____"
    Me.txtAmount.Text = "0.00"
    Me.txtch_dd_no.Text = ""
    Me.cbopayment.ListIndex = 0
    Me.cboParty.ListIndex = 0
    temp_balance = 0
    Me.txtDiscount.Text = "0.00"
    Me.txtNarration.Text = ""
    Me.txtreceipt_no.SetFocus
    temp_Amount = 0
    
End If


End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
formname = "collection"
temp_day = Format(Date, "dd")
temp_month = Format(Date, "mm")
temp_year = Format(Date, "yyyy")
Me.txt_receipt_date.Text = temp_day & "/" & temp_month & "/" & temp_year

Set rec1 = Db.OpenRecordset("select * from LedgerMaster ")
While Not rec1.EOF
Me.cboParty.AddItem (rec1("AccName"))
Me.cboParty.ItemData(cboParty.NewIndex) = rec1("AccId")
rec1.MoveNext
Wend
Me.cboParty.ListIndex = 0
Me.cbopayment.ListIndex = 0
Set rec1 = Db.OpenRecordset("select max(receiptno) as max_slno from partycollection")
If Not IsNull(rec1!max_slno) Then
    temp_slno = rec1!max_slno + 1
Else
    temp_slno = 1
End If
Me.txtreceipt_no.Text = temp_slno

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
'db.Execute ("delete * from selectedinvoice")
End Sub

Private Sub txt_ch_dd_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cboBank.SetFocus
End If
End Sub

Private Sub txt_receipt_date_GotFocus()
txt_receipt_date.SelStart = 0
txt_receipt_date.SelLength = Len(txt_receipt_date.Text)
End Sub

Private Sub txt_receipt_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
temp_day = Left((Me.txt_receipt_date.Text), 2)
temp_month = Mid((Me.txt_receipt_date.Text), 4, 2)
temp_year = Right((Me.txt_receipt_date.Text), 4)

Accperiod_day = Left(AccountingPeriod, 2)
Accperiod_month = Mid(AccountingPeriod, 4, 2)
Accperiod_year = Right(AccountingPeriod, 4)
Set rec1 = Db.OpenRecordset("select max(Receiptno) as max_slno from partycollection where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
If Not IsNull(rec1!max_slno) Then
    Me.txtreceipt_no.Text = rec1!max_slno + 1
Else
    Me.txtreceipt_no.Text = 1
End If
    Me.cboParty.SetFocus
End If
End Sub

Private Sub txtAmount_GotFocus()
txtAmount.SelStart = 0
txtAmount.SelLength = Len(txtAmount.Text)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtDiscount.SetFocus
End If
'If KeyCode = 113 Then
'    frminvoice_list.Show vbModal
'End If
'If KeyCode = 13 Then
'    ans = MsgBox("Save the Receipt?", vbYesNo)
'    If ans = 6 Then
'
'    If cbopayment.Text = "Cash" Then
'        Db.Execute ("insert into partycollection (Receiptno,ReceiptDate,Party,Paidby,Amount) values(" & Me.txtreceipt_no.Text & ",'" & Me.txt_receipt_date.Text & "','" & temp_pcode & "','" & Me.cbopayment.Text & "'," & Me.txtAmount.Text & ")")
'        Db.Execute ("insert into partydrledger (TDate,Partyname,Particulars,Dr,Cr,Balance,Slno) values('" & Format(Me.txt_receipt_date.Text, "dd/mm/yyyy") & "','" & temp_pcode & "','Payment by " & cbopayment.Text & "[" & Trim(Me.txtreceipt_no.Text) & "]',0," & Val(Me.txtAmount.Text) & "," & temp_ledger_balance - Val(Me.txtAmount.Text) & "," & temp_ledger_slno & ")")
'    Else
'        Db.Execute ("insert into partycollection (Receiptno,ReceiptDate,Party,Paidby,Ch_DD_No,Ch_DD_Date,Banker,Amount) values(" & Me.txtreceipt_no.Text & ",'" & Me.txt_receipt_date.Text & "','" & temp_code & "','" & Me.cbopayment.Text & "','" & Me.txtch_dd_no.Text & "','" & Me.txt_ch_dd_date.Text & "','" & Me.txtbanker.Text & "'," & Me.txtAmount.Text & ")")
'        Db.Execute ("insert into partydrledger (TDate,Partyname,Particulars,Dr,Cr,Balance,Slno) values('" & Format(Me.txt_receipt_date.Text, "dd/mm/yyyy") & "','" & temp_pcode & "','Payment by " & cbopayment.Text & "[" & Me.txtch_dd_no.Text & "]',0," & Val(Me.txtAmount.Text) & "," & temp_ledger_balance - Val(Me.txtAmount.Text) & "," & temp_ledger_slno & ")")
'    End If
'
'    temp_amount = Val(txtAmount.Text)
'
'    Set rec1 = Db.OpenRecordset("select * from selectedinvoice")
'    While Not rec1.EOF
'        If rec1("amount") < temp_amount Then
'            Db.Execute ("update invoicehead set paid=paid+" & rec1("amount") & ",balance=balance-" & rec1("amount") & " where invoiceno=" & rec1("invoiceno") & " and pcode='" & temp_pcode & "'")
'            Set rec2 = Db.OpenRecordset("select inv_type from invoicehead where invoiceno=" & rec1("invoiceno") & " and pcode='" & temp_pcode & "'")
'            If Not rec2.EOF Then
'                Db.Execute ("insert into collectiondetails (receipt_no,invoiceno,amount,inv_type) values(" & Me.txtreceipt_no.Text & "," & rec1("invoiceno") & "," & rec1("amount") & ",'" & rec2("inv_type") & "')")
'            End If
'        Else
'            Db.Execute ("update invoicehead set paid=paid+" & temp_amount & ",balance=balance-" & temp_amount & " where invoiceno=" & rec1("invoiceno") & " and pcode='" & temp_pcode & "'")
'            Set rec2 = Db.OpenRecordset("select inv_type from invoicehead where invoiceno=" & rec1("invoiceno") & " and pcode='" & temp_pcode & "'")
'            If Not rec2.EOF Then
'                Db.Execute ("insert into collectiondetails (receipt_no,invoiceno,amount,inv_type) values(" & Me.txtreceipt_no.Text & "," & rec1("invoiceno") & "," & temp_amount & ",'" & rec2("inv_type") & "')")
'            End If
'        End If
'        temp_amount = temp_amount - rec1("amount")
'        rec1.MoveNext
'    Wend
'    Me.txtreceipt_no.Text = Val(txtreceipt_no.Text) + 1
'    Me.txt_ch_dd_date.Text = "__/__/____"
'    Me.txtAmount.Text = "0.00"
'    Me.txtbanker.Text = ""
'    Me.txtch_dd_no.Text = ""
'    Me.cbopayment.ListIndex = 0
'    Db.Execute ("delete * from selectedinvoice")
'    Me.txtreceipt_no.SetFocus
'    End If
'
'End If
End Sub

Private Sub txtbanker_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtAmount.SetFocus
End If
End Sub

Private Sub txtbanker_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtch_dd_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txt_ch_dd_date.SetFocus
End If
End Sub
Private Sub txtDiscount_GotFocus()
Me.txtDiscount.SelStart = 0
Me.txtDiscount.SelLength = Len(Me.txtDiscount.Text)
End Sub
Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Me.cbopayment.Text = "Cash" Then
    Me.txtNarration.SetFocus
    Else
    Me.txtch_dd_no.SetFocus
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

Private Sub txtreceipt_no_GotFocus()
Me.txtreceipt_no.SelStart = 0
Me.txtreceipt_no.SelLength = Len(Me.txtreceipt_no.Text)
End Sub

Private Sub txtreceipt_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txt_receipt_date.SetFocus
End If
End Sub
