VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmContra 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contra Voucher"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   5415
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtNarration 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1200
         MaxLength       =   189
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtAmount 
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
         Left            =   1200
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cboCrAccount 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Cr Account"
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
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cboDrAccount 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtVoucherNo 
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
         Left            =   1200
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtContraDate 
         Height          =   315
         Left            =   3720
         TabIndex        =   0
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label7 
         Caption         =   "Balance"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBalance 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Dr Account"
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
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Voucher No"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset

Private Sub cboCrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtAmount.SetFocus
End If
End Sub

Private Sub cboDrAccount_Change()
Set rec1 = Db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex) & " and SlNo=(select max(SlNo) from LedgerTran where AccId=" & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex) & ")")
If Not rec1.EOF Then
Me.lblBalance.Caption = Format(rec1("Balance"), "######0.00")
End If
Me.cboDrAccount.ToolTipText = rec1("AccId")
End Sub
Private Sub cboDrAccount_Click()
cboDrAccount_Change
End Sub
Private Sub cboDrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cboCrAccount.SetFocus
End If
End Sub

Private Sub cmdSave_Click()
ans = MsgBox("Save This?", vbYesNo)
If ans = 6 Then
temp_day = Left((Me.txtContraDate.Text), 2)
temp_month = Mid((Me.txtContraDate.Text), 4, 2)
temp_year = Right((Me.txtContraDate.Text), 4)

Accperiod_day = Left(AccountingPeriod, 2)
Accperiod_month = Mid(AccountingPeriod, 4, 2)
Accperiod_year = Right(AccountingPeriod, 4)
    Set rec1 = Db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex) & " and Slno=(select Max(SlNo) from LedgerTran where AccId=" & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex) & " and TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "#)")
    If Not rec1.EOF Then
    Dr_Account_Slno = rec1("SlNo") + 1
    Dr_Account_Balance = rec1("Balance")
    Else
    Dr_Account_Slno = 1
    Dr_Account_Balance = 0
    End If
    Set rec1 = Db.OpenRecordset("select Groupid from LedgerMaster where AccID= " & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex))
    If Not rec1.EOF Then
    Db.Execute ("update ledgerTran set SlNo=SlNo+1,Balance=Balance+" & Val(Me.txtAmount.Text) & " where AccId=" & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex) & " and SlNo >= " & Dr_Account_Slno)
    Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & Dr_Account_Slno & ",'" & Me.txtContraDate.Text & "','" & Me.cboDrAccount.Text & "'," & Me.txtAmount.Text & ",0," & Dr_Account_Balance + Val(Me.txtAmount.Text) & "," & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex) & ",'" & Me.txtNarration.Text & "','Contra'," & Me.txtVoucherNo.Text & "," & rec1("Groupid") & ")")
    End If
    Set rec1 = Db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboCrAccount.ItemData(Me.cboCrAccount.ListIndex) & " and Slno=(select Max(SlNo) from LedgerTran where AccId=" & Me.cboCrAccount.ItemData(Me.cboCrAccount.ListIndex) & " and TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "#)")
    If Not rec1.EOF Then
    Cr_Account_Slno = rec1("SlNo") + 1
    Cr_Account_Balance = rec1("Balance")
    Else
    Dr_Account_Slno = 1
    Dr_Account_Balance = 0
    End If
    Set rec1 = Db.OpenRecordset("select Groupid from LedgerMaster where AccID= " & Me.cboCrAccount.ItemData(Me.cboCrAccount.ListIndex))
    Db.Execute ("update ledgerTran set SlNo=SlNo+1,Balance=Balance-" & Val(Me.txtAmount.Text) & " where AccId=" & Me.cboCrAccount.ItemData(Me.cboCrAccount.ListIndex) & " and SlNo >= " & Cr_Account_Slno)
    Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & Cr_Account_Slno & ",'" & Me.txtContraDate.Text & "','" & Me.cboCrAccount.Text & "',0," & Me.txtAmount.Text & "," & Cr_Account_Balance & " - " & Val(Me.txtAmount.Text) & "," & Me.cboCrAccount.ItemData(Me.cboCrAccount.ListIndex) & ",'" & Me.txtNarration.Text & "','Contra'," & Me.txtVoucherNo.Text & "," & rec1("GroupId") & ")")
    
    Db.Execute ("insert into ContraVoucher (SlNo,ContraDate,DrAccId,DrAccName,CrAccId,CrAccName,Amount,Remarks) values(" & Me.txtVoucherNo.Text & ",'" & Me.txtContraDate.Text & "'," & Me.cboDrAccount.ItemData(Me.cboDrAccount.ListIndex) & ",'" & Me.cboDrAccount.Text & "'," & Me.cboCrAccount.ItemData(Me.cboCrAccount.ListIndex) & ",'" & Me.cboCrAccount.Text & "'," & Me.txtAmount.Text & ",'" & Me.txtNarration.Text & "')")
    Me.txtVoucherNo.Text = Val(Me.txtVoucherNo.Text) + 1
    Me.txtAmount.Text = "0.00"
    Me.txtNarration.Text = ""
    Me.txtContraDate.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.txtContraDate.Text = Format(Date, "dd/mm/yyyy")
Set rec1 = Db.OpenRecordset("select max(slno) as max_no from Contravoucher where ContraDate=#" & Format(Date, "mm") & "/" & Format(Date, "dd") & "/" & Format(Date, "yyyy") & "#")
If Not IsNull(rec1!Max_no) Then
Me.txtVoucherNo.Text = rec1!Max_no + 1
Else
Me.txtVoucherNo.Text = 1
End If
Set rec1 = Db.OpenRecordset("select * from LedgerMaster where Groupname like 'Cash-In-Hand' or groupname like 'Bank Accounts'")
While Not rec1.EOF
Me.cboDrAccount.AddItem (rec1("AccName"))
Me.cboCrAccount.AddItem (rec1("AccName"))
Me.cboDrAccount.ItemData(Me.cboDrAccount.NewIndex) = rec1("AccId")
Me.cboCrAccount.ItemData(Me.cboCrAccount.NewIndex) = rec1("AccId")
rec1.MoveNext
Wend
If Me.cboDrAccount.ListCount > 0 Then
Me.cboDrAccount.ListIndex = 0
Me.cboCrAccount.ListIndex = 0
End If
End Sub

Private Sub txtAmount_GotFocus()
Me.txtAmount.SelStart = 0
Me.txtAmount.SelLength = Len(Me.txtAmount.Text)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtNarration.SetFocus
End If
End Sub

Private Sub txtContraDate_GotFocus()
Me.txtContraDate.SelStart = 0
Me.txtContraDate.SelLength = Len(Me.txtContraDate.Text)
End Sub
Private Sub txtContraDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
temp_day = Left((Me.txtContraDate.Text), 2)
temp_month = Mid((Me.txtContraDate.Text), 4, 2)
temp_year = Right((Me.txtContraDate.Text), 4)

Accperiod_day = Left(AccountingPeriod, 2)
Accperiod_month = Mid(AccountingPeriod, 4, 2)
Accperiod_year = Right(AccountingPeriod, 4)
Set rec1 = Db.OpenRecordset("select max(SlNo) as max_no from ContraVoucher where ContraDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
If Not IsNull(rec1!Max_no) Then
Me.txtVoucherNo.Text = rec1!Max_no + 1
Else
Me.txtVoucherNo.Text = 1
End If
Me.cboDrAccount.SetFocus
End If
End Sub

Private Sub txtNarration_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cmdSave.SetFocus
End If
End Sub
