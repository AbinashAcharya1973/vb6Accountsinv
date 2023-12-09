VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmReceiptVoucher 
   BackColor       =   &H00E3E3E3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Voucher"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8640
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
      Left            =   5768
      TabIndex        =   28
      Top             =   6930
      Width           =   975
   End
   Begin VB.TextBox txtbalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   6360
      Width           =   1095
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
      Left            =   4478
      TabIndex        =   21
      Top             =   6930
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   8415
      Begin VB.TextBox TxtCrAmount 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   6720
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox CboCrAccount 
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
         TabIndex        =   4
         Top             =   600
         Width           =   3855
      End
      Begin VB.ComboBox CboGroup 
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
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label10 
         Caption         =   "Receive From"
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
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LBLADR2 
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   1320
         Width           =   6615
      End
      Begin VB.Label LblAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Label6 
         Caption         =   "Cr Amount"
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
         Left            =   5640
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Credit Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Acc Group"
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
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8415
      Begin VB.TextBox TxtDrAmount 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox CboDrAccount 
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
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox TxtReceiptNo 
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
         TabIndex        =   1
         Text            =   "0"
         Top             =   120
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TxtReceiptDate 
         Height          =   315
         Left            =   1560
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
      Begin VB.Label Label3 
         Caption         =   "Debit Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Dr Amount"
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
         Left            =   5640
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Receive By"
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
         Width           =   1335
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
         TabIndex        =   12
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
         Left            =   5640
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   8415
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmReceiptVoucher.frx":0000
         Height          =   2055
         Left            =   120
         OleObjectBlob   =   "FrmReceiptVoucher.frx":0014
         TabIndex        =   20
         Top             =   120
         Width           =   8175
      End
      Begin VB.TextBox TxtNarration 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   8175
      End
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "SAVE"
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
      Left            =   1898
      TabIndex        =   7
      Top             =   6930
      Width           =   975
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "EDIT"
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
      Left            =   3188
      TabIndex        =   6
      Top             =   6930
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TempReceipt"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   975
   End
End
Attribute VB_Name = "FrmReceiptVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As DAO.Recordset, rec As DAO.Recordset, rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset, rec5 As Recordset, TEMPDELETE
Attribute rec.VB_VarUserMemId = 1073938432
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432
Attribute rec5.VB_VarUserMemId = 1073938432
Attribute TEMPDELETE.VB_VarUserMemId = 1073938432

Private Sub CboCrAccount_Change()
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex))
    If Not IsNull(rec1!Address1) Then
        Me.LblAddress.Caption = rec1("Address1")
    Else
        Me.LblAddress.Caption = ""
    End If
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex))
    If Not IsNull(rec1!Address2) Then
        Me.LBLADR2.Caption = rec1("aDDRESS2")
    Else
        Me.LBLADR2.Caption = ""
    End If

    Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex))
    If Not IsNull(rs!max_dr) Then
        temp_dr = rs!max_dr
    End If
    Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex))
    If Not IsNull(rs!max_cr) Then
        temp_cr = rs!max_cr
    End If
    Set rs = db.OpenRecordset("select * from ledgermaster where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex))
    If Not rs.EOF Then
        If rs("BalanceType") = "Dr" Then
            temp_dr = temp_dr + rs("OBalance")
        End If
        If rs("BalanceType") = "Cr" Then
            temp_cr = temp_cr + rs("OBalance")
        End If
        Me.txtbalance.Text = Format(temp_dr - temp_cr, "#######0.00")
    Else
        Me.txtbalance.Text = 0
    End If

End Sub
Private Sub CboCrAccount_Click()
    CboCrAccount_Change
End Sub

Private Sub cboCrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtCrAmount.SetFocus
    End If
End Sub

Private Sub CboDrAccount_Change()
'If Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) = 7 Then
'Me.TxtNarration.Text = "M.R NO"
'Else
'Me.TxtNarration.Text = "CHQ/DD"
'End If
End Sub

Private Sub CboDrAccount_Click()
'CboDrAccount_Change
End Sub

Private Sub cboDrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.CboGroup.SetFocus
    End If
End Sub

Private Sub cboGroup_Change()
    Set rec1 = db.OpenRecordset("select * from LedgerMAster where GroupID=" & Me.CboGroup.ItemData(Me.CboGroup.ListIndex))
    Me.CboCrAccount.Clear
    While Not rec1.EOF
        Me.CboCrAccount.AddItem (rec1("AccName"))
        Me.CboCrAccount.ItemData(Me.CboCrAccount.NewIndex) = rec1("AccID")
        rec1.MoveNext
    Wend
    If Me.CboCrAccount.ListCount > 0 Then
        Me.CboCrAccount.ListIndex = 0
    End If
End Sub
Private Sub cboGroup_Click()
    cboGroup_Change
End Sub

Private Sub cboGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.CboCrAccount.SetFocus
    End If
    If KeyCode = 27 Then
        If Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) = 7 Then
            Me.TxtNarration.Text = "M.R NO-"
        Else
            Me.TxtNarration.Text = "CHQ NO-"
        End If
        Me.TxtNarration.SetFocus
    End If
End Sub
Private Sub cmddelete_Click()
    db.Execute ("delete * from tempreceipt")
    Data1.Refresh
    Me.TxtDrAmount.Text = "0.00"
    Me.TxtReceiptDate.SetFocus
    Me.TxtReceiptNo.Locked = False
    TEMPDELETE = "y"
End Sub
Private Sub cmdedit_Click()
    db.Execute ("delete * from TempReceipt")
    Data1.Refresh
    Me.TxtDrAmount.Text = "0.00"
    Me.TxtReceiptNo.Locked = False
    Me.TxtReceiptDate.SetFocus
End Sub

Private Sub cmdprint_Click()
frmreceiptprint.Show 0
End Sub

Private Sub cmdsave_Click()
On Error GoTo errtrap
'-----Change----------
    Dim rs As DAO.Recordset
    temp_day = Left((Me.TxtReceiptDate.Text), 2)
    temp_month = Mid((Me.TxtReceiptDate.Text), 4, 2)
    temp_year = Right((Me.TxtReceiptDate.Text), 4)
    temp_from = temp_month & "/" & temp_day & "/" & temp_year
    Accperiod_day = Left(AccountingPeriod, 2)
    Accperiod_month = Mid(AccountingPeriod, 4, 2)
    Accperiod_year = Right(AccountingPeriod, 4)
    ans = MsgBox("Confirm This Transaction?", vbYesNo)
    If ans = 6 Then
        Set rs = db.OpenRecordset("select * from TempReceipt")
        If Not rs.EOF Then


            Set rec = db.OpenRecordset("select * from ReceiptDetails where ReceiptDate = #" & temp_from & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
            If Not rec.EOF Then
                Set rec5 = db.OpenRecordset("select * from receipthead where ReceiptDate = #" & temp_from & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
                If Not rec5.EOF Then
                    DrAccId = rec5("AccId")
                End If
                While Not rec.EOF
                    Set rec1 = db.OpenRecordset("select * from LedgertRan where AccId=" & DrAccId & " and VoucherType='Receipt' and Tdate=#" & temp_from & "# and TranAccId=" & rec("AccId") & " and VoucherSlno=" & Me.TxtReceiptNo.Text)
                    If Not rec1.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec1("AccId") & " and SlNo>=" & rec1("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Receipt' and VoucherSlno=" & Me.TxtReceiptNo.Text & " and TDate =#" & temp_from & "# and TranAccId=" & rec("AccId"))
                    End If
                    Set rec1 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec("AccId") & " and Tdate=#" & temp_from & "# and VoucherType='Receipt' and VoucherslNo=" & Me.TxtReceiptNo.Text & " and TranAccId=" & DrAccId)
                    If Not rec1.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec("AccId") & " and SlNo>=" & rec1("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec("AccId") & " and VoucherType='Receipt' and VoucherSlno=" & Me.TxtReceiptNo & " and TDate =#" & temp_from & "# and TranAccId=" & DrAccId)
                    End If
                    rec.MoveNext
                Wend
                db.Execute ("delete * from ReceiptHead where ReceiptNo=" & Me.TxtReceiptNo.Text & " and ReceiptDate=#" & temp_from & "#")
                db.Execute ("Delete * from ReceiptDetails where ReceiptDate = #" & temp_from & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
            End If
            '-----------------Bill Wiswe Adjustment------------------------
            '-----------------Delete Previous Entry-------------
            Set rec = db.OpenRecordset("select * from TempReceipt")
            If Not rec.EOF Then
                Set rec1 = db.OpenRecordset("select * from CollectionDetails where ReciptNo=" & Me.TxtReceiptNo.Text & " and ReciptDate=#" & temp_from & "#")
                If Not rec1.EOF Then
                    While Not rec1.EOF
                        db.Execute ("update Invoicehead set Paid=Paid-" & rec1("Amount") & ",Balance=Balance+" & rec1("Amount") & " where InvNo=" & rec1("InvNo") & " and AccId=" & rec1("AccId"))
                        rec1.MoveNext
                    Wend
                End If
                db.Execute ("delete * from CollectionDetails where ReciptNo=" & Me.TxtReceiptNo.Text & " and ReciptDate=#" & temp_from & "#")
            End If
            '-------------------------------------------------------
            '=============New Entry===========
            Set rec = db.OpenRecordset("select * from TempReceipt")
            If Not rec.EOF Then
                db.Execute ("insert into ReceiptHead (ReceiptNo,ReceiptDate,AccId,AccName,Amount,ParentTrn,ChildTran,Narration) values(" & Me.TxtReceiptNo.Text & ",'" & Me.TxtReceiptDate.Text & "'," & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ",'" & Me.CboDrAccount.Text & "'," & Me.TxtDrAmount.Text & ",'DEBIT','CREDIT','" & Trim(Me.TxtNarration.Text) & "')")
                While Not rec.EOF
                    db.Execute ("insert into ReceiptDetails (ReceiptNo,ReceiptDate,AccId,AccName,Amount,TranType) values(" & Me.TxtReceiptNo.Text & ",'" & Me.TxtReceiptDate.Text & "'," & rec("AccId") & ",'" & rec("AccName") & "'," & rec("Amount") & ",'CREDIT')")
                    Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex))
                    If Not rec3.EOF Then
                        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & " and SlNo=(select max(slno) from LedgerTran where AccID=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ")")
                        If Not rec4.EOF Then
                            temp_LegerSlno = rec4("SlNo") + 1
                        Else
                            temp_LegerSlno = 1
                        End If
                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & temp_LegerSlno & ",'" & Me.TxtReceiptDate.Text & "','To " & rec("AccName") & "'," & rec("Amount") & ",0,0," & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ",'" & Trim(Me.TxtNarration.Text) & "','Receipt'," & Me.TxtReceiptNo.Text & "," & rec3("GroupId") & "," & rec("AccId") & ")")
                    End If
                    Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccId=" & rec("AccId"))
                    If Not rec3.EOF Then
                        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec("AccId") & " and SlNo=(select max(slno) from LedgerTran where AccID=" & rec("AccId") & ")")
                        If Not rec4.EOF Then
                            Child_LedgerSlno = rec4("SlNo") + 1
                        Else
                            Child_LedgerSlno = 1
                        End If
                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Child_LedgerSlno & ",'" & Me.TxtReceiptDate.Text & "','By " & Me.CboDrAccount.Text & "',0," & rec("Amount") & ",0," & rec("AccId") & ",'" & Trim(Me.TxtNarration.Text) & "','Receipt'," & Me.TxtReceiptNo.Text & "," & rec3("GroupId") & "," & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ")")
                    End If

                    '------------------Bill Wise New Entry----------------
                    '    Diff_Amount = Val(rec("Amount"))
                    '        Set rec1 = db.OpenRecordset("select * from Invoicehead where AccId=" & rec("AccId") & " and Balance>0")
                    '        While Not rec1.EOF
                    '        temp_balance = rec1("Balance")
                    '        If Diff_Amount > 0 Then
                    '            If rec1("Balance") < Diff_Amount Then
                    '            db.Execute ("insert into CollectionDetails (ReciptNo,ReciptDate,InvNo,Amount,AccId) values(" & Me.TxtReceiptNo.Text & ",'" & Me.TxtReceiptDate.Text & "'," & rec1("InvNo") & "," & rec1("Balance") & "," & rec("AccId") & ")")
                    '            db.Execute ("update InvoiceHead set paid=paid+" & rec1("Balance") & ",Balance=Balance-" & rec1("Balance") & " where AccId=" & rec("AccId") & " and InvNo=" & rec1("InvNo"))
                    '            End If
                    '            If rec1("Balance") > Diff_Amount Then
                    '            db.Execute ("insert into CollectionDetails (ReciptNo,ReciptDate,InvNo,Amount,AccId) values(" & Me.TxtReceiptNo.Text & ",'" & Me.TxtReceiptDate.Text & "'," & rec1("InvNo") & "," & Diff_Amount & "," & rec("AccId") & ")")
                    '            db.Execute ("update invoicehead set Paid=Paid+" & Diff_Amount & ",Balance=Balance-" & Diff_Amount & " where AccId=" & rec("AccId") & " and invno=" & rec1("InvNo"))
                    '            End If
                    '        End If
                    '        rec1.MoveNext
                    '        Diff_Amount = Val(Diff_Amount) - Val(temp_balance)
                    '        Wend


                    rec.MoveNext
                Wend
            End If

            db.Execute ("delete * from TempReceipt")
            Data1.Refresh
            Me.TxtNarration.Text = ""
            Me.TxtDrAmount.Text = "0.00"
            Me.TxtReceiptDate.SetFocus
            Me.TxtReceiptNo.Locked = True
        End If
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub
Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Me.TxtDrAmount.Text = Val(Me.TxtDrAmount.Text) - Val(Me.DBGrid1.Columns(2))
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Me.TxtReceiptDate.Text = Format(Date, "dd/mm/yyyy")
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where Groupname like 'Cash-In-Hand' or groupname like 'Bank Accounts'")
    While Not rec1.EOF
        Me.CboDrAccount.AddItem (rec1("AccName"))
        Me.CboDrAccount.ItemData(Me.CboDrAccount.NewIndex) = rec1("AccId")
        rec1.MoveNext
    Wend
    If Me.CboDrAccount.ListCount > 0 Then
        Me.CboDrAccount.ListIndex = 0
    End If
    Set rec1 = db.OpenRecordset("select * from Groups")
    While Not rec1.EOF
        Me.CboGroup.AddItem (rec1("GroupName"))
        Me.CboGroup.ItemData(Me.CboGroup.NewIndex) = rec1("GroupID")
        rec1.MoveNext
    Wend
    If Me.CboGroup.ListCount > 0 Then
        Me.CboGroup.ListIndex = 0
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("delete * from tempReceipt")
    Data1.Refresh
End Sub

Private Sub TxtCrAmount_GotFocus()
    Me.TxtCrAmount.SelStart = 0
    Me.TxtCrAmount.SelLength = Len(Me.TxtCrAmount.Text)
End Sub
Private Sub TxtCrAmount_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errtrap
    If KeyCode = 13 Then
        If Val(Me.TxtCrAmount.Text) > 0 Then
            Set rec1 = db.OpenRecordset("select * from TempReceipt where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex))
            If rec1.EOF Then
                Me.TxtDrAmount.Text = Format(Val(Me.TxtDrAmount.Text) + Val(Me.TxtCrAmount.Text), "############0.00")
                db.Execute ("insert into tempReceipt (AccId,AccName,Amount,TranType) values(" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex) & ",'" & Me.CboCrAccount.Text & "'," & Me.TxtCrAmount.Text & ",'CREDIT')")
                Data1.Refresh
                Me.TxtCrAmount.Text = "0.00"
                Me.CboCrAccount.ListIndex = 0
                Me.CboGroup.SetFocus
            Else
                MsgBox "Allready Exists", vbCritical
            End If
        Else
            MsgBox "Zero Not Enter", vbCritical
        End If
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub TxtNarration_GotFocus()
    Me.TxtNarration.SelStart = 0
    Me.TxtNarration.SelLength = Len(Me.TxtNarration.Text)
End Sub

Private Sub txtNarration_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.CmdSave.SetFocus
    End If
End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TxtReceiptDate_GotFocus()
    Me.TxtReceiptDate.SelStart = 0
    Me.TxtReceiptDate.SelLength = Len(Me.TxtReceiptDate.Text)
End Sub
Private Sub TxtReceiptDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errtrap
    If KeyCode = 13 Then
        temp_day = Left((Me.TxtReceiptDate.Text), 2)
        temp_month = Mid((Me.TxtReceiptDate.Text), 4, 2)
        temp_year = Right((Me.TxtReceiptDate.Text), 4)

        Accperiod_day = Left(AccountingPeriod, 2)
        Accperiod_month = Mid(AccountingPeriod, 4, 2)
        Accperiod_year = Right(AccountingPeriod, 4)
        Set rec1 = db.OpenRecordset("select max(ReceiptNo) as max_slno from ReceiptHead where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
        If Not IsNull(rec1!max_slno) Then
            Me.TxtReceiptNo.Text = rec1!max_slno + 1
        Else
            Me.TxtReceiptNo.Text = 1
        End If
        If Me.TxtReceiptNo.Locked = False Then
            Me.TxtReceiptNo.SetFocus
        Else
            Me.CboDrAccount.SetFocus
        End If
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub
Private Sub TxtReceiptNo_GotFocus()
    Me.TxtReceiptNo.SelStart = 0
    Me.TxtReceiptNo.SelLength = Len(Me.TxtReceiptNo.Text)
End Sub
Private Sub TxtReceiptNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.TxtReceiptDate.Text), 2)
        temp_month = Mid((Me.TxtReceiptDate.Text), 4, 2)
        temp_year = Right((Me.TxtReceiptDate.Text), 4)
        db.Execute ("delete * from tempreceipt")
        Me.TxtDrAmount.Text = "0.00"
        Set rec1 = db.OpenRecordset("select * from ReceiptDetails where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
        If Not rec1.EOF Then
            While Not rec1.EOF
                db.Execute ("insert into TempReceipt (AccId,AccName,Amount,TranType) values(" & rec1("AccId") & ",'" & rec1("AccName") & "'," & rec1("Amount") & ",'" & rec1("TranType") & "')")
                Me.TxtDrAmount.Text = Format(Val(Me.TxtDrAmount.Text) + rec1("Amount"), "#############0.00")
                rec1.MoveNext
            Wend
            Data1.Refresh
        End If
        Set rec1 = db.OpenRecordset("select * from ReceiptHead where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
        If Not rec1.EOF Then
            Set rec2 = db.OpenRecordset("select * from ledgermaster where accid=" & rec1("AccId"))
            If Not rec2.EOF Then
                Me.TxtNarration.Text = Trim(rec1("Narration"))
                Me.CboDrAccount.Text = rec2("AccName")
            End If
        End If
        Me.CboGroup.SetFocus
        '---------Deleteing Voucher----------------
        If TEMPDELETE = "y" Then
            ans = MsgBox("Confirm Delete?", vbYesNo)
            If ans = 6 Then
                Set rec1 = db.OpenRecordset("select * from receipthead where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
                If Not rec1.EOF Then
                    Set rec2 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Receipt' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.TxtReceiptNo.Text)
                    If Not rec2.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec2("Dr") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec2("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Receipt' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.TxtReceiptNo.Text)
                    End If
                End If
                Set rec1 = db.OpenRecordset("select * from ReceiptDetails where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
                If Not rec1.EOF Then
                    While Not rec1.EOF
                        Set rec2 = db.OpenRecordset("select * from LedgerMAster Where AccId=" & rec1("AccId"))
                        If Not rec2.EOF Then
                            temp_sign = rec2("Dr")
                            Set rec3 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Receipt' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.TxtReceiptNo.Text)
                            If Not rec3.EOF Then
                                db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance " & temp_sign & rec3("Cr") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec3("SlNo"))
                                db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Receipt' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.TxtReceiptNo.Text)
                            End If
                        End If
                        rec1.MoveNext
                    Wend
                End If
                '--Bill Wise Adjust---------------
                Set rec1 = db.OpenRecordset("select * from CollectionDetails where ReciptNo=" & Me.TxtReceiptNo.Text & " and ReciptDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                If Not rec1.EOF Then
                    While Not rec1.EOF
                        db.Execute ("update Invoicehead set Paid=Paid-" & rec1("Amount") & ",Balance=Balance+" & rec1("Amount") & " where InvNo=" & rec1("InvNo") & " and AccId=" & rec1("AccId"))
                        rec1.MoveNext
                    Wend
                End If
                db.Execute ("delete * from CollectionDetails where ReciptNo=" & Me.TxtReceiptNo.Text & " and ReciptDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")

                db.Execute ("delete * from Receipthead where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
                db.Execute ("delete * from ReceiptDetails where ReceiptDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and ReceiptNo=" & Me.TxtReceiptNo.Text)
            End If
            db.Execute ("delete * from tempreceipt")
            Data1.Refresh
            Me.TxtDrAmount.Text = "0.00"
            Me.TxtReceiptDate.SetFocus
            Me.TxtReceiptNo.Locked = True
        End If
        TEMPDELETE = "n"
    End If
End Sub
