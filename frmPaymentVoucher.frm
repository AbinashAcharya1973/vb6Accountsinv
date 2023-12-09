VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPaymentVoucher 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Voucher"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
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
      Left            =   5648
      TabIndex        =   25
      Top             =   6360
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
      TabIndex        =   21
      Top             =   6360
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
      Left            =   3228
      TabIndex        =   20
      Top             =   6360
      Width           =   975
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
      Left            =   2018
      TabIndex        =   16
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E3E3E3&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   8415
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
         TabIndex        =   18
         Top             =   2280
         Width           =   8175
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmPaymentVoucher.frx":0000
         Height          =   2055
         Left            =   120
         OleObjectBlob   =   "frmPaymentVoucher.frx":0014
         TabIndex        =   17
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8415
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
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TxtCrAmount 
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtSlNo 
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   120
         Width           =   1455
      End
      Begin MSMask.MaskEdBox TxtPDate 
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
         Caption         =   "Paid From"
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
         TabIndex        =   23
         Top             =   600
         Width           =   1395
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1395
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
         Left            =   6000
         TabIndex        =   13
         Top             =   600
         Width           =   1095
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
         Left            =   6000
         TabIndex        =   12
         Top             =   120
         Width           =   855
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
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   8415
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
         Width           =   4215
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
         TabIndex        =   4
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TxtDrAmount 
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
         Left            =   6840
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Paid To"
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
         TabIndex        =   24
         Top             =   600
         Width           =   1395
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
         Width           =   6855
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
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1335
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
         Left            =   6000
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
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
      RecordSource    =   "TempPayment"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmPaymentVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim rs As Recordset, rec As DAO.Recordset, rec1 As DAO.Recordset, rec2 As DAO.Recordset, rec3 As Recordset, rec4 As Recordset, rec5 As Recordset, VDELETE
Attribute rec.VB_VarUserMemId = 1073938432
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432
Attribute rec5.VB_VarUserMemId = 1073938432
Attribute VDELETE.VB_VarUserMemId = 1073938432
Private Sub cboCrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.CboGroup.SetFocus
    End If
End Sub

Private Sub CboDrAccount_Change()
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex))
    If Not IsNull(rec1!Address1) Then
        Me.LblAddress.Caption = rec1("Address1")
    Else
        Me.LblAddress.Caption = ""
    End If
End Sub
Private Sub CboDrAccount_Click()
    CboDrAccount_Change
End Sub
Private Sub cboDrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtDrAmount.SetFocus
    End If
End Sub
Private Sub cboGroup_Change()
    Set rec1 = db.OpenRecordset("select * from LedgerMAster where GroupID=" & Me.CboGroup.ItemData(Me.CboGroup.ListIndex))
    Me.CboDrAccount.Clear
    While Not rec1.EOF
        Me.CboDrAccount.AddItem (rec1("AccName"))
        Me.CboDrAccount.ItemData(Me.CboDrAccount.NewIndex) = rec1("AccID")
        rec1.MoveNext
    Wend
    If Me.CboDrAccount.ListCount > 0 Then
        Me.CboDrAccount.ListIndex = 0
    End If
End Sub
Private Sub cboGroup_Click()
    cboGroup_Change
End Sub
Private Sub cboGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Me.TxtNarration.SetFocus
    End If
    If KeyCode = 13 Then
        Me.CboDrAccount.SetFocus
    End If
End Sub

Private Sub cmddelete_Click()
    db.Execute ("delete * from temppayment")
    Data1.Refresh
    Me.TxtCrAmount.Text = "0.00"
    VDELETE = "y"
    Me.txtSlNo.Locked = False
    Me.TxtPDate.SetFocus
End Sub

Private Sub cmdedit_Click()
    Me.txtSlNo.Locked = False
    Me.TxtPDate.SetFocus
    db.Execute ("delete * from TempPayment")
    Data1.Refresh
    Me.TxtCrAmount.Text = "0.00"
End Sub

Private Sub cmdprint_Click()
frmpaymentprint.Show 0
End Sub

Private Sub cmdsave_Click()
'-------Change----------
    ans = MsgBox("SAVE THIS?", vbYesNo)
    If ans = 6 Then
        Set rs = db.OpenRecordset("select* from TempPayment")
        If Not rs.EOF Then
            temp_day = Left((Me.TxtPDate.Text), 2)
            temp_month = Mid((Me.TxtPDate.Text), 4, 2)
            temp_year = Right((Me.TxtPDate.Text), 4)
            temp_from = temp_month & "/" & temp_day & "/" & temp_year
            Accperiod_day = Left(AccountingPeriod, 2)
            Accperiod_month = Mid(AccountingPeriod, 4, 2)
            Accperiod_year = Right(AccountingPeriod, 4)


            Set rec = db.OpenRecordset("select * from PaymentDetails where SlNo=" & Me.txtSlNo.Text & " and PDate=#" & temp_from & "#")
            If Not rec.EOF Then
                Set rec5 = db.OpenRecordset("SELECT * FROM PAYMENTHEAD WHERE SlNo=" & Me.txtSlNo.Text & " and PDate=#" & temp_from & "#")
                If Not rec5.EOF Then
                    CrAccId = rec5("AccId")
                End If
                While Not rec.EOF
                    Set rec1 = db.OpenRecordset("select * from LedgertRan where AccId=" & CrAccId & " and VoucherType='Payment' and Tdate=#" & temp_from & "# and TranAccId=" & rec("AccId") & " and VoucherSlno=" & Me.txtSlNo.Text)
                    If Not rec1.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec1("AccId") & " and SlNo>=" & rec1("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Payment' and VoucherSlno=" & Me.txtSlNo.Text & " and TDate =#" & temp_from & "# and TranAccId=" & rec("AccId"))
                    End If
                    Set rec1 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec("AccId") & " and Tdate=#" & temp_from & "# and VoucherType='Payment' and VoucherslNo=" & Me.txtSlNo.Text & " and TranAccId=" & CrAccId)
                    If Not rec1.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec("AccId") & " and SlNo>=" & rec1("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec("AccId") & " and VoucherType='Payment' and VoucherSlno=" & Me.txtSlNo.Text & " and TDate =#" & temp_from & "# and TranAccId=" & CrAccId)
                    End If
                    rec.MoveNext
                Wend
                db.Execute ("delete * from PaymentHead where Slno=" & Me.txtSlNo.Text & " and PDate=#" & temp_from & "#")
                db.Execute ("delete * from PaymentDetails where SlNo=" & Me.txtSlNo.Text & " and PDate=#" & temp_from & "#")
            End If

            '=============New Entry===========
            Set rec = db.OpenRecordset("SELECT * FROM TEMPPAYMENT")
            If Not rec.EOF Then
                db.Execute ("insert into PaymentHead (SlNo,PDate,AccId,AccName,Amount,ParentTrn,ChildTran,Narration) values(" & Me.txtSlNo.Text & ",'" & Me.TxtPDate.Text & "'," & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex) & ",'" & Me.CboCrAccount.Text & "'," & Me.TxtCrAmount.Text & ",'CREDIT','DEBIT','" & Trim(Me.TxtNarration.Text) & "')")
                While Not rec.EOF
                    db.Execute ("insert into PaymentDetails (SlNo,PDate,AccId,AccName,Amount,TranType) values(" & Me.txtSlNo.Text & ",'" & Me.TxtPDate.Text & "'," & rec("AccId") & ",'" & rec("AccName") & "'," & rec("Amount") & ",'" & rec("TranType") & "')")
                    Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex))
                    If Not rec3.EOF Then
                        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex) & " and SlNo=(select max(slno) from LedgerTran where AccID=" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex) & " )")
                        If Not rec4.EOF Then
                            CrAccount_Slno = rec4("SlNo") + 1
                        Else
                            CrAccount_Slno = 1
                        End If
                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & CrAccount_Slno & ",'" & Me.TxtPDate.Text & "','By " & rec("AccName") & "',0," & rec("Amount") & ",0," & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex) & ",'" & Trim(Me.TxtNarration.Text) & "','Payment'," & Me.txtSlNo.Text & "," & rec3("GroupId") & "," & rec("AccId") & ")")
                    End If
                    Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccId=" & rec("AccId"))
                    If Not rec3.EOF Then
                        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec("AccId") & " and SlNo=(select max(slno) from LedgerTran where AccID=" & rec("AccId") & ")")
                        If Not rec4.EOF Then
                            Ledger_slno = rec4("SlNo") + 1
                        Else
                            Ledger_slno = 1
                        End If
                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Ledger_slno & ",'" & Me.TxtPDate.Text & "','By " & Me.CboCrAccount.Text & "'," & rec("Amount") & ",0,0," & rec("AccId") & ",'" & Trim(Me.TxtNarration.Text) & "','Payment'," & Me.txtSlNo.Text & "," & rec3("GroupId") & "," & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex) & ")")
                    End If
                    rec.MoveNext
                Wend
            End If

            db.Execute ("delete * from temppayment")
            Data1.Refresh
            Me.TxtNarration.Text = ""
            Ledger_Balance = 0
            Ledger_slno = 0
            CrAccount_Balance = 0
            CrAccount_Slno = 0
            Me.TxtCrAmount.Text = "0.00"
            Me.txtSlNo.Locked = True
            Me.TxtPDate.SetFocus
        End If
    End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Me.TxtCrAmount.Text = Format(Val(Me.TxtCrAmount.Text) - Val(Me.DBGrid1.Columns(2)), "#########0.00")
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Me.TxtPDate.Text = Format(Date, "dd/mm/yyyy")
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where Groupname like 'Cash-In-Hand' or groupname like 'Bank Accounts'")
    While Not rec1.EOF
        Me.CboCrAccount.AddItem (rec1("AccName"))
        Me.CboCrAccount.ItemData(Me.CboCrAccount.NewIndex) = rec1("AccId")
        rec1.MoveNext
    Wend
    If Me.CboCrAccount.ListCount > 0 Then
        Me.CboCrAccount.ListIndex = 0
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
    frmMain.mnuvoucher.Enabled = False
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("DELETE * FROM TEMPPAYMENT")
    frmMain.mnuvoucher.Enabled = True
End Sub
Private Sub TxtDrAmount_GotFocus()
    Me.TxtDrAmount.SelStart = 0
    Me.TxtDrAmount.SelLength = Len(Me.TxtDrAmount.Text)
End Sub
Private Sub TxtDrAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Val(Me.TxtDrAmount.Text) > 0 Then
            Set rec1 = db.OpenRecordset("select * from TempPayment where AccId=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex))
            If rec1.EOF Then
                Me.TxtCrAmount.Text = Format(Val(Me.TxtCrAmount.Text) + Val(Me.TxtDrAmount.Text), "############0.00")
                db.Execute ("insert into tempPayment (AccId,AccName,Amount,TranType) values(" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ",'" & Me.CboDrAccount.Text & "'," & Me.TxtDrAmount.Text & ",'DEBIT')")
                Data1.Refresh
                Me.TxtDrAmount.Text = "0.00"
                Me.CboDrAccount.ListIndex = 0
                Me.CboGroup.SetFocus
            Else
                MsgBox "Allready Exists", vbCritical
            End If
        Else
            MsgBox "Zero Not Enter", vbCritical
        End If
    End If
End Sub

Private Sub txtNarration_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.CmdSave.SetFocus
    End If
End Sub

Private Sub TxtPDate_GotFocus()
    Me.TxtPDate.SelStart = 0
    Me.TxtPDate.SelLength = Len(Me.TxtPDate.Text)
End Sub
Private Sub TxtPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.TxtPDate.Text), 2)
        temp_month = Mid((Me.TxtPDate.Text), 4, 2)
        temp_year = Right((Me.TxtPDate.Text), 4)

        Accperiod_day = Left(AccountingPeriod, 2)
        Accperiod_month = Mid(AccountingPeriod, 4, 2)
        Accperiod_year = Right(AccountingPeriod, 4)
        Set rec1 = db.OpenRecordset("select max(SlNo) as max_slno from PaymentHead where PDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
        If Not IsNull(rec1!max_slno) Then
            Me.txtSlNo.Text = rec1!max_slno + 1
        Else
            Me.txtSlNo.Text = 1
        End If
        If Me.txtSlNo.Locked = False Then
            Me.txtSlNo.SetFocus
        Else
            Me.CboCrAccount.SetFocus
        End If
    End If
End Sub
Private Sub txtslno_GotFocus()
    Me.txtSlNo.SelStart = 0
    Me.txtSlNo.SelLength = Len(Me.txtSlNo.Text)
End Sub
Private Sub txtslno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.TxtPDate.Text), 2)
        temp_month = Mid((Me.TxtPDate.Text), 4, 2)
        temp_year = Right((Me.TxtPDate.Text), 4)
        db.Execute ("delete * from temppayment")
        Set rec1 = db.OpenRecordset("SELECT * FROM PAYMENTHEAD WHERE PDATE=#" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
        If Not rec1.EOF Then
            Me.TxtCrAmount.Text = Format(rec1("Amount"), "###############0.00")
            Me.TxtNarration.Text = rec1("Narration")
            Set rec2 = db.OpenRecordset("SELECT * FROM LEDGERMASTER WHERE ACCID=" & rec1("aCCiD"))
            If Not rec2.EOF Then
                Me.CboCrAccount.Text = rec2("ACCNAME")
            End If
        End If
        Set rec1 = db.OpenRecordset("select * from Paymentdetails where PDATE=#" & temp_month & "/" & temp_day & "/" & temp_year & "# AND SlNo=" & Me.txtSlNo.Text)
        If Not rec1.EOF Then
            While Not rec1.EOF
                db.Execute ("insert into temppayment (AccId,AccName,Amount,TranType) values(" & rec1("AccId") & ",'" & rec1("AccName") & "'," & rec1("Amount") & ",'" & rec1("TranType") & "')")
                rec1.MoveNext
            Wend
            Data1.Refresh
        End If
        Me.CboGroup.SetFocus


        '---------Deleteing Voucher----------------
        If VDELETE = "y" Then
            ans = MsgBox("Confirm Delete?", vbYesNo)
            If ans = 6 Then
                Set rec1 = db.OpenRecordset("select * from Paymenthead where PDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
                If Not rec1.EOF Then
                    Set rec2 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Payment' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                    If Not rec2.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec2("Dr") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec2("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Payment' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                    End If
                End If
                '--------------Payment Details delete------------
                Set rec1 = db.OpenRecordset("select * from PaymentDetails where PDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
                If Not rec1.EOF Then
                    While Not rec1.EOF
                        Set rec2 = db.OpenRecordset("select * from LedgerMaster Where AccId=" & rec1("AccId"))
                        If Not rec2.EOF Then
                            temp_sign = rec2("Dr")
                            Set rec3 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Payment' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                            If Not rec3.EOF Then
                                db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance " & temp_sign & rec3("Cr") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec3("SlNo"))
                                db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Payment' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtSlNo.Text)
                            End If
                        End If
                        rec1.MoveNext
                    Wend
                End If
                db.Execute ("delete * from Paymenthead where PDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
                db.Execute ("delete * from PaymentDetails where PDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtSlNo.Text)
            End If
            db.Execute ("delete * from temppayment")
            Data1.Refresh
            Me.TxtCrAmount.Text = "0.00"
            Me.TxtPDate.SetFocus
            Me.txtSlNo.Locked = True
            VDELETE = "n"
            Me.txtSlNo.Locked = True
        End If
    End If
End Sub
Private Sub txtSlNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCapture Me.hwnd
End Sub
