VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmContraVoucher 
   BackColor       =   &H00E3E3E3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contra Voucher"
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
   ShowInTaskbar   =   0   'False
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
      Left            =   3840
      TabIndex        =   6
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
      Left            =   2520
      TabIndex        =   5
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   8415
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmContraVoucher.frx":0000
         Height          =   2055
         Left            =   120
         OleObjectBlob   =   "FrmContraVoucher.frx":0014
         TabIndex        =   19
         Top             =   120
         Width           =   8175
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TempContra"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1140
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
         MaxLength       =   49
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2280
         Width           =   8175
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8415
      Begin VB.TextBox TxtSlNo 
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
         TabIndex        =   13
         Text            =   "0"
         Top             =   120
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
         TabIndex        =   1
         Top             =   600
         Width           =   3855
      End
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
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   600
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TxtCDate 
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
      Begin VB.Label Label2 
         Caption         =   "To A/C"
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
         TabIndex        =   20
         Top             =   600
         Width           =   1335
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   120
         Width           =   735
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1320
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
         Top             =   240
         Width           =   3855
      End
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
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "From A/C"
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
         TabIndex        =   21
         Top             =   240
         Width           =   1335
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
         TabIndex        =   10
         Top             =   480
         Width           =   1455
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
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
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
      Left            =   5160
      TabIndex        =   7
      Top             =   6360
      Width           =   975
   End
End
Attribute VB_Name = "FrmContraVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As DAO.Recordset, rec1 As Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset, rec5 As Recordset, CDELETE
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432
Attribute rec5.VB_VarUserMemId = 1073938432
Attribute CDELETE.VB_VarUserMemId = 1073938432
Private Sub cboCrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Me.txtNarration.SetFocus
    End If
    If KeyCode = 13 Then
        Me.TxtCrAmount.SetFocus
    End If
End Sub

Private Sub CboDrAccount_Change()
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where Groupname = 'Cash-In-Hand' or groupname = 'Bank Accounts' and AccId <> " & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex))
    Me.CboCrAccount.Clear
    While Not rec1.EOF
        Me.CboCrAccount.AddItem (rec1("AccName"))
        Me.CboCrAccount.ItemData(Me.CboCrAccount.NewIndex) = rec1("AccId")
        rec1.MoveNext
    Wend
    If Me.CboCrAccount.ListCount > 0 Then
        Me.CboCrAccount.ListIndex = 0
    End If
End Sub
Private Sub CboDrAccount_Click()
    CboDrAccount_Change
End Sub
Private Sub cboDrAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.CboCrAccount.SetFocus
    End If
End Sub
Private Sub CmdDelete_Click()
    db.Execute ("delete * from tempContra")
    Data1.Refresh
    Me.TxtCDate.SetFocus
    CDELETE = "y"
    Me.txtslno.Locked = False
End Sub
Private Sub CmdEdit_Click()
    db.Execute ("delete * from tempcontra")
    Data1.Refresh
    Me.TxtDrAmount.Text = "0.00"
    Me.txtslno.Locked = False
    Me.TxtCDate.SetFocus
End Sub
Private Sub CmdSave_Click()
On Error GoTo errtrap
    temp_day = Left((Me.TxtCDate.Text), 2)
    temp_month = Mid((Me.TxtCDate.Text), 4, 2)
    temp_year = Right((Me.TxtCDate.Text), 4)

    Accperiod_day = Left(AccountingPeriod, 2)
    Accperiod_month = Mid(AccountingPeriod, 4, 2)
    Accperiod_year = Right(AccountingPeriod, 4)
    ans = MsgBox("Confirm This Transaction?", vbYesNo)
    If ans = 6 Then
        Set rec = db.OpenRecordset("select * from TempContra")
        If Not rec.EOF Then
            Set rec1 = db.OpenRecordset("select * from Contrahead where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
            If Not rec1.EOF Then
                Set rec2 = db.OpenRecordset("select * from ledgertran where AccId=" & rec1("AccId") & " and VoucherType='Contra' and VoucherSlno=" & Me.txtslno.Text & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                If Not rec2.EOF Then
                    db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec2("Dr") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec2("SlNo"))
                    db.Execute ("delete * from LedgerTran where AccId=" & rec1("AccId") & " and VoucherType='Contra' and VoucherSlno=" & Me.txtslno.Text & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                End If
                '-------------------DELETING CONTRADETAILS ENTRY--------------
                Set rec2 = db.OpenRecordset("select * from ContraDetails where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
                If Not rec2.EOF Then
                    While Not rec2.EOF
                        Set rec3 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec2("AccId") & " and VoucherType='Contra' and VoucherSlNo=" & Me.txtslno.Text & " and TDate= #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                        If Not rec3.EOF Then
                            db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec2("AccId") & " and SlNo>=" & rec3("SlNo"))
                            db.Execute ("delete * from LedgerTran where AccId=" & rec2("AccId") & " and VoucherType='Contra' and VoucherSlNo=" & Me.txtslno.Text & " and TDate= #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                        End If
                        rec2.MoveNext
                    Wend
                End If
                db.Execute ("Delete * from ContraDetails where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SLno=" & Me.txtslno.Text)
                db.Execute ("Delete * from ContraHead where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SLno=" & Me.txtslno.Text)
            End If
            '------------New Transaction  Head Entry-----------
            db.Execute ("insert into ContraHead (SlNo,CDate,AccId,AccName,Amount,ParentTrn,ChildTran,Narration) values(" & Me.txtslno.Text & ",'" & Me.TxtCDate.Text & "'," & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ",'" & Me.CboDrAccount.Text & "'," & Me.TxtDrAmount.Text & ",'DEBIT','CREDIT','" & Trim(Me.txtNarration.Text) & "')")
            Set rec4 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex))
            If Not rec4.EOF Then
                Set rec3 = db.OpenRecordset("select * from LedgerTran where AccId=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & " and SlNo=(select max(SlNo) from LedgerTran where Accid=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ")")
                If Not rec3.EOF Then
                    temp_LegerSlno = rec3("SlNo") + 1
                    Temp_LedgerBalance = rec3("Balance")
                Else
                    temp_LegerSlno = 1
                    Temp_LedgerBalance = 0
                End If
                db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & temp_LegerSlno & ",'" & Me.TxtCDate.Text & "','By Bank/CAsh Account'," & Me.TxtDrAmount.Text & ",0," & Temp_LedgerBalance + Val(Me.TxtDrAmount.Text) & "," & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex) & ",'" & Trim(Me.txtNarration.Text) & "','Contra'," & Me.txtslno.Text & "," & rec4("GroupId") & ")")
            End If

            '-----Child Transaction Entry---------------
            Set rec5 = db.OpenRecordset("select * from tempcontra")
            While Not rec5.EOF
                db.Execute ("insert into ContraDetails (SlNo,CDate,AccId,AccName,Amount,TranType) values(" & Me.txtslno.Text & ",'" & Me.TxtCDate.Text & "'," & rec5("AccId") & ",'" & rec5("AccName") & "'," & rec5("Amount") & ",'CREDIT')")
                Set rec1 = db.OpenRecordset("select * from LedgerMaster  where AccId=" & rec5("AccId"))
                If Not rec1.EOF Then
                    Set rec2 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec5("AccId") & " and VoucherType='Contra' and VoucherSlNo=" & Me.txtslno.Text & " and TDate= #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                    If Not rec2.EOF Then
                        db.Execute ("update LedgerTran set Cr=Cr+" & rec5("Amount") & " where AccId=" & rec5("AccId") & " and VoucherType='Contra' and VoucherSlNo=" & Me.txtslno.Text & " and TDate= #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                    Else
                        Set rec3 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec5("AccId") & " and SlNo=(select max(SlNo) from LedgerTran where Accid=" & rec5("AccId") & ")")
                        If Not rec3.EOF Then
                            Child_LedgerSlno = rec3("SlNo") + 1
                            Child_LedgerBalance = rec3("Balance")
                        Else
                            Child_LedgerSlno = 1
                            Child_LedgerBalance = 0
                        End If
                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(" & Child_LedgerSlno & ",'" & Me.TxtCDate.Text & "','By Bank/CAsh Account',0," & rec5("Amount") & "," & Child_LedgerBalance + rec5("Amount") & "," & rec5("AccId") & ",'" & Trim(Me.txtNarration.Text) & "','Contra'," & Me.txtslno.Text & "," & rec1("GroupId") & ")")
                    End If
                    rec5.MoveNext
                End If
            Wend

            db.Execute ("delete * from TempContra")
            Data1.Refresh
            Me.TxtDrAmount.Text = "0.00"
            Me.TxtCDate.SetFocus
            Me.txtslno.Locked = True
        End If
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Me.TxtDrAmount.Text = Format(Val(Me.TxtDrAmount.Text) - Val(Me.DBGrid1.Columns(2)), "##########0.00")
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Me.Data1.databasename = dbname
    Me.TxtCDate.Text = Format(Date, "dd/mm/yyyy")
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where Groupname like 'Cash-In-Hand' or groupname like 'Bank Accounts'")
    While Not rec1.EOF
        Me.CboDrAccount.AddItem (rec1("AccName"))
        'Me.CboCrAccount.AddItem (rec1("AccName"))
        Me.CboDrAccount.ItemData(Me.CboDrAccount.NewIndex) = rec1("AccId")
        'Me.CboCrAccount.ItemData(Me.CboCrAccount.NewIndex) = rec1("AccId")
        rec1.MoveNext
    Wend
    If Me.CboDrAccount.ListCount > 0 Then
        Me.CboDrAccount.ListIndex = 0
        'Me.CboCrAccount.ListIndex = 0
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("DELETE * FROM TEMPCONTRA")
End Sub

Private Sub TxtCDate_GotFocus()
    Me.TxtCDate.SelStart = 0
    Me.TxtCDate.SelLength = Len(Me.TxtCDate.Text)
End Sub
Private Sub TxtCDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.TxtCDate.Text), 2)
        temp_month = Mid((Me.TxtCDate.Text), 4, 2)
        temp_year = Right((Me.TxtCDate.Text), 4)

        Accperiod_day = Left(AccountingPeriod, 2)
        Accperiod_month = Mid(AccountingPeriod, 4, 2)
        Accperiod_year = Right(AccountingPeriod, 4)
        Set rec1 = db.OpenRecordset("select max(SlNo) as max_slno from ContraHead where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
        If Not IsNull(rec1!max_slno) Then
            Me.txtslno.Text = rec1!max_slno + 1
        Else
            Me.txtslno.Text = 1
        End If
        If Me.txtslno.Locked = False Then
            Me.txtslno.SetFocus
        Else
            Me.CboDrAccount.SetFocus
        End If
    End If
End Sub

Private Sub TxtCrAmount_GotFocus()
    Me.TxtCrAmount.SelStart = 0
    Me.TxtCrAmount.SelLength = Len(Me.TxtCrAmount.Text)
End Sub
Private Sub TxtCrAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Val(Me.TxtCrAmount.Text) > 0 Then
            Set rec1 = db.OpenRecordset("select * from TempContra where AccId=" & Me.CboDrAccount.ItemData(Me.CboDrAccount.ListIndex))
            If rec1.EOF Then
                Me.TxtDrAmount.Text = Format(Val(Me.TxtDrAmount.Text) + Val(Me.TxtCrAmount.Text), "############0.00")
                db.Execute ("insert into tempContra (AccId,AccName,Amount,TranType) values(" & Me.CboCrAccount.ItemData(Me.CboCrAccount.ListIndex) & ",'" & Me.CboCrAccount.Text & "'," & Me.TxtCrAmount.Text & ",'CREDIT')")
                Data1.Refresh
                Me.TxtCrAmount.Text = "0.00"
                Me.CboCrAccount.ListIndex = 0
                Me.CboCrAccount.SetFocus
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
        Me.cmdSave.SetFocus
    End If
End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtslno_GotFocus()
    Me.txtslno.SelStart = 0
    Me.txtslno.SelLength = Len(Me.txtslno.Text)
End Sub
Private Sub txtSlno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.TxtCDate.Text), 2)
        temp_month = Mid((Me.TxtCDate.Text), 4, 2)
        temp_year = Right((Me.TxtCDate.Text), 4)
        db.Execute ("delete * from TempContra")
        Set rec1 = db.OpenRecordset("select * from Contrahead where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
        If Not rec1.EOF Then
            Set rec2 = db.OpenRecordset("select * from LedgerMaster where AccId=" & rec1("AccId"))
            If Not rec2.EOF Then
                Me.CboDrAccount.Text = rec2("AccName")
            End If
            Me.TxtDrAmount.Text = Format(rec1("Amount"), "#############0.00")
        End If
        Set rec2 = db.OpenRecordset("select * from ContraDetails where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
        If Not rec2.EOF Then
            While Not rec2.EOF
                db.Execute ("insert into TempContra (AccId,AccName,Amount,TranType) values(" & rec2("AccId") & ",'" & rec2("AccName") & "'," & rec2("Amount") & ",'" & rec2("TranType") & "')")
                rec2.MoveNext
            Wend
            Data1.Refresh
        End If
        Me.CboDrAccount.SetFocus

        '-------------Delete--------------

        If CDELETE = "y" Then
            ans = MsgBox("Confirm Delete?", vbYesNo)
            If ans = 6 Then
                Set rec1 = db.OpenRecordset("select * from Contrahead where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
                If Not rec1.EOF Then
                    Set rec2 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Contra' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtslno.Text)
                    If Not rec2.EOF Then
                        db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec2("Dr") & " where AccId=" & rec1("AccId") & " and SlNo>=" & rec2("SlNo"))
                        db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Contra' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtslno.Text)
                    End If
                End If
                '--------------Contra Details delete------------
                Set rec1 = db.OpenRecordset("select * from ContraDetails where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
                If Not rec1.EOF Then
                    While Not rec1.EOF
                        Set rec2 = db.OpenRecordset("select * from LedgerMaster Where AccId=" & rec1("AccId"))
                        If Not rec2.EOF Then
                            Set rec3 = db.OpenRecordset("select * from ledgertran Where AccId=" & rec1("AccId") & " and VoucherType='Contra' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtslno.Text)
                            If Not rec3.EOF Then
                                db.Execute ("update LedgerTran set SLno=SlNo-1 where AccId=" & rec1("AccId") & " and SlNo>=" & rec3("SlNo"))
                                db.Execute ("delete * from LedgerTran Where AccId=" & rec1("AccId") & " and VoucherType='Contra' and  TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and VoucherSlno=" & Me.txtslno.Text)
                            End If
                        End If
                        rec1.MoveNext
                    Wend
                End If
                db.Execute ("delete * from Contrahead where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
                db.Execute ("delete * from ContraDetails where CDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "# and SlNo=" & Me.txtslno.Text)
            End If
            db.Execute ("delete * from tempContra")
            Data1.Refresh
            Me.TxtDrAmount.Text = "0.00"
            Me.TxtCDate.SetFocus
            Me.txtslno.Locked = True
            CDELETE = "n"
            Me.txtslno.Locked = True
        End If
    End If
End Sub
