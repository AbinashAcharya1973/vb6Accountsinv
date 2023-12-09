VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmJournal 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Entry"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8160
   Begin VB.TextBox txtNarration 
      BackColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   7935
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmJournal.frx":0000
         Height          =   3615
         Left            =   120
         OleObjectBlob   =   "frmJournal.frx":0014
         TabIndex        =   7
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tempjournal"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      Begin MSMask.MaskEdBox txtdate 
         Height          =   495
         Left            =   6120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtslno 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sl. No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Narration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dramount, cramount, balance, temp_journal_no, rec1 As DAO.Recordset, rec2 As DAO.Recordset, rec3 As DAO.Recordset
Attribute cramount.VB_VarUserMemId = 1073938432
Attribute balance.VB_VarUserMemId = 1073938432
Attribute temp_journal_no.VB_VarUserMemId = 1073938432
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Private Sub CmdSave_Click()
    temp_day = Left((Me.txtdate.Text), 2)
    temp_month = Mid((Me.txtdate.Text), 4, 2)
    temp_year = Right((Me.txtdate.Text), 4)

    Accperiod_day = Left(AccountingPeriod, 2)
    Accperiod_month = Mid(AccountingPeriod, 4, 2)
    Accperiod_year = Right(AccountingPeriod, 4)
    If Me.TxtNarration.Text <> "" Then
        ans = MsgBox("Save the Journal?", vbYesNo)
        If ans = 6 Then
            temp_sq = 0
            Set rs = db.OpenRecordset("select * from tempjournal")
            While Not rs.EOF
                Set rec1 = db.OpenRecordset("select * from ledgermaster where accname='" & rs("AccNAme") & "'")
                temp_id = rec1("accid")

                db.Execute ("insert into JournalTr (Slno,TDate,Particulars,Dr,Cr,AccID,Remarks,Sub_Sq) values(" & Me.txtslno.Text & ",'" & Me.txtdate.Text & "','" & rs("AccName") & "'," & rs("Dr") & "," & rs("Cr") & "," & rec1("AccID") & ",'" & Me.TxtNarration.Text & "'," & temp_sq & ")")
                '----------If problem Change it----------------
                Set rec3 = db.OpenRecordset("select * from ledgertran where accid=" & rec1("accid") & " and slno=(select max(slno) from ledgertran where accid=" & rec1("accid") & " and TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "#)")
                If Not rec3.EOF Then
                    tempslno = rec3("slno") + 1
                    tempbalance = rec3("balance")
                Else
                    tempslno = 1
                    tempbalance = 0
                End If
                If rs("dr") <> 0 Then
                    tempsign = Trim(rec1("dr"))
                    Set rec2 = db.OpenRecordset("select * from tempjournal where cr>0")
                    If rec2("cr") >= rs("dr") Then
                        db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance " & tempsign & rs("dr") & " where AccId=" & temp_id & " and SlNo > " & tempslno - 1)
                        db.Execute ("insert into ledgertran (slno,tdate,particulars,dr,cr,balance,accid,vouchertype,voucherslno,Remarks) values(" & tempslno & ",'" & Me.txtdate.Text & "','To " & rec2("accname") & "'," & rs("dr") & ",0," & tempbalance & tempsign & rs("dr") & "," & temp_id & ",'Journal Voucher'," & Me.txtslno.Text & ",'" & Me.TxtNarration.Text & "')")
                    Else
                        While Not rec2.EOF
                            db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance " & tempsign & rs("Cr") & " where AccId=" & temp_id & " and SlNo > " & tempslno - 1)
                            db.Execute ("insert into ledgertran (slno,tdate,particulars,dr,cr,balance,accid,vouchertype,voucherslno,Remarks) values(" & tempslno & ",'" & Me.txtdate.Text & "','To " & rec2("accname") & "'," & rec2("cr") & ",0," & tempbalance & tempsign & rec2("cr") & "," & temp_id & ",'Journal Voucher'," & Me.txtslno.Text & ",'" & Me.TxtNarration.Text & "')")
                            tempslno = tempslno + 1
                            If tempsign = "+" Then
                                tempbalance = tempbalance + rec2("cr")
                            End If
                            If tempsign = "-" Then
                                tempbalance = tempbalance - rec2("cr")
                            End If
                            rec2.MoveNext
                        Wend
                    End If
                End If
                If rs("cr") <> 0 Then
                    tempsign = Trim(rec1("cr"))
                    Set rec2 = db.OpenRecordset("select * from tempjournal where dr>0")
                    If rec2("dr") >= rs("cr") Then
                        db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance " & tempsign & rs("Cr") & " where AccId=" & temp_id & " and SlNo > " & tempslno - 1)
                        db.Execute ("insert into ledgertran (slno,tdate,particulars,dr,cr,balance,accid,vouchertype,voucherslno,Remarks) values(" & tempslno & ",'" & Me.txtdate.Text & "','By " & rec2("accname") & "',0," & rs("cr") & "," & tempbalance & tempsign & rs("cr") & "," & temp_id & ",'Journal Voucher'," & Me.txtslno.Text & ",'" & Me.TxtNarration.Text & "')")
                    Else
                        While Not rec2.EOF
                            db.Execute ("update LedgerTran set SlNo=SlNo+1,Balance=Balance " & tempsign & rs("Dr") & " where AccId=" & temp_id & " and SlNo > " & tempslno - 1)
                            db.Execute ("insert into ledgertran (slno,tdate,particulars,dr,cr,balance,accid,vouchertype,voucherslno,Remarks) values(" & tempslno & ",'" & Me.txtdate.Text & "','By " & rec2("accname") & "',0," & rec2("dr") & "," & tempbalance & tempsign & rec2("dr") & "," & temp_id & ",'Journal Voucher'," & Me.txtslno.Text & ",'" & Me.TxtNarration.Text & "')")
                            tempslno = tempslno + 1
                            If tempsign = "+" Then
                                tempbalance = tempbalance + rec2("dr")
                            End If
                            If tempsign = "-" Then
                                tempbalance = tempbalance - rec2("dr")
                            End If
                            rec2.MoveNext
                        Wend
                    End If
                End If
                temp_sq = temp_sq + 1
                rs.MoveNext
            Wend
            Me.txtslno.Text = Val(Me.txtslno.Text) + 1
            db.Execute ("delete * from tempjournal")
            dramount = 0
            cramount = 0
            balance = 0
            Data1.Refresh
        End If
    End If
    Me.TxtNarration.Text = ""
    Me.TxtNarration.Locked = True
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If DBGrid1.Col = 1 Then
            frmacclist.Show 0
            If dramount <> cramount Then
                If dramount > cramount Then
                    DBGrid1.Columns(3) = Abs(balance)
                End If
                If dramount < cramount Then
                    DBGrid1.Columns(2) = Abs(balance)
                End If
            End If
        End If

    End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
    If DBGrid1.Col = 0 Then
        If Chr(KeyAscii) = "B" Or Chr(KeyAscii) = "b" Then
            DBGrid1.Columns(0) = "By"

        End If

    End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If LastCol = 2 Then
        dramount = dramount + Val(DBGrid1.Columns(2))
        balance = dramount - cramount
    End If
    If LastCol = 3 Then
        cramount = cramount + Val(DBGrid1.Columns(3))
        balance = dramount - cramount
    End If
    If LastCol = 4 Then
        If dramount = cramount Then
            Me.TxtNarration.Locked = False
            Me.TxtNarration.SetFocus
        End If

    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    temp_date = Format(Date, "dd")
    temp_month = Format(Date, "mm")
    temp_year = Format(Date, "yyyy")
    Me.txtdate.Text = temp_date & "/" & temp_month & "/" & temp_year
    Set rs = db.OpenRecordset("select max(slno) as max_slno from JournalTr where Tdate=#" & temp_month & "/" & temp_date & "/" & temp_year & "#")
    If Not IsNull(rs!max_slno) Then
        Me.txtslno.Text = rs!max_slno + 1
    Else
        Me.txtslno.Text = 1
    End If
    Me.Data1.databasename = dbname
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("delete * from tempjournal")
End Sub

Private Sub txtdate_GotFocus()
    Me.txtdate.SelStart = 0
    Me.txtdate.SelLength = Len(Me.txtdate.Text)
End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.txtdate.Text), 2)
        temp_month = Mid((Me.txtdate.Text), 4, 2)
        temp_year = Right((Me.txtdate.Text), 4)

        Accperiod_day = Left(AccountingPeriod, 2)
        Accperiod_month = Mid(AccountingPeriod, 4, 2)
        Accperiod_year = Right(AccountingPeriod, 4)
        Set rec1 = db.OpenRecordset("select Max(Slno) as max_slno from JournalTr where TDate = #" & temp_month & "/" & temp_day & "/" & temp_year & "#")
        If Not IsNull(rec1!max_slno) Then
            Me.txtslno.Text = rec1!max_slno + 1
        Else
            Me.txtslno.Text = 1
        End If
        Me.DBGrid1.SetFocus
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
