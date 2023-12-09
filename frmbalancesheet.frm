VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmbalancesheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sheet"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10950
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1680
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   6480
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmbalancesheet.frx":0000
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmbalancesheet.frx":0014
      TabIndex        =   0
      Top             =   720
      Width           =   10695
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
      RecordSource    =   "Balancesheet"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "A S S E T S"
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
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "L I A B I L I T I E S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label2 
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
      Left            =   8280
      TabIndex        =   2
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label1 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   6120
      Width           =   2535
   End
End
Attribute VB_Name = "frmbalancesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As DAO.Recordset, rec As Recordset, rec1 As Recordset, rec2 As Recordset
Attribute rec.VB_VarUserMemId = 1073938432
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432

Private Sub Command1_Click()
    Me.CrystalReport1.ReportFileName = App.Path & "\Balancesheet.rpt"
    Me.CrystalReport1.PrintReport
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    db.Execute ("delete * from Balancesheet")
    Me.Top = 0
    Me.Left = 0
    Dim Total_Dr As Currency, temp_cr As Currency, temp_dr As Currency, Total_Cr As Currency
    Me.Top = 0
    Me.Left = 0
    temp_slno = 1
    cr_slno = 0
    Set rs = db.OpenRecordset("select * from groups where GP_On_Bs='y' and GroupNature='Liabilites' order by Bl_Pref ASC")
    While Not rs.EOF

        Set rec = db.OpenRecordset("select * from ledgermaster where groupid=" & rs("groupid"))
        While Not rec.EOF

            Set rec1 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Dr) Then
                temp_dr = rec1!Total_Dr
            Else
                temp_dr = 0
            End If
            Set rec1 = db.OpenRecordset("select sum(cr) as total_cr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Cr) Then
                temp_cr = rec1!Total_Cr
            Else
                temp_cr = 0
            End If

            If rec("oBalance") > 0 And rec("Balancetype") = "Dr" Then
                temp_dr = temp_dr + rec("OBalance")
            End If
            If rec("oBalance") > 0 And rec("Balancetype") = "Cr" Then
                temp_cr = temp_cr + rec("OBalance")
            End If
            If temp_dr > temp_cr Then
                TEMP_DR_BALANCE = temp_dr - temp_cr
                Total_Dr = Total_Dr + TEMP_DR_BALANCE
            End If
            If temp_cr > temp_dr Then
                TEMP_CR_BALANCE = temp_cr - temp_dr
                Total_Cr = Total_Cr + TEMP_CR_BALANCE
            End If
            If temp_dr = temp_cr Then
                temp_dr = 0
                temp_cr = 0
            End If

            rec.MoveNext
        Wend
        '    If Total_Cr > Total_Dr Then
        Cr_balance = Total_Cr
        '    End If
        '    If Total_Dr > Total_Cr Then
        dr_balance = Total_Dr
        '    End If

        If Cr_balance > 0 Then
            db.Execute ("insert into Balancesheet (SlNo,Lib_Particulars,LibAmount) values(" & temp_slno & ",'" & UCase(rs("groupname")) & "'," & Total_Cr & ")")
            temp_slno = temp_slno + 1
        End If
        If dr_balance > 0 Then
            db.Execute ("insert into Balancesheet (SlNo,Ast_Particulars,AstAmount) values(" & temp_slno & ",'" & UCase(rs("groupname")) & "'," & Total_Dr & ")")
            temp_slno = temp_slno + 1
        End If

        If Total_Dr = Total_Cr Then

        End If
        Cr_balance = 0
        dr_balance = 0
        Total_Dr = 0
        Total_Cr = 0
        temp_dr = 0
        temp_cr = 0
        X = 0
        rs.MoveNext

    Wend

    '---------------------without grouping of ledger--------------------------------------------------------
    Set rs = db.OpenRecordset("select * from groups where GP_On_BS='n' and GroupNature='Liabilites' order by Bl_Pref ASC")
    While Not rs.EOF
        Set rec = db.OpenRecordset("select * from ledgermaster where groupid=" & rs("groupid"))
        While Not rec.EOF
            Set rec1 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Dr) Then
                temp_dr = rec1!Total_Dr
            Else
                temp_dr = 0
            End If
            Set rec1 = db.OpenRecordset("select sum(cr) as total_cr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Cr) Then
                temp_cr = rec1!Total_Cr
            Else
                temp_cr = 0
            End If
            If rec("oBalance") > 0 And rec("Balancetype") = "Dr" Then
                temp_dr = temp_dr + rec("OBalance")
            End If
            If rec("oBalance") > 0 And rec("Balancetype") = "Cr" Then
                temp_cr = temp_cr + rec("OBalance")
            End If
            If temp_dr > temp_cr Then
                db.Execute ("insert into Balancesheet (SlNo,Ast_Particulars,AstAmount) values(" & temp_slno & ",'" & UCase(rec("ACCNAME")) & "'," & temp_dr - temp_cr & ")")
                temp_slno = temp_slno + 1
            End If
            If temp_cr > temp_dr Then
                db.Execute ("insert into Balancesheet (slno,Lib_Particulars,LibAmount) values(" & temp_slno & ",'" & UCase(rec("accname")) & "'," & temp_cr - temp_dr & ")")
                temp_slno = temp_slno + 1
            End If

            temp_dr = 0
            temp_cr = 0
            rec.MoveNext
        Wend

        Total_Dr = 0
        Total_Cr = 0
        rs.MoveNext
    Wend
    Total_Dr = 0
    Total_Cr = 0
    temp_dr = 0
    temp_cr = 0
    max_slno = temp_slno
    temp_slno = 1
    '------------------------------Asset-----------------------------

    Set rs = db.OpenRecordset("select * from groups where GP_On_Bs='y' and GroupNature='Asset' order by Bl_Pref ASC")
    While Not rs.EOF

        Set rec = db.OpenRecordset("select * from ledgermaster where groupid=" & rs("groupid"))
        While Not rec.EOF

            Set rec1 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Dr) Then
                temp_dr = rec1!Total_Dr
            Else
                temp_dr = 0
            End If
            Set rec1 = db.OpenRecordset("select sum(cr) as total_cr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Cr) Then
                temp_cr = rec1!Total_Cr
            Else
                temp_cr = 0
            End If

            If rec("oBalance") > 0 And rec("Balancetype") = "Dr" Then
                temp_dr = temp_dr + rec("OBalance")
            End If
            If rec("oBalance") > 0 And rec("Balancetype") = "Cr" Then
                temp_cr = temp_cr + rec("OBalance")
            End If
            If temp_dr > temp_cr Then
                TEMP_DR_BALANCE = temp_dr - temp_cr
                Total_Dr = Total_Dr + TEMP_DR_BALANCE
            End If
            If temp_cr > temp_dr Then
                TEMP_CR_BALANCE = temp_cr - temp_dr
                Total_Cr = Total_Cr + TEMP_CR_BALANCE
            End If
            If temp_dr = temp_cr Then
                temp_dr = 0
                temp_cr = 0
            End If

            rec.MoveNext
        Wend
        '    If Total_Cr > Total_Dr Then
        Cr_balance = Total_Cr
        '    End If
        '    If Total_Dr > Total_Cr Then
        dr_balance = Total_Dr
        '    End If

        If dr_balance > 0 Then
            'If temp_slno < max_slno Then
            'Db.Execute ("update Balancesheet set Ast_Particulars='" & UCase(rs("groupname")) & "',AstAmount=" & Total_Dr & " where SlNo=" & temp_slno)
            'temp_slno = temp_slno + 1
            'Else
            If rs("groupname") = "Stock-In-Hand" Then
                'If temp_slno >= max_slno Then
                db.Execute ("insert into Balancesheet (SlNo,Ast_Particulars,AstAmount) values(" & temp_slno & ",'" & UCase(rs("groupname")) & "'," & Total_Cr & ")")
                'Else
                'Db.Execute ("update Balancesheet set Ast_Particulars='" & UCase(rs("groupname")) & "',AstAmount= " & Total_Cr & " where slno=" & temp_slno)
                'temp_slno = temp_slno + 1
                'End If
            Else
                'If temp_slno >= max_slno Then
                db.Execute ("insert into Balancesheet (SlNo,Ast_Particulars,AstAmount) values(" & temp_slno & ",'" & UCase(rs("groupname")) & "'," & Total_Dr & ")")
                'temp_slno = temp_slno + 1
                ' Else
                ' Db.Execute ("update Balancesheet set Ast_Particulars='" & UCase(rs("groupname")) & "',AstAmount= " & Total_Dr & " where slno=" & temp_slno)
                'temp_slno = temp_slno + 1
                'End If
                'temp_slno = temp_slno + 1
            End If
        End If
        If Cr_balance > 0 Then
            If rs("groupname") <> "Stock-In-Hand" Then
                db.Execute ("insert into Balancesheet (SlNo,Lib_Particulars,LibAmount) values(" & temp_slno & ",'" & UCase(rs("groupname")) & "'," & Total_Cr & ")")
                temp_slno = temp_slno + 1
            End If
        End If
        If Total_Dr = Total_Cr Then

        End If
        Cr_balance = 0
        dr_balance = 0
        Total_Dr = 0
        Total_Cr = 0
        temp_dr = 0
        temp_cr = 0
        X = 0
        rs.MoveNext

    Wend

    '---------------------without grouping of ledger--------------------------------------------------------
    Set rs = db.OpenRecordset("select * from groups where GP_On_BS='n' and GroupNature='Asset' order by Bl_Pref ASC")
    While Not rs.EOF
        Set rec = db.OpenRecordset("select * from ledgermaster where groupid=" & rs("groupid"))
        While Not rec.EOF
            Set rec1 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Dr) Then
                temp_dr = rec1!Total_Dr
            Else
                temp_dr = 0
            End If
            Set rec1 = db.OpenRecordset("select sum(cr) as total_cr from ledgertran where accid=" & rec("accid"))
            If Not IsNull(rec1!Total_Cr) Then
                temp_cr = rec1!Total_Cr
            Else
                temp_cr = 0
            End If
            If rec("oBalance") > 0 And rec("Balancetype") = "Dr" Then
                temp_dr = temp_dr + rec("OBalance")
            End If
            If rec("oBalance") > 0 And rec("Balancetype") = "Cr" Then
                temp_cr = temp_cr + rec("OBalance")
            End If
            If temp_dr > temp_cr Then
                'If temp_slno < max_slno Then
                'Db.Execute ("update Balancesheet set Ast_Particulars='" & UCase(rec("AccName")) & "',AstAmount=" & temp_dr - temp_cr & " where SlNo=" & temp_slno)
                'temp_slno = temp_slno + 1
                'Else
                db.Execute ("insert into Balancesheet (Slno,Ast_Particulars,AstAmount) values(" & temp_slno & ",'" & UCase(rec("accname")) & "'," & temp_dr - temp_cr & ")")
                temp_slno = temp_slno + 1
                'End If
            End If
            If temp_cr > temp_dr Then
                db.Execute ("insert into Balancesheet (SlNo,Lib_Particulars,LibAmount) values(" & temp_slno & ",'" & UCase(rec("accname")) & "'," & temp_cr - temp_dr & ")")
                temp_slno = temp_slno + 1
            End If

            If temp_dr = temp_cr Then

            End If
            temp_dr = 0
            temp_cr = 0
            rec.MoveNext
        Wend

        Total_Dr = 0
        Total_Cr = 0
        rs.MoveNext
    Wend
    Set rec1 = db.OpenRecordset("select * from P_LAcc where Dr_Particulars='NET PROFIT'")
    If Not rec1.EOF Then

        db.Execute ("insert into Balancesheet (Slno,lib_Particulars,libAmount) values(" & temp_slno & ",'NET PROFIT'," & rec1("DRAMOUNT") & ")")
        'Db.Execute ("update Balancesheet set libAmount=libAmount+" & rec1("DRAMOUNT") & " where lib_particulars='PROPRIETORS CAPITAL ACCOUNT'")
    End If
    '//Net loss---------------
    Set rec1 = db.OpenRecordset("select * from P_LAcc where Cr_Particulars='NET LOSS'")
    If Not rec1.EOF Then
        db.Execute ("insert into Balancesheet (SlNo,Ast_Particulars,AstAmount) values(" & temp_slno & ",'NET LOSS'," & rec1("CRAMOUNT") & ")")
        'Db.Execute ("update Balancesheet set libAmount=libAmount-" & rec1("DRAMOUNT") & " where lib_particulars='PROPRIETORS CAPITAL ACCOUNT'")
    End If

    Set rs = db.OpenRecordset("select sum(LibAmount) as trail_dr from Balancesheet")
    If Not IsNull(rs!trail_dr) Then
        Me.Label1.Caption = Format(rs!trail_dr, "########0.00")
    End If
    Set rs = db.OpenRecordset("select sum(AstAmount) as trail_cr from Balancesheet")
    If Not IsNull(rs!trail_Cr) Then
        Me.Label2.Caption = Format(rs!trail_Cr, "########0.00")
    End If

    '-----------USING OPENING BALANCE ----------------------------------------------
    'Set rs = Db.OpenRecordset("SELECT SUM(OBALANCE) AS SUM_DR_BALANCE FROM LEDGERMASTER WHERE BALANCETYPE='Dr'")
    'If Not IsNull(rs!SUM_DR_BALANCE) Then
    '    TEMP_OP_DR_BALANCE = rs!SUM_DR_BALANCE
    'Else
    '    TEMP_OP_DR_BALANCE = 0
    'End If
    '
    'Set rs = Db.OpenRecordset("SELECT SUM(OBALANCE) AS SUM_CR_BALANCE FROM LEDGERMASTER WHERE BALANCETYPE='Cr'")
    'If Not IsNull(rs!SUM_CR_BALANCE) Then
    '    TEMP_OP_CR_BALANCE = rs!SUM_CR_BALANCE
    'Else
    '    TEMP_OP_CR_BALANCE = 0
    'End If
    '
    'If TEMP_OP_DR_BALANCE <> TEMP_OP_CR_BALANCE Then
    '    opening_balance_diff = Abs(TEMP_OP_DR_BALANCE - TEMP_OP_CR_BALANCE)
    '
    '    Set rs = Db.OpenRecordset("select sum(dr) as trail_dr from trailbalance")
    '    If Not IsNull(rs!trail_dr) Then
    '        TRAIL_DR_OP = rs!trail_dr
    '    End If
    '    Set rs = Db.OpenRecordset("select sum(cr) as trail_cr from trailbalance")
    '    If Not IsNull(rs!trail_Cr) Then
    '        TRAIL_CR_OP = rs!trail_Cr
    '    End If
    '    If TRAIL_DR_OP > TRAIL_CR_OP Then
    '        Db.Execute ("insert into trailbalance (AccName,Cr) values('Difference In Opening Balances'," & Abs(TRAIL_DR_OP - TRAIL_CR_OP) & ")")
    '    End If
    '    If TRAIL_CR_OP > TRAIL_DR_OP Then
    '        Db.Execute ("insert into trailbalance (AccName,Dr) values('Difference In Opening Balances'," & Abs(TRAIL_DR_OP - TRAIL_CR_OP) & ")")
    '    End If
    'End If
    '
    ''If TEMP_OP_DR_BALANCE > TEMP_OP_CR_BALANCE Then
    ''    db.Execute ("insert into trailbalance (AccName,dr) values('Difference In Opening Balances'," & TEMP_OP_DR_BALANCE - TEMP_OP_CR_BALANCE & ")")
    ''End If
    ''If TEMP_OP_CR_BALANCE > TEMP_OP_DR_BALANCE Then
    ''    db.Execute ("insert into trailbalance (AccName,cr) values('Difference In Opening Balances'," & TEMP_OP_CR_BALANCE - TEMP_OP_DR_BALANCE & ")")
    ''End If
    '
    ''----------------------------------------------------------------------
    'Set rs = Db.OpenRecordset("select sum(dr) as trail_dr from trailbalance")
    'If Not IsNull(rs!trail_dr) Then
    '    Me.Label1.Caption = Format(rs!trail_dr, "########0.00")
    'End If
    'Set rs = Db.OpenRecordset("select sum(cr) as trail_cr from trailbalance")
    'If Not IsNull(rs!trail_Cr) Then
    '    Me.Label2.Caption = Format(rs!trail_Cr, "########0.00")
    'End If
    Data1.databasename = dbname
    Data1.Refresh
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

