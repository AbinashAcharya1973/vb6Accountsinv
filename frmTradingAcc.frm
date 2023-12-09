VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmTradingAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trading Account"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   10590
   Begin VB.CommandButton cmdplacc 
      Appearance      =   0  'Flat
      Caption         =   "Next to P L Account->"
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
      Left            =   7920
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1680
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print the Statement"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   7080
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmTradingAcc.frx":0000
      Height          =   6135
      Left            =   120
      OleObjectBlob   =   "frmTradingAcc.frx":0014
      TabIndex        =   0
      Top             =   720
      Width           =   10455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TradingAc"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSMask.MaskEdBox txtTo 
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   120
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
   Begin MSMask.MaskEdBox txtFrom 
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   120
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
   Begin VB.Label Label3 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "frmTradingAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As DAO.Recordset, rec3 As DAO.Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432

Private Sub cmdplacc_Click()
    frmPLAcc.Show 0
End Sub

Private Sub cmdprint_Click()
    db.Execute ("update tradingac set fromdt='" & Me.txtfrom.Text & "',todt='" & Me.txtto.Text & "'")
    CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Dim opstock, clstock, purchase, sale, preturn, salereturn
    Me.CrystalReport1.ReportFileName = App.Path & "\trdacc.rpt"
    db.Execute ("DELETE * FROM TRADINGAC")
    opstock = 0
    clstock = 0
    purchase = 0
    preturn = 0
    sale = 0
    salereturn = 0
    '//--OPENING STOCK-------------------
    Set rec1 = db.OpenRecordset("select * from ledgermaster where accname like 'OPENING STOCK*'")
    If Not rec1.EOF Then
        Set rec2 = db.OpenRecordset("SELECT SUM(DR) AS OSTOCK FROM LEDGERTRAN WHERE ACCID=" & rec1("ACCID"))
        If Not IsNull(rec2!OSTOCK) Then
            opstock = rec1("obalance") + rec2!OSTOCK
        Else
            opstock = rec1("obalance")
        End If
    End If
    '//-----------CLOSING STOCK---------
    Set rec1 = db.OpenRecordset("select * from ledgermaster where accname like 'CLOSING STOCK*'")
    If Not rec1.EOF Then
        Set rec2 = db.OpenRecordset("SELECT SUM(CR) AS CSTOCK FROM LEDGERTRAN WHERE ACCID=" & rec1("ACCID"))
        If Not IsNull(rec2!CSTOCK) Then
            clstock = rec2!CSTOCK
        Else
            clstock = 0
        End If
    End If
    '//-----------------
    Set rec1 = db.OpenRecordset("select * from ledgermaster where accname like 'Purchase*'")
    If Not rec1.EOF Then

        Set rec2 = db.OpenRecordset("select sum(dr) as dr_total from ledgertran where accid=" & rec1("accid"))
        If Not IsNull(rec2!dr_total) Then
            purchase = rec2!dr_total
        End If
    End If
    Set rec1 = db.OpenRecordset("select * from ledgermaster where accname like 'PURCHASE RETURN*'")
    If Not rec1.EOF Then
        '----------------PURCHASE RETURN--------------------------------------
        Set rec2 = db.OpenRecordset("select sum(cr) as cr_total from ledgertran where accid=" & rec1("accid"))
        If Not IsNull(rec2!cr_total) Then
            preturn = rec2!cr_total
        End If

        If rec1("obalance") > 0 And rec1("balancetype") = "Dr" Then
            purchase = purchase + rec1("obalance")
        End If
        If rec1("obalance") > 0 And rec1("balancetype") = "Cr" Then
            preturn = preturn + rec1("obalance")
        End If

    End If

    Set rec1 = db.OpenRecordset("select * from ledgermaster where accname like 'SALES ACCOUNT'")
    If Not rec1.EOF Then
        '-------------SALES -------------------------------------------------------
        Set rec2 = db.OpenRecordset("select sum(dr) as dr_total from ledgertran where accid=" & rec1("accid"))
        If Not IsNull(rec2!dr_total) Then
            saleretun = rec2!dr_total
        End If

        Set rec2 = db.OpenRecordset("select sum(cr) as cr_total from ledgertran where accid=" & rec1("accid"))
        If Not IsNull(rec2!cr_total) Then
            sale = rec2!cr_total - saleretun
        End If

        If rec1("obalance") > 0 And rec1("balancetype") = "Dr" Then
            salereturn = salereturn + rec1("obalance")
        End If
        If rec1("obalance") > 0 And rec1("balancetype") = "Cr" Then
            sale = sale + rec1("obalance")
        End If
    End If

    Set rec1 = db.OpenRecordset("select * from ledgermaster where accname like 'SALES RETURN*'")
    If Not rec1.EOF Then
        '-------------SALES RETURN-------------------------------------------------------
        Set rec2 = db.OpenRecordset("select sum(dr) as dr_total from ledgertran where accid=" & rec1("accid"))
        If Not IsNull(rec2!dr_total) Then
            saleretun = rec2!dr_total
        End If

        'Set rec2 = Db.OpenRecordset("select sum(cr) as cr_total from ledgertran where accid=" & rec1("accid"))
        'If Not IsNull(rec2!cr_total) Then
        ' sale = rec2!cr_total
        'End If

        'If rec1("obalance") > 0 And rec1("balancetype") = "Dr" Then
        saleretun = saleretun + rec1("obalance")
        ' End If
        'If rec1("obalance") > 0 And rec1("balancetype") = "Cr" Then
        '    sale = sale + rec1("obalance")
        'End If

    End If

    ''------------------INSERTING RECORDS INTO TRADING ACCOUNT-------------------
    ''---1. OPENING STOCK--------
    'Db.Execute ("INSERT INTO TRADINGAC (slno,DR_PARTICULARS,DrAmount) values(1,'Opening Stock'," & opstock & ")")
    ''---2. PURCHASE ------------
    'Db.Execute ("INSERT INTO TRADINGAC (slno,DR_PARTICULARS,DrAmt,DrAmount) values(2,'Purchase'," & purchase & "," & purchase - preturn & ")")
    ''---3. PURCHASE RETURN------
    'Db.Execute ("INSERT INTO TRADINGAC (slno,CR_PARTICULARS,CrAmt,CrAmount) values(3,'Purchase Return'," & preturn & "," & preturn & ")")
    '
    ''---1. SALES------
    'Db.Execute ("UPDATE TRADINGAC SET CR_PARTICULARS='Sales',CrAmt=" & sale & ",CrAmount=" & sale & " where slno=1")
    ''---2. SALES RETURN------
    ''Db.Execute ("INSERT INTO TRADINGAC (slno,DR_PARTICULARS,DrAmt,DrAmount) values(3,'SALES RETURN'," & saleretun & "," & saleretun & ")")
    'Db.Execute ("UPDATE TRADINGAC SET DR_PARTICULARS='Sales Return',DrAmt=" & saleretun & ",DrAmount=" & saleretun & " where slno=3")
    ''---3. CLOSING STOCK------
    'Db.Execute ("UPDATE TRADINGAC SET CR_PARTICULARS='Closing Stock',CrAmount=" & clstock & " where slno=2")
    ''-----OTHER EXPENCES-----------------------
    '----------------------------------------
    '------------------INSERTING RECORDS INTO TRADING ACCOUNT-------------------
    '---1. OPENING STOCK--------
    db.Execute ("INSERT INTO TRADINGAC (slno,DR_PARTICULARS,DrAmount) values(1,'Opening Stock'," & opstock & ")")
    '---2. PURCHASE ------------
    db.Execute ("INSERT INTO TRADINGAC (slno,DR_PARTICULARS,DrAmt) values(2,'Purchase'," & purchase & ")")
    '---3. PURCHASE RETURN------
    db.Execute ("INSERT INTO TRADINGAC (slno,DR_PARTICULARS,DrAmt,DrAmount) values(3,'Less Purchase Return'," & preturn & "," & purchase - preturn & ")")
    '---1. SALES------
    db.Execute ("UPDATE TRADINGAC SET CR_PARTICULARS='Sales',CrAmt=" & sale & " where slno=1")
    '---2. SALES RETURN------
    db.Execute ("UPDATE TRADINGAC SET CR_PARTICULARS='Less Sales Return',CrAmt=" & saleretun & ",CrAmount=" & sale - saleretun & " where slno=2")
    '---3. CLOSING STOCK------
    db.Execute ("UPDATE TRADINGAC SET CR_PARTICULARS='Closing Stock',CrAmount=" & clstock & " where slno=3")
    '-----OTHER EXPENCES-----------------------
    '-----------------------------------------
    tempslno = 4
    Set rec1 = db.OpenRecordset("select * from groups where GroupNature='Expences' and Affect_GP='Y'")
    While Not rec1.EOF
        Set rec2 = db.OpenRecordset("select * from ledgermaster where groupid=" & rec1("groupid"))
        While Not rec2.EOF
            temp_dr = 0
            temp_cr = 0
            Set rec3 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rec2("accid"))
            If Not IsNull(rec3!Total_Dr) Then
                temp_dr = rec3!Total_Dr
            End If
            Set rec3 = db.OpenRecordset("select sum(cr) as total_cr from ledgertran where accid=" & rec2("accid"))
            If Not IsNull(rec3!Total_Cr) Then
                temp_cr = rec3!Total_Cr
            End If

            If rec2("obalance") > 0 And rec2("balancetype") = "Dr" Then
                temp_dr = temp_dr + rec2("obalance")
            End If
            If rec2("obalance") > 0 And rec2("balancetype") = "Cr" Then
                temp_cr = temp_cr + rec2("obalance")
            End If
            If temp_dr > temp_cr Then
                db.Execute ("insert into tradingac (slno,DR_PARTICULARS,DRAMOUNT) VALUES(" & tempslno & ",'" & rec2("AccName") & "'," & temp_dr - temp_cr & ")")
            End If
            If temp_cr > temp_dr Then
                db.Execute ("insert into tradingac (slno,CR_PARTICULARS,CRAMOUNT) VALUES(" & tempslno & ",'" & rec2("AccName") & "'," & temp_cr - temp_dr & ")")
            End If
            tempslno = tempslno + 1
            rec2.MoveNext
        Wend
        rec1.MoveNext
    Wend
    temp_total_dr = 0
    temp_total_cr = 0
    Set rec1 = db.OpenRecordset("select sum(dramount) as total_dr from tradingac")
    If Not IsNull(rec1!Total_Dr) Then
        temp_total_dr = rec1!Total_Dr
    End If
    Set rec1 = db.OpenRecordset("select sum(cramount) as total_cr from tradingac")
    If Not IsNull(rec1!Total_Cr) Then
        temp_total_cr = rec1!Total_Cr
    End If
    '---------gross profit---------------------
    If temp_total_cr > temp_total_dr Then
        net_diff = temp_total_cr - temp_total_dr
        db.Execute ("insert into tradingac (slno,DR_PARTICULARS,DRAMOUNT) VALUES(" & tempslno + 1 & ",'GROSS PROFIT c/o'," & net_diff & ")")
    End If
    '---------gross loss-------------------
    If temp_total_dr > temp_total_cr Then
        net_diff = temp_total_dr - temp_total_cr
        db.Execute ("insert into tradingac (slno,CR_PARTICULARS,CRAMOUNT) VALUES(" & tempslno + 1 & ",'GROSS LOSS c/o'," & net_diff & ")")
    End If
    db.Execute ("UPDATE TRADINGAC SET todt='" & Me.txtto.Text & "'")
    Data1.databasename = dbname
    Data1.Refresh
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtFrom_GotFocus()
    Me.txtfrom.SelStart = 0
    Me.txtfrom.SelLength = Len(Me.txtfrom.Text)
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtto.SetFocus
    End If
End Sub

Private Sub txtTo_GotFocus()
    Me.txtto.SelStart = 0
    Me.txtto.SelLength = Len(Me.txtto.Text)
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        db.Execute ("update tradingac set fromdt='" & Me.txtfrom.Text & "',todt='" & Me.txtto.Text & "'")
    End If
End Sub
