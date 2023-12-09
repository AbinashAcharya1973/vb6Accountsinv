VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmTrailBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trail Balance"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9090
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      Begin VB.TextBox TxtDatefrom 
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
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtdateto 
         Height          =   375
         Left            =   3000
         TabIndex        =   0
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         Caption         =   "From"
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
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "To"
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
         Left            =   2640
         TabIndex        =   6
         Top             =   120
         Width           =   375
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print/View"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   7920
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmTrailBalance.frx":0000
      Height          =   6855
      Left            =   120
      OleObjectBlob   =   "frmTrailBalance.frx":0014
      TabIndex        =   1
      Top             =   720
      Width           =   8895
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
      RecordSource    =   "TrailBalance"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   7080
      TabIndex        =   3
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   7560
      Width           =   1695
   End
End
Attribute VB_Name = "frmTrailBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec As DAO.Recordset
Attribute rec.VB_VarUserMemId = 1073938432
Private Sub cmdprint_Click()
    CrystalReport1.PrintReport
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    Dim Total_Dr As Currency, temp_cr As Currency, temp_dr As Currency, Total_Cr As Currency
    Me.CrystalReport1.ReportFileName = App.Path & "\TRAILBAL.RPT"
    Me.Top = 0
    Me.Left = 0
    Me.txtdatefrom.Text = AccountingPeriod
    Set rs = db.OpenRecordset("select * from groups where GP_On_TB='y'")
    While Not rs.EOF

        Set rec = db.OpenRecordset("select * from ledgermaster where groupid=" & rs("groupid"))
        While Not rec.EOF
            X = X + 1
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


            '        If rs("GROUPID") = 17 Then
            '        Debug.Print rec("ACCNAME") & "," & temp_dr & "," & temp_cr
            '        End If
            rec.MoveNext
        Wend
        '    If Total_Dr > Total_Cr Then
        '        Db.Execute ("insert into trailbalance (AccName,dr) values('" & rs("groupname") & "'," & Total_Dr - Total_Cr & ")")
        '    End If
        '    If Total_Cr > Total_Dr Then
        '        Db.Execute ("insert into trailbalance (AccName,cr) values('" & rs("groupname") & "'," & Total_Cr - Total_Dr & ")")
        '    End If
        db.Execute ("insert into trailbalance (AccName,dr,Cr) values('" & rs("groupname") & "'," & Total_Dr & "," & Total_Cr & ")")
        If Total_Dr = Total_Cr Then

        End If
        Total_Dr = 0
        Total_Cr = 0
        temp_dr = 0
        temp_cr = 0
        X = 0
        rs.MoveNext

    Wend

    '---------------------without grouping of ledger--------------------------------------------------------
    Set rs = db.OpenRecordset("select * from groups where GP_On_TB='n'")
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
                db.Execute ("insert into trailbalance (AccName,dr) values('" & rec("accname") & "'," & temp_dr - temp_cr & ")")
            End If
            If temp_cr > temp_dr Then
                db.Execute ("insert into trailbalance (AccName,cr) values('" & rec("accname") & "'," & temp_cr - temp_dr & ")")
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
    '-----------USING OPENING BALANCE ----------------------------------------------
    Set rs = db.OpenRecordset("SELECT SUM(OBALANCE) AS SUM_DR_BALANCE FROM LEDGERMASTER WHERE BALANCETYPE='Dr'")
    If Not IsNull(rs!SUM_DR_BALANCE) Then
        TEMP_OP_DR_BALANCE = rs!SUM_DR_BALANCE
    Else
        TEMP_OP_DR_BALANCE = 0
    End If

    Set rs = db.OpenRecordset("SELECT SUM(OBALANCE) AS SUM_CR_BALANCE FROM LEDGERMASTER WHERE BALANCETYPE='Cr'")
    If Not IsNull(rs!SUM_CR_BALANCE) Then
        TEMP_OP_CR_BALANCE = rs!SUM_CR_BALANCE
    Else
        TEMP_OP_CR_BALANCE = 0
    End If

    If TEMP_OP_DR_BALANCE <> TEMP_OP_CR_BALANCE Then
        opening_balance_diff = Abs(TEMP_OP_DR_BALANCE - TEMP_OP_CR_BALANCE)

        Set rs = db.OpenRecordset("select sum(dr) as trail_dr from trailbalance")
        If Not IsNull(rs!trail_dr) Then
            TRAIL_DR_OP = rs!trail_dr
        End If
        Set rs = db.OpenRecordset("select sum(cr) as trail_cr from trailbalance")
        If Not IsNull(rs!trail_Cr) Then
            TRAIL_CR_OP = rs!trail_Cr
        End If
        If TRAIL_DR_OP > TRAIL_CR_OP Then
            db.Execute ("insert into trailbalance (AccName,Cr) values('Difference In Opening Balances'," & Abs(TRAIL_DR_OP - TRAIL_CR_OP) & ")")
        End If
        If TRAIL_CR_OP > TRAIL_DR_OP Then
            db.Execute ("insert into trailbalance (AccName,Dr) values('Difference In Opening Balances'," & Abs(TRAIL_DR_OP - TRAIL_CR_OP) & ")")
        End If
    End If

    'If TEMP_OP_DR_BALANCE > TEMP_OP_CR_BALANCE Then
    '    db.Execute ("insert into trailbalance (AccName,dr) values('Difference In Opening Balances'," & TEMP_OP_DR_BALANCE - TEMP_OP_CR_BALANCE & ")")
    'End If
    'If TEMP_OP_CR_BALANCE > TEMP_OP_DR_BALANCE Then
    '    db.Execute ("insert into trailbalance (AccName,cr) values('Difference In Opening Balances'," & TEMP_OP_CR_BALANCE - TEMP_OP_DR_BALANCE & ")")
    'End If

    '----------------------------------------------------------------------
    Set rs = db.OpenRecordset("select sum(dr) as trail_dr from trailbalance")
    If Not IsNull(rs!trail_dr) Then
        Me.Label1.Caption = Format(rs!trail_dr, "########0.00")
    End If
    Set rs = db.OpenRecordset("select sum(cr) as trail_cr from trailbalance")
    If Not IsNull(rs!trail_Cr) Then
        Me.Label2.Caption = Format(rs!trail_Cr, "########0.00")
    End If
    Data1.databasename = dbname
    Data1.RecordSource = ("select * from trailbalance where Dr>0 or Cr>0")
    Data1.Refresh
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("delete * from trailbalance")
End Sub

Public Sub Trail1()
    Set rs = db.OpenRecordset("select * from ledgermaster")
    While Not rs.EOF
        Set rec1 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rs("accid"))
        If Not IsNull(rec1!Total_Dr) Then
            Total_Dr = rec1!Total_Dr
        Else
            Total_Dr = 0
        End If
        Set rec1 = db.OpenRecordset("select sum(cr) as total_cr from ledgertran where accid=" & rs("accid"))
        If Not IsNull(rec1!Total_Cr) Then
            Total_Cr = rec1!Total_Cr
        Else
            Total_Cr = 0
        End If

        If Total_Dr > Total_Cr Then
            db.Execute ("insert into trailbalance (AccName,dr) values('" & rs("accname") & "'," & Total_Dr - Total_Cr & ")")
        End If
        If Total_Cr > Total_Dr Then
            db.Execute ("insert into trailbalance (AccName,cr) values('" & rs("accname") & "'," & Total_Cr - Total_Dr & ")")
        End If
        If Total_Dr = Total_Cr Then

        End If
        rs.MoveNext
    Wend

End Sub

Private Sub txtdateto_GotFocus()
    Me.txtdateto.SelStart = 0
    Me.txtdateto.SelLength = Len(Me.txtdateto.Text)
End Sub
Private Sub txtdateto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Total_Dr As Currency, temp_cr As Currency, temp_dr As Currency, Total_Cr As Currency
    If KeyCode = 13 Then
        db.Execute ("delete * from trailbalance")
        db.Execute ("INSERT INTO  TrailBalance (TDATE) VALUES('" & Me.txtdateto.Text & "')")
        temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
        temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
        temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)

        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year

        temp_to_date = Mid(Me.txtdateto.Text, 1, 2)
        temp_to_month = Mid(Me.txtdateto.Text, 4, 2)
        temp_to_year = Mid(Me.txtdateto.Text, 7, 4)

        temp_to = temp_to_month & "/" & temp_to_date & "/" & temp_to_year
        '-------------------------------------------------------------------
        Set rs = db.OpenRecordset("select * from groups where GP_On_TB='y'")
        While Not rs.EOF

            Set rec = db.OpenRecordset("select * from ledgermaster where groupid=" & rs("groupid"))
            While Not rec.EOF
                X = X + 1
                Set rec1 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rec("accid") & " and tdate between #" & temp_from & "# and #" & temp_to & "#")
                If Not IsNull(rec1!Total_Dr) Then
                    temp_dr = rec1!Total_Dr
                Else
                    temp_dr = 0
                End If
                Set rec1 = db.OpenRecordset("select sum(cr) as total_cr from ledgertran where accid=" & rec("accid") & " and tdate between #" & temp_from & "# and #" & temp_to & "#")
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
            db.Execute ("insert into trailbalance (AccName,dr,Cr) values('" & rs("groupname") & "'," & Total_Dr & "," & Total_Cr & ")")
            If Total_Dr = Total_Cr Then

            End If
            Total_Dr = 0
            Total_Cr = 0
            temp_dr = 0
            temp_cr = 0
            X = 0
            rs.MoveNext

        Wend

        '---------------------without grouping of ledger--------------------------------------------------------
        Set rs = db.OpenRecordset("select * from groups where GP_On_TB='n'")
        While Not rs.EOF
            Set rec = db.OpenRecordset("select * from ledgermaster where groupid=" & rs("groupid"))
            While Not rec.EOF
                Set rec1 = db.OpenRecordset("select sum(dr) as total_dr from ledgertran where accid=" & rec("accid") & " and tdate between #" & temp_from & "# and #" & temp_to & "#")
                If Not IsNull(rec1!Total_Dr) Then
                    temp_dr = rec1!Total_Dr
                Else
                    temp_dr = 0
                End If
                Set rec1 = db.OpenRecordset("select sum(Cr) as total_cr from ledgertran where accid=" & rec("accid") & " and tdate between #" & temp_from & "# and #" & temp_to & "#")
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
                    db.Execute ("insert into trailbalance (AccName,dr) values('" & rec("accname") & "'," & temp_dr - temp_cr & ")")
                End If
                If temp_cr > temp_dr Then
                    db.Execute ("insert into trailbalance (AccName,cr) values('" & rec("accname") & "'," & temp_cr - temp_dr & ")")
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
        '-----------USING OPENING BALANCE ----------------------------------------------
        Set rs = db.OpenRecordset("SELECT SUM(OBALANCE) AS SUM_DR_BALANCE FROM LEDGERMASTER WHERE BALANCETYPE='Dr'")
        If Not IsNull(rs!SUM_DR_BALANCE) Then
            TEMP_OP_DR_BALANCE = rs!SUM_DR_BALANCE
        Else
            TEMP_OP_DR_BALANCE = 0
        End If

        Set rs = db.OpenRecordset("SELECT SUM(OBALANCE) AS SUM_CR_BALANCE FROM LEDGERMASTER WHERE BALANCETYPE='Cr'")
        If Not IsNull(rs!SUM_CR_BALANCE) Then
            TEMP_OP_CR_BALANCE = rs!SUM_CR_BALANCE
        Else
            TEMP_OP_CR_BALANCE = 0
        End If

        If TEMP_OP_DR_BALANCE <> TEMP_OP_CR_BALANCE Then
            opening_balance_diff = Abs(TEMP_OP_DR_BALANCE - TEMP_OP_CR_BALANCE)

            Set rs = db.OpenRecordset("select sum(dr) as trail_dr from trailbalance")
            If Not IsNull(rs!trail_dr) Then
                TRAIL_DR_OP = rs!trail_dr
            End If
            Set rs = db.OpenRecordset("select sum(cr) as trail_cr from trailbalance")
            If Not IsNull(rs!trail_Cr) Then
                TRAIL_CR_OP = rs!trail_Cr
            End If
            If TRAIL_DR_OP > TRAIL_CR_OP Then
                db.Execute ("insert into trailbalance (AccName,Cr) values('Difference In Opening Balances'," & Abs(TRAIL_DR_OP - TRAIL_CR_OP) & ")")
            End If
            If TRAIL_CR_OP > TRAIL_DR_OP Then
                db.Execute ("insert into trailbalance (AccName,Dr) values('Difference In Opening Balances'," & Abs(TRAIL_DR_OP - TRAIL_CR_OP) & ")")
            End If
        End If

        '----------------------------------------------------------------------
        Set rs = db.OpenRecordset("select sum(dr) as trail_dr from trailbalance")
        If Not IsNull(rs!trail_dr) Then
            Me.Label1.Caption = Format(rs!trail_dr, "########0.00")
        End If
        Set rs = db.OpenRecordset("select sum(cr) as trail_cr from trailbalance")
        If Not IsNull(rs!trail_Cr) Then
            Me.Label2.Caption = Format(rs!trail_Cr, "########0.00")
        End If
        Data1.databasename = dbname
        Data1.RecordSource = ("select * from trailbalance where Dr>0 or Cr>0")
        Data1.Refresh
    End If
End Sub
