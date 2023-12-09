VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "Crystl32.OCX"
Begin VB.Form frmPLAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P & L Account"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1560
      Top             =   7080
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
      Left            =   4200
      TabIndex        =   3
      Top             =   7080
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPLAcc.frx":0000
      Height          =   6135
      Left            =   120
      OleObjectBlob   =   "frmPLAcc.frx":0014
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "P_LAcc"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSMask.MaskEdBox txtTo 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
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
      TabIndex        =   7
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
      TabIndex        =   6
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
Attribute VB_Name = "frmPLAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As DAO.Recordset, rec3 As DAO.Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Private Sub CMDPRINT_Click()
    db.Execute ("update P_LACC set todt='" & Me.txtTo.Text & "'")
    CrystalReport1.PrintReport
End Sub
Private Sub Form_Load()
    Me.CrystalReport1.ReportFileName = App.Path & "\pl_acc.rpt"
    Dim opstock, clstock, purchase, sale, preturn, salereturn
    db.Execute ("DELETE * FROM P_LAcc")
    opstock = 0
    clstock = 0
    purchase = 0
    preturn = 0
    sale = 0
    salereturn = 0
    tempslno = 1
    cr_no = 1
    Set rec1 = db.OpenRecordset("select * from TRADINGAC where slno=(select max(slno) from TRADINGAC)")

    If Not rec1.EOF Then
        If rec1("DR_PARTICULARS") = "GROSS PROFIT c/o" Then
            db.Execute ("insert into P_LACC (slno,Cr_Particulars,CRAMOUNT) VALUES(" & tempslno & ",'GROSS PROFIT b/f'," & rec1("DRAMOUNT") & ")")
            cr_no = tempslno + 1
        End If
        If rec1("CR_PARTICULARS") = "GROSS LOSS c/o" Then
            db.Execute ("insert into P_LACC (slno,DR_PARTICULARS,DRAMOUNT) VALUES(" & tempslno & ",'GROSS LOSS b/f'," & rec1("CRAMOUNT") & ")")
        End If
    End If
    tempslno = tempslno + 1

    '-----OTHER EXPENCES-----------------------
    Set rec1 = db.OpenRecordset("select * from groups where GroupNature='Expences' and Affect_GP='N'")
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
                db.Execute ("insert into P_LAcc (slno,DR_PARTICULARS,DRAMOUNT) VALUES(" & tempslno & ",'" & rec2("AccName") & "'," & temp_dr - temp_cr & ")")
                tempslno = tempslno + 1
            End If
            If temp_cr > temp_dr Then
                db.Execute ("insert into P_LAcc (slno,DR_PARTICULARS,DRAMOUNT) VALUES(" & tempslno & ",'" & rec2("AccName") & "'," & temp_cr - temp_dr & ")")
                tempslno = tempslno + 1
            End If

            rec2.MoveNext
        Wend
        rec1.MoveNext
    Wend
    max_slno = tempslno
    tempslno = 1

    '-------INSERTING INCOMES----------------------
    Set rec1 = db.OpenRecordset("select * from groups where GroupNature='Income' and Affect_GP='N'")
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
            'If tempslno <= max_slno Then
            If cr_no <= max_slno Then
                If temp_dr > temp_cr Then
                    'Db.Execute ("insert into P_LAcc (CR_PARTICULARS,CRAMOUNT) values('" & rec2("AccName") & "'," & temp_dr - temp_cr & ")")
                    db.Execute ("update P_LAcc set CR_PARTICULARS='" & rec2("AccName") & "',CRAMOUNT=" & temp_dr - temp_cr & " where slno=" & cr_no)
                    'tempslno = tempslno + 1
                    cr_no = cr_no + 1
                End If
                If temp_cr > temp_dr Then
                    db.Execute ("update P_LAcc set CR_PARTICULARS='" & rec2("AccName") & "',CRAMOUNT=" & temp_cr - temp_dr & " where slno=" & cr_no)
                    tempslno = tempslno + 1
                    cr_no = cr_no + 1
                End If
            Else
                If temp_dr > temp_cr Then
                    db.Execute ("insert into P_LAcc (CR_PARTICULARS,CRAMOUNT) values('" & rec2("AccName") & "'," & temp_dr - temp_cr & ")")
                    tempslno = tempslno + 1
                End If
                If temp_cr > temp_dr Then
                    db.Execute ("insert into P_LAcc (CR_PARTICULARS,CRAMOUNT) values('" & rec2("AccName") & "'," & temp_cr - temp_dr & ")")
                    tempslno = tempslno + 1
                End If
            End If
            rec2.MoveNext
        Wend
        rec1.MoveNext
    Wend

    temp_total_dr = 0
    temp_total_cr = 0
    Set rec1 = db.OpenRecordset("select sum(dramount) as total_dr from p_lacc")
    If Not IsNull(rec1!Total_Dr) Then
        temp_total_dr = rec1!Total_Dr
    End If
    Set rec1 = db.OpenRecordset("select sum(cramount) as total_cr from p_lacc")
    If Not IsNull(rec1!Total_Cr) Then
        temp_total_cr = rec1!Total_Cr
    End If
    '---------NET profit---------------------
    If temp_total_cr > temp_total_dr Then
        net_diff = temp_total_cr - temp_total_dr
        db.Execute ("insert into P_LACC (slno,DR_PARTICULARS,DRAMOUNT) VALUES(" & tempslno + 1 & ",'NET PROFIT'," & net_diff & ")")
    End If
    '---------NET loss-------------------
    If temp_total_dr > temp_total_cr Then
        net_diff = temp_total_dr - temp_total_cr
        db.Execute ("insert into P_LACC (slno,CR_PARTICULARS,CRAMOUNT) VALUES(" & tempslno + 1 & ",'NET LOSS'," & net_diff & ")")
    End If
    Data1.databasename = dbname
    Data1.Refresh
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtTo.SetFocus
    End If
End Sub

Private Sub txtTo_GotFocus()
    Me.txtTo.SelStart = 0
    Me.txtTo.SelLength = Len(Me.txtTo.Text)
End Sub
