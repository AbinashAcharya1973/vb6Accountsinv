VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "Crystl32.OCX"
Begin VB.Form frmCashBookView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Book View"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10935
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   10320
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdjournalprint 
      Caption         =   "Journal Print"
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
      Left            =   8760
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3720
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Print"
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
      Left            =   7080
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   5520
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox txtview 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmCashBookView.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdgenerate 
      Caption         =   "Generate Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin MSMask.MaskEdBox txtdateto 
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   360
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
   Begin MSMask.MaskEdBox txtdatefrom 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   1200
      TabIndex        =   3
      Top             =   360
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
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmCashBookView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As DAO.Recordset, rec1 As DAO.Recordset, rs As DAO.Recordset
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rs.VB_VarUserMemId = 1073938432
Private Sub cmdgenerate_Click()
    Dim pageno, lineno
    pageno = 1
    lineno = 1
    Kill ("d:\testfile.txt")
    Set fso = CreateObject("Scripting.FileSystemObject")


    If Me.txtdatefrom.Text <> "__/__/____" And Me.txtdateto.Text <> "__/__/____" Then
        Set fh = fso.CreateTextFile("d:\testfile.txt", True)

        temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
        temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
        temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)

        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year

        temp_to_date = Mid(Me.txtdateto.Text, 1, 2)
        temp_to_month = Mid(Me.txtdateto.Text, 4, 2)
        temp_to_year = Mid(Me.txtdateto.Text, 7, 4)

        temp_to = temp_to_month & "/" & temp_to_date & "/" & temp_to_year
        '-----------------------PAGE HEADER-------------------------------------
        fh.WriteLine ("Page No-" & pageno)
        fh.WriteLine (Chr(27) + Chr(14) + Chr(14) & "SHANKAR TEXTILES" & Chr(20) + Chr(10))
        fh.WriteLine ("JAUNLIAPATTY" & Space(35) & "Financial Year : 2006-2007")
        fh.WriteLine ("CUTTACK" & Space(40) & "Assessment Year : 2007-2008")
        fh.WriteLine (Chr(27) + Chr(69) & "CASH BOOK (SUMMERY)" & Chr(27) + Chr(70))
        fh.WriteLine (String(80, "-"))
        fh.WriteLine (Chr(27) + Chr(71) & "Voucher-Typ." & Space(2) & "Details" & Space(32 - Len("Details")) & Space(2) & Space(15 - Len("Dr")) & "Dr" & Space(2) & Space(15 - Len("Dr")) & "Cr" & Chr(27) + Chr(72))
        fh.WriteLine (String(80, "-"))
        '-------------------------------------------------------------------------------
        lineno = 10

        Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where accid=7 AND TDATE<#" & temp_from & "#")
        If Not IsNull(rs!max_dr) Then
            temp_dr = rs!max_dr
        Else
            temp_dr = 0
        End If
        Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where accid=7 AND TDATE<#" & temp_from & "#")
        If Not IsNull(rs!max_cr) Then
            temp_cr = rs!max_cr
        Else
            temp_cr = 0
        End If

        Set rs = db.OpenRecordset("select * from ledgermaster where accid=7")
        If Not rs.EOF Then
            If rs("BalanceType") = "Dr" Then
                temp_dr = temp_dr + rs("OBalance")
            End If
            If rs("BalanceType") = "Cr" Then
                temp_cr = temp_cr + rs("OBalance")
            End If

        End If

        Set rec = db.OpenRecordset("select * from ledgertran where accid=7 and tdate between #" & temp_from & "# and #" & temp_to & "# order by tdate")
        '-----OPENING BALANCE--------------------------------------------

        If temp_cr > temp_dr Then
            fh.WriteLine (Chr(27) + Chr(71) & Space(16) & "BY BALANCE B/D" & Space(47) & Space(15 - Len(Format(temp_cr - temp_dr, "########0.00")))) & Format(temp_cr - temp_dr, "########0.00") & Chr(27) + Chr(72)
            temp_cr = temp_cr - temp_dr
            temp_dr = 0
            lineno = lineno + 1
        End If

        If temp_dr > temp_cr Then
            fh.WriteLine (Chr(27) + Chr(71) & Space(16) & "TO BALANCE B/D" & Space(32 - Len("TO BALANCE B/D")) & Space(15 - Len(Format(temp_dr - temp_cr, "########0.00")))) & Format(temp_dr - temp_cr, "########0.00") & Chr(27) + Chr(72)
            temp_dr = temp_dr - temp_cr
            temp_cr = 0
            lineno = lineno + 1
        End If



        If Not rec.EOF Then
            c_date = rec("tdate")
            fh.WriteLine (c_date)
            fh.WriteLine ("----------")
            lineno = lineno + 2
        End If

        DAY_TOTAL_dr = temp_dr
        DAY_TOTAL_Cr = temp_cr
        While Not rec.EOF
            Set rec1 = db.OpenRecordset("select * from ledgermaster where accid=" & rec("tranaccid"))
            If c_date = rec("tdate") Then
                If lineno > 40 Then
                    fh.WriteLine (String(80, "-"))
                    fh.WriteLine ("C/D" & Space(45) & Space(15 - Len(Format(DAY_TOTAL_dr, "########0.00"))) & (Format(DAY_TOTAL_dr, "########0.00")) & Space(2) & Space(15 - Len(Format(DAY_TOTAL_Cr, "########0.00"))) & (Format(DAY_TOTAL_Cr, "########0.00")))
                    fh.WriteLine (Chr(12))    'ESCAPE CODE FOR NEXT PAGE
                    '-Page Header------------------------
                    fh.WriteLine ("Page No-" & pageno)
                    fh.WriteLine (Chr(27) + Chr(14) + Chr(14) & "SHANKAR TEXTILES" & Chr(20) + Chr(10))
                    fh.WriteLine ("JAUNLIAPATTY" & Space(35) & "Financial Year : 2006-2007")
                    fh.WriteLine ("CUTTACK" & Space(40) & "Assessment Year : 2007-2008")
                    fh.WriteLine (Chr(27) + Chr(69) & "CASH BOOK (SUMMERY)" & Chr(27) + Chr(70))
                    fh.WriteLine (String(80, "-"))
                    fh.WriteLine (Chr(27) + Chr(71) & "Voucher-Typ." & Space(2) & "Details" & Space(32 - Len("Details")) & Space(2) & Space(15 - Len("Dr")) & "Dr" & Space(2) & Space(15 - Len("Dr")) & "Cr" & Chr(27) + Chr(72))
                    fh.WriteLine (String(80, "-"))
                    fh.WriteLine ("B/D" & Space(45) & Space(15 - Len(Format(DAY_TOTAL_dr, "########0.00"))) & (Format(DAY_TOTAL_dr, "########0.00")) & Space(2) & Space(15 - Len(Format(DAY_TOTAL_Cr, "########0.00"))) & (Format(DAY_TOTAL_Cr, "########0.00")))
                    c_date = rec("tdate")
                    fh.WriteLine (c_date)
                    fh.WriteLine ("----------")
                    lineno = lineno + 10
                End If
                fh.WriteLine (Trim(rec("vouchertype")) & Space(14 - Len(rec("vouchertype"))) & Space(2) & rec1("Accname") & Space(50 - Len(rec1("Accname"))) & Space(2) & Space(15 - Len(Format(rec("dr"), "###########.00"))) & Format(rec("dr"), "##############.00") & Space(2) & Space(15 - Len(Format(rec("cr"), "###########.00"))) & Format(rec("cr"), "##############.00"))
                lineno = lineno + 1
                If Not IsNull(rec1("address1")) Then
                    fh.WriteLine (Space(16) & rec1("address1"))
                    lineno = lineno + 1
                End If
                If Not IsNull(rec("remarks")) Then
                    fh.WriteLine (Space(16) & rec("remarks"))
                End If
                lineno = lineno + 1
                DAY_TOTAL_dr = DAY_TOTAL_dr + rec("dr")
                DAY_TOTAL_Cr = DAY_TOTAL_Cr + rec("cr")
            Else
                TEMP_DR_BALANCE = 0
                TEMP_CR_BALANCE = 0
                fh.WriteLine (Space(48) & String(15, "-") & Space(2) & String(15, "-"))
                fh.WriteLine (Space(48) & Space(15 - Len(Format(DAY_TOTAL_dr, "########0.00"))) & Format(DAY_TOTAL_dr, "########0.00") & Space(2) & Space(15 - Len(Format(DAY_TOTAL_Cr, "########0.00"))) & Format(DAY_TOTAL_Cr, "########0.00"))
                lineno = lineno + 2
                If DAY_TOTAL_Cr > DAY_TOTAL_dr Then
                    fh.WriteLine (Chr(27) + Chr(71) & Space(16) & "BY BALANCE C/D" & Space(32 - Len("TO BALANCE B/D")) & Space(17) & Space(15 - Len(Format(DAY_TOTAL_Cr - DAY_TOTAL_dr, "########0.00")))) & Format(DAY_TOTAL_Cr - DAY_TOTAL_dr, "########0.00") & Chr(27) + Chr(72)
                    lineno = lineno + 1
                    temp_cr = DAY_TOTAL_Cr - DAY_TOTAL_dr
                    TEMP_DR_BALANCE = DAY_TOTAL_Cr - DAY_TOTAL_dr
                    temp_dr = 0
                End If

                If temp_dr > temp_cr Then
                    fh.WriteLine (Chr(27) + Chr(71) & Space(16) & "TO BALANCE C/D" & Space(32 - Len("TO BALANCE B/D")) & Space(17) & Space(15 - Len(Format(DAY_TOTAL_dr - DAY_TOTAL_Cr, "########0.00")))) & Format(DAY_TOTAL_dr - DAY_TOTAL_Cr, "########0.00") & Chr(27) + Chr(72)
                    lineno = lineno + 1
                    temp_dr = DAY_TOTAL_dr - DAY_TOTAL_Cr
                    TEMP_CR_BALANCE = DAY_TOTAL_dr - DAY_TOTAL_Cr
                    temp_cr = 0
                End If
                fh.WriteLine (Space(48) & Space(15 - Len(Format(DAY_TOTAL_dr + TEMP_DR_BALANCE, "########0.00"))) & Format(DAY_TOTAL_dr + TEMP_DR_BALANCE, "########0.00") & Space(2) & Space(15 - Len(Format(DAY_TOTAL_Cr + TEMP_CR_BALANCE, "########0.00"))) & Format(DAY_TOTAL_Cr + TEMP_CR_BALANCE, "########0.00"))
                lineno = lineno + 1
                '-----------------------------GO TO NEXT PAGE-------------------------------------------

                fh.WriteLine (" ")
                DAY_TOTAL_dr = temp_dr
                DAY_TOTAL_Cr = temp_cr
                fh.WriteLine (Chr(12))    'ESCAPE CODE FOR NEXT PAGE
                lineno = 1
                pageno = pageno + 1
                '-----------------------PAGE HEADER-------------------------------------
                fh.WriteLine ("Page No-" & pageno)
                fh.WriteLine (Chr(27) + Chr(14) + Chr(14) & "SHANKAR TEXTILES" & Chr(20) + Chr(10))
                fh.WriteLine ("JAUNLIAPATTY" & Space(35) & "Financial Year : 2006-2007")
                fh.WriteLine ("CUTTACK" & Space(40) & "Assessment Year : 2007-2008")
                fh.WriteLine (Chr(27) + Chr(69) & "CASH BOOK (SUMMERY)" & Chr(27) + Chr(70))
                fh.WriteLine (String(80, "-"))
                fh.WriteLine (Chr(27) + Chr(71) & "Voucher-Typ." & Space(2) & "Details" & Space(32 - Len("Details")) & Space(2) & Space(15 - Len("Dr")) & "Dr" & Space(2) & Space(15 - Len("Dr")) & "Cr" & Chr(27) + Chr(72))
                fh.WriteLine (String(80, "-"))
                lineno = lineno + 8
                '-------------------------------------------------------------------------------
                If temp_cr > temp_dr Then
                    fh.WriteLine (Chr(27) + Chr(71) & Space(16) & "BY BALANCE B/D" & Space(47) & Space(15 - Len(Format(temp_cr - temp_dr, "########0.00")))) & Format(temp_cr - temp_dr, "########0.00") & Chr(27) + Chr(72)
                End If
                lineno = lineno + 1
                If temp_dr > temp_cr Then
                    fh.WriteLine (Chr(27) + Chr(71) & Space(16) & "TO BALANCE B/D" & Space(32 - Len("TO BALANCE B/D")) & Space(15 - Len(Format(temp_dr - temp_cr, "########0.00")))) & Format(temp_dr - temp_cr, "########0.00") & Chr(27) + Chr(72)
                End If
                lineno = lineno + 1
                c_date = rec("tdate")
                fh.WriteLine (c_date)
                fh.WriteLine ("----------")
                fh.WriteLine (Trim(rec("vouchertype")) & Space(14 - Len(rec("vouchertype"))) & Space(2) & rec1("Accname") & Space(50 - Len(rec1("Accname"))) & Space(2) & Space(15 - Len(Format(rec("dr"), "###########.00"))) & Format(rec("dr"), "##############.00") & Space(2) & Space(15 - Len(Format(rec("cr"), "###########.00"))) & Format(rec("cr"), "##############.00"))
                lineno = lineno + 3
                If Not IsNull(rec1("address1")) Then
                    fh.WriteLine (Space(16) & rec1("address1"))
                End If
                lineno = lineno + 1
                If Not IsNull(rec("remarks")) Then
                    fh.WriteLine (Space(16) & rec("remarks"))
                End If
                lineno = lineno + 1
                DAY_TOTAL_dr = DAY_TOTAL_dr + rec("dr")
                DAY_TOTAL_Cr = DAY_TOTAL_Cr + rec("cr")
            End If
            rec.MoveNext
        Wend
        fh.WriteLine (Chr(12))
        fh.Close
    End If
    Me.txtview.FileName = "d:\testfile.txt"
End Sub

Private Sub cmdjournalprint_Click()
    db.Execute ("Delete * from JournalPrint")
    Set rec1 = db.OpenRecordset("select * from ledgertran where VoucherType='Receipt' or VoucherType='Payment' or VoucherType='Journal' or VoucherType='Contra' order by tdate")
    While Not rec1.EOF
        Set rec2 = db.OpenRecordset("select * from ledgermaster where accid=" & rec1("TranAccid"))
        If Not rec2.EOF Then
            db.Execute ("insert into JournalPrint (Tdate,VoucherType,Accname,Narration,Dr,Cr,Accid,Groupname,Tranaccid,VoucherNo,Address) values('" & rec1("Tdate") & "','" & rec1("VoucherType") & "','" & rec1("Particulars") & "','" & rec1("remarks") & "'," & rec1("Dr") & "," & rec1("Cr") & "," & rec1("Accid") & ",'" & rec2("Groupname") & "'," & rec1("Tranaccid") & "," & rec1("VoucherSlNo") & ",'" & rec2("Address1") & "')")
        End If
        rec1.MoveNext
    Wend
    db.Close

    Me.CrystalReport2.ReportFileName = App.Path & "\dailybook.rpt"
    Me.CrystalReport2.SelectionFormula = "{Journalprint.Accid} <> 10 and {Journalprint.Tranaccid} <> 10"
    Me.CrystalReport2.PrintReport
End Sub

Private Sub CMDPRINT_Click()
'txtview.SelPrint Printer.hDC
    PrintTXTFile "d:\testfile.txt"
End Sub

Public Sub PrintTXTFile(FileName As String)
    Dim X As Integer
    Dim s As String

    X = FreeFile
    On Error GoTo HandleError
    Open FileName For Input As X
    filenum = FreeFile
    Open "lpt1:" For Output As filenum

    Do While Not EOF(X)
        Line Input #X, s
        Print #filenum, s
    Loop
    Close #X
    Close #filenum
    Exit Sub
HandleError:
    MsgBox "Error :" & Err.Description, vbCritical, "Printing File..."
End Sub
Private Sub Command1_Click()
    db.Execute ("delete * from Cashbookprint")
    tdate = #4/1/2008#
    lastdate = #3/31/2009#
    adate = Format(tdate, "mm")
    bdate = Format(tdate, "dd")
    rdate = Format(tdate, "yyyy")
    ldate = adate & "/" & bdate & "/" & rdate
    Drop = 0
    CrOp = 0
    While tdate <= lastdate
        temp_dr = 0
        temp_cr = 0
        Opdr = 0
        OpCr = 0
        Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where accid=10 AND TDATE<#" & ldate & "#")
        If Not IsNull(rs!max_dr) Then
            temp_dr = rs!max_dr
        Else
            temp_dr = 0
        End If
        Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where accid=10 AND TDATE<#" & ldate & "#")
        If Not IsNull(rs!max_cr) Then
            temp_cr = rs!max_cr
        Else
            temp_cr = 0
        End If
        Set rs = db.OpenRecordset("select * from ledgermaster where accid=10")
        If Not rs.EOF Then
            If rs("BalanceType") = "Dr" Then
                temp_dr = temp_dr + rs("OBalance")
            End If
            If rs("BalanceType") = "Cr" Then
                temp_cr = temp_cr + rs("OBalance")
            End If

        End If
        If temp_cr > temp_dr Then
            CrOp = temp_cr - temp_dr
            Drop = 0
        End If
        If temp_dr > temp_cr Then
            Drop = temp_dr - temp_cr
            CrOp = 0
        End If
        DAY_TOTAL_dr = 0
        DAY_TOTAL_Cr = 0
        Set rec1 = db.OpenRecordset("select * from ledgertran where accid=10 and tdate=#" & ldate & "#")
        While Not rec1.EOF
            db.Execute ("insert into Cashbookprint (Tdate,Voucher,Account,Narration,Dr,Cr,BalanceDr,BalanceCr,Opdr,Opcr,Accid) values('" & tdate & "','" & rec1("VoucherType") & "','" & rec1("Particulars") & "','" & rec1("Remarks") & "'," & rec1("Dr") & "," & rec1("Cr") & ",0,0," & Drop & "," & CrOp & "," & rec1("tranaccid") & ")")
            DAY_TOTAL_dr = DAY_TOTAL_dr + rec1("Dr")
            DAY_TOTAL_Cr = DAY_TOTAL_Cr + rec1("Cr")
            rec1.MoveNext
        Wend
        'If (DAY_TOTAL_dr + Drop) > (DAY_TOTAL_Cr + CrOp) Then
        ' Drop = (DAY_TOTAL_dr + Drop) - (DAY_TOTAL_Cr + CrOp)
        ' CrOp = DAY_TOTAL_Cr
        'Db.Execute ("update Cashbookprint set BalanceDr=" & Drop & ",BalanceCr=" & CrOp & " where tdate=#" & ldate & "#")
        'Else
        'CrOp = (DAY_TOTAL_Cr + CrOp) - (DAY_TOTAL_dr + Drop)
        'Drop = DAY_TOTAL_dr
        'Db.Execute ("update Cashbookprint set BalanceDr=" & Drop & ",BalanceCr=" & CrOp & " where tdate=#" & ldate & "#")
        'End If
        db.Execute ("update Cashbookprint set BalanceDr=" & DAY_TOTAL_dr + Drop & ",BalanceCr=" & DAY_TOTAL_Cr + CrOp & " where tdate=#" & ldate & "#")
        tdate = tdate + 1
        Debug.Print tdate
        adate = Format(tdate, "mm")
        bdate = Format(tdate, "dd")
        rdate = Format(tdate, "yyyy")
        ldate = adate & "/" & bdate & "/" & rdate
    Wend
    db.Close
    Set db = Nothing
    Me.CrystalReport1.ReportFileName = App.Path & "\cashprint.rpt"
    Me.CrystalReport1.PrintReport

End Sub

Private Sub Command2_Click()

End Sub

