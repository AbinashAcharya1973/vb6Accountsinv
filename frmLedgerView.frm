VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmLedgerView 
   BackColor       =   &H00E3E3E3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ledger View"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12975
   Begin VB.CommandButton cmdpdfprint 
      Caption         =   "Print Ledger Statement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1650
      TabIndex        =   25
      Top             =   8250
      Width           =   2640
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
      RecordSource    =   "LedgerTran"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Width           =   12735
      Begin Crystal.CrystalReport CrystalReport2 
         Left            =   5280
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Ledger Statement"
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
         Left            =   300
         TabIndex        =   22
         Top             =   150
         Visible         =   0   'False
         Width           =   2895
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5760
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.Line Line2 
         X1              =   6360
         X2              =   6360
         Y1              =   0
         Y2              =   1440
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Opening Balance :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   21
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lbl_op_dr 
         Alignment       =   1  'Right Justify
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
         Left            =   9120
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lbl_op_cr 
         Alignment       =   1  'Right Justify
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
         Left            =   10920
         TabIndex        =   19
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Total :"
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
         Left            =   6120
         TabIndex        =   18
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lbldr 
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
         Left            =   9120
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblcr 
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
         Left            =   10920
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   6360
         X2              =   12720
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Closing Balance :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lbldr_bal 
         Alignment       =   1  'Right Justify
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
         Left            =   9120
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblcr_bal 
         Alignment       =   1  'Right Justify
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
         Left            =   10920
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   12735
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmLedgerView.frx":0000
         Height          =   5175
         Left            =   120
         OleObjectBlob   =   "frmLedgerView.frx":0014
         TabIndex        =   7
         Top             =   120
         Width           =   12495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   7320
         TabIndex        =   8
         Top             =   0
         Width           =   5295
         Begin MSMask.MaskEdBox txtdateto 
            Height          =   375
            Left            =   3240
            TabIndex        =   9
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
         Begin MSMask.MaskEdBox txtdatefrom 
            Height          =   375
            Left            =   720
            TabIndex        =   10
            Top             =   120
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
            Left            =   2880
            TabIndex        =   12
            Top             =   120
            Width           =   375
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
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.ComboBox cboLedger 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   5175
      End
      Begin VB.ComboBox cboGroup 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label lbladr2 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   1560
         Width           =   5175
      End
      Begin VB.Label lbladr1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Ledger"
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
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLedgerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec1 As DAO.Recordset, rec2 As Recordset
Private Sub cboGroup_Change()
Set rec1 = db.OpenRecordset("select * from LedgerMaster where GroupID=" & Me.cbogroup.ItemData(Me.cbogroup.ListIndex))
Me.cboLedger.Clear
If Not rec1.EOF Then
    While Not rec1.EOF
    Me.cboLedger.AddItem (rec1("AccName"))
    Me.cboLedger.ItemData(Me.cboLedger.NewIndex) = rec1("AccID")
    rec1.MoveNext
    Wend
End If
If Me.cboLedger.ListCount > 0 Then
Me.cboLedger.ListIndex = 0
End If
End Sub

Private Sub cboGroup_Click()
cboGroup_Change
End Sub
Private Sub cboGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cboLedger.SetFocus
End If
End Sub
Private Sub cboLedger_Change()
Me.txtdatefrom.Text = "__/__/____"
Me.txtdateto.Text = "__/__/____"

Me.lblcr.Caption = ""
Me.lblcr_bal.Caption = ""
Me.lbldr.Caption = ""
Me.lbldr_bal.Caption = ""
Set rec1 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
If Not IsNull(rec1!Address1) Then
Me.LblAdr1.Caption = rec1("Address1")
Else
Me.LblAdr1.Caption = ""
End If
Set rec1 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
If Not IsNull(rec1!Address2) Then
Me.lbladr2.Caption = rec1("aDDRESS2")
Else
Me.lbladr2.Caption = ""
End If

Data1.RecordSource = "select * from LedgerTran where AccId=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " order by Tdate,LedgerTran.`Slno` ASC"
Data1.Refresh
Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where AccId=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
If Not IsNull(rs!max_dr) Then
    temp_dr = rs!max_dr
    Me.lbldr.Caption = Format(rs!max_dr, "########0.00")
End If
Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where AccId=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
If Not IsNull(rs!max_cr) Then
    temp_cr = rs!max_cr
    Me.lblcr.Caption = Format(rs!max_cr, "########0.00")
End If
Set rs = db.OpenRecordset("select * from ledgermaster where AccId=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
If Not rs.EOF Then
    If rs("BalanceType") = "Dr" Then
        temp_dr = temp_dr + rs("OBalance")
        Me.lbl_op_dr.Caption = Format(str(Format(rs("OBalance"))), "#######0.00")
        Me.lbl_op_cr.Caption = ""
    End If
    If rs("BalanceType") = "Cr" Then
        temp_cr = temp_cr + rs("OBalance")
        Me.lbl_op_cr.Caption = Format(str(Format(rs("OBalance"))), "#######0.00")
        Me.lbl_op_dr.Caption = ""
    End If
    
End If
If temp_cr > temp_dr Then
    Me.lblcr_bal.Caption = Format(temp_cr - temp_dr, "########0.00")
End If
If temp_dr > temp_cr Then
    Me.lbldr_bal.Caption = Format(temp_dr - temp_cr, "########0.00")
End If

End Sub
Private Sub cboLedger_Click()
cboLedger_Change
End Sub
Private Sub cboLedger_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtdatefrom.SetFocus
End If
End Sub

Private Sub cmdpdfprint_Click()
    Dim objPDF As New clsPDFCreator
    Dim OFile As String
    Dim dht, dwdth As Double
    Dim xbranch As String

    OFile = App.Path & "\ledger.pdf"

    objPDF.Title = "Ledger Statement"
    objPDF.PaperSize = pdfA4
    objPDF.ScaleMode = pdfCentimeter
    objPDF.Margin = 0
    objPDF.Orientation = pdfPortrait

    objPDF.InitPDFFile OFile
    objPDF.LoadFont "Fnt1", "Times New Roman"
    objPDF.LoadFont "Fnt3", "Times New Roman", pdfBold
    objPDF.LoadFont "Fnt4", "Arial Black", pdfBold

    objPDF.BeginPage
    objPDF.DrawText 10.3, 28, "NINE STAR", "Fnt3", 10, pdfCenter
    objPDF.DrawText 10.3, 27.5, "LEDGER STATEMENT", "Fnt4", 14, pdfCenter
    objPDF.DrawText 10.3, 27, Me.cboLedger.Text, "Fnt3", 8, pdfCenter

    objPDF.DrawText 1, 26.1, "Date", "Fnt2", 9.75
    objPDF.DrawText 3.5, 26.1, "Particulars", "Fnt3", 9.75
    objPDF.DrawText 9, 26.1, "Voucher Type", "Fnt4", 9.75
    objPDF.DrawText 12, 26.1, "Voucher No.", "Fnt5", 9.75
    objPDF.DrawText 15, 26.1, "Dr. Amount", "Fnt6", 9.75
    objPDF.DrawText 18, 26.1, "Cr. Amount", "Fnt7", 9.75
    objPDF.DrawLine 0, 26, 22, 26, Stroked, 2
    Dim lineno As Single
    lineno = 25.5
    If Me.txtdatefrom.Text = "__/__/____" And Me.txtdateto.Text = "__/__/____" Then
        Set rec1 = db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " order by Tdate,LedgerTran.`Slno` ASC")
    Else
        objPDF.DrawText 10.3, 26.5, "From :" & Me.txtdatefrom.Text & "   To :" & Me.txtdateto.Text, "Fnt3", 8, pdfCenter
        temp_date1 = Split(Me.txtdatefrom.Text, "/")
        temp_date2 = Split(Me.txtdateto.Text, "/")
        temp_from = temp_date1(1) & "/" & temp_date1(0) & "/" & temp_date1(2)
        temp_to = temp_date2(1) & "/" & temp_date2(0) & "/" & temp_date2(2)
        Set rec1 = db.OpenRecordset("select * from ledgertran where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " and tdate between #" & temp_from & "# and #" & temp_to & "# order by tdate,slno asc")
    End If
    While Not rec1.EOF
        If lineno <= 4 Then
            objPDF.EndPage
            objPDF.BeginPage
            objPDF.DrawText 1, 27, "Date", "Fnt2", 9.75
            objPDF.DrawText 3.5, 27, "Particulars", "Fnt3", 9.75
            objPDF.DrawText 9, 27, "Voucher Type", "Fnt4", 9.75
            objPDF.DrawText 12, 27, "Voucher No.", "Fnt5", 9.75
            objPDF.DrawText 15, 27, "Dr. Amount", "Fnt6", 9.75
            objPDF.DrawText 18, 27, "Cr. Amount", "Fnt7", 9.75
            objPDF.DrawLine 0, 26.5, 22, 26.5, Stroked, 2
            lineno = 26
        End If
        objPDF.DrawText 1, lineno, rec1("tdate"), "Fnt1", 9.75
        objPDF.DrawText 3.5, lineno, rec1("particulars"), "Fnt1", 9.75
        objPDF.DrawText 9, lineno, rec1("vouchertype"), "Fnt1", 9.75
        objPDF.DrawText 13.5, lineno, rec1("voucherslno"), "Fnt1", 9.75, pdfAlignRight
        objPDF.DrawText 17, lineno, Format(rec1("dr"), "######0.00"), "Fnt1", 9.75, pdfAlignRight
        objPDF.DrawText 20, lineno, Format(rec1("cr"), "#######0.00"), "Fnt1", 9.75, pdfAlignRight
        lineno = lineno - 0.4
        rec1.MoveNext
    Wend
    objPDF.DrawLine 0, 4, 22, 4, Stroked, 2
    objPDF.DrawText 10, 3.5, "Opening Balance", "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 17, 3.5, Format(Val(Me.lbl_op_dr.Caption), "######0.00"), "Fnt3", 10, pdfAlignRight
    objPDF.DrawText 20, 3.5, Format(Val(Me.lbl_op_cr.Caption), "#######0.00"), "Fnt3", 10, pdfAlignRight
    objPDF.DrawText 10, 3, "Total", "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 17, 3, Format(Val(Me.lbldr.Caption), "#######0.00"), "Fnt3", 10, pdfAlignRight
    objPDF.DrawText 20, 3, Format(Val(Me.lblcr.Caption), "#######0.00"), "Fnt3", 10, pdfAlignRight
    objPDF.DrawText 10, 2.5, "Closing Balance", "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 17, 2.5, Format(Val(Me.lbldr_bal.Caption), "########0.00"), "Fnt3", 10, pdfAlignRight
    objPDF.DrawText 20, 2.5, Format(Val(Me.lblcr_bal.Caption), "#########0.00"), "Fnt3", 10, pdfAlignRight
    objPDF.EndPage


    '##########################################################################

    objPDF.ClosePDFFile



    Call Shell("rundll32.exe url.dll,FileProtocolHandler " & (OFile), vbMaximizedFocus)

End Sub

Private Sub cmdprint_Click()

If Me.txtdatefrom.Text = "__/__/____" And Me.txtdateto.Text = "__/__/____" Then
CrystalReport1.SelectionFormula = "{LedgerMaster.AccID} =" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex)
CrystalReport1.PrintReport
End If
If Me.txtdatefrom.Text <> "__/__/____" And Me.txtdateto.Text <> "__/__/____" Then
temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)
                
temp_to_date = Mid(Me.txtdateto.Text, 1, 2)
temp_to_month = Mid(Me.txtdateto.Text, 4, 2)
temp_to_year = Mid(Me.txtdateto.Text, 7, 4)

    
End If
Set rec1 = db.OpenRecordset("select * from ledgertran where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
If rec1.EOF Then
    db.Close
    CrystalReport3.SelectionFormula = "{LedgerMaster.AccID} = " & Me.cboLedger.ItemData(Me.cboLedger.ListIndex)
    CrystalReport3.PrintReport
    
Else
    db.Close
    CrystalReport2.SelectionFormula = "{LedgerMaster.AccID} =" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " and {LedgerTran.TDate} in Date (" & temp_from_year & "," & temp_from_month & "," & temp_from_date & ") to Date (" & temp_to_year & "," & temp_to_month & "," & temp_to_date & ")"
    CrystalReport2.PrintReport
End If
Set db = OpenDatabase(dbname)
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
Me.Top = 0
Me.Left = 0
Me.CrystalReport1.ReportFileName = App.Path & "\ledger.rpt"
Me.CrystalReport2.ReportFileName = App.Path & "\ledger_dtwise.rpt"
Data1.databasename = dbname
Set rec1 = db.OpenRecordset("select * from Groups")
While Not rec1.EOF
Me.cbogroup.AddItem (rec1("GroupName"))
Me.cbogroup.ItemData(Me.cbogroup.NewIndex) = rec1("GroupID")
rec1.MoveNext
Wend
If Me.cbogroup.ListCount > 0 Then
Me.cbogroup.ListIndex = 0
End If
Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtdatefrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
        temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
        temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)
        
        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year
        
        temp_to_date = Mid(AccountingPeriod, 1, 2)
        temp_to_month = Mid(AccountingPeriod, 4, 2)
        temp_to_year = Mid(AccountingPeriod, 7, 4)
        
        temp_to = temp_to_month & "/" & temp_to_date & "/" & temp_to_year
If temp_from < temp_to Then
MsgBox "Out of Date", vbCritical
Else
Me.txtdateto.SetFocus
End If

End If
End Sub

Private Sub txtdateto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Me.txtdatefrom.Text <> "__/__/____" And Me.txtdateto.Text <> "__/__/____" Then
        
        Me.lblcr.Caption = ""
        Me.lblcr_bal.Caption = ""
        Me.lbldr.Caption = ""
        Me.lbldr_bal.Caption = ""
        Me.lbl_op_cr.Caption = ""
        Me.lbl_op_dr.Caption = ""
        temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
        temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
        temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)
        
        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year
        
        temp_to_date = Mid(Me.txtdateto.Text, 1, 2)
        temp_to_month = Mid(Me.txtdateto.Text, 4, 2)
        temp_to_year = Mid(Me.txtdateto.Text, 7, 4)
        
        temp_to = temp_to_month & "/" & temp_to_date & "/" & temp_to_year
        
        Data1.RecordSource = "select * from ledgertran where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " and tdate between #" & temp_from & "# and #" & temp_to & "# order by tdate"
        Data1.Refresh
    End If
    'FINDING OPENING BALANCE-------------FOR THE DATE---------
    Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " AND TDATE<#" & temp_from & "#")
    If Not IsNull(rs!max_dr) Then
        temp_dr = rs!max_dr
    Else
        temp_dr = 0
    End If
    Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " AND TDATE<#" & temp_from & "#")
    If Not IsNull(rs!max_cr) Then
        temp_cr = rs!max_cr
    Else
        temp_cr = 0
    End If
    Set rs = db.OpenRecordset("select * from ledgermaster where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
    If Not rs.EOF Then
        If rs("BalanceType") = "Dr" Then
            temp_dr = temp_dr + rs("OBalance")
        End If
        If rs("BalanceType") = "Cr" Then
            temp_cr = temp_cr + rs("OBalance")
        End If
    
    End If
    db.Execute ("update ledgermaster set from_dt='" & Me.txtdatefrom.Text & "',to_dt='" & Me.txtdateto.Text & "',dr_op=0,cr_op=0")
    If temp_cr > temp_dr Then
        Me.lbl_op_cr.Caption = Format(temp_cr - temp_dr, "########0.00")
        db.Execute ("update ledgermaster set from_dt='" & Me.txtdatefrom.Text & "',to_dt='" & Me.txtdateto.Text & "',dr_op=0,cr_op=" & Val(Me.lbl_op_cr.Caption))
    End If
    If temp_dr > temp_cr Then
        Me.lbl_op_dr.Caption = Format(temp_dr - temp_cr, "########0.00")
        db.Execute ("update ledgermaster set from_dt='" & Me.txtdatefrom.Text & "',to_dt='" & Me.txtdateto.Text & "',dr_op=" & Val(Me.lbl_op_dr.Caption) & ",cr_op=0")
    End If
    'FINDING CLOSING BALANCE------------FOR THE DATE----------
    Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " and tdate between #" & temp_from & "# and #" & temp_to & "#")
    If Not IsNull(rs!max_dr) Then
        temp_dr = rs!max_dr
        Me.lbldr.Caption = Format(rs!max_dr, "########0.00")
    End If
    Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex) & " and tdate between #" & temp_from & "# and #" & temp_to & "#")
    If Not IsNull(rs!max_cr) Then
        temp_cr = rs!max_cr
        Me.lblcr.Caption = Format(rs!max_cr, "########0.00")
    End If
    Set rs = db.OpenRecordset("select * from ledgermaster where accid=" & Me.cboLedger.ItemData(Me.cboLedger.ListIndex))
    If Not rs.EOF Then
        If Me.lbl_op_dr.Caption <> "" Then
            temp_dr = temp_dr + Val(Me.lbl_op_dr.Caption)
            Me.lbl_op_cr.Caption = ""
        End If
        If Me.lbl_op_cr.Caption <> "" Then
            temp_cr = temp_cr + Val(Me.lbl_op_cr.Caption)
            Me.lbl_op_dr.Caption = ""
        End If
    End If
    If temp_cr > temp_dr Then
        Me.lblcr_bal.Caption = Format(temp_cr - temp_dr, "########0.00")
    End If
    If temp_dr > temp_cr Then
        Me.lbldr_bal.Caption = Format(temp_dr - temp_cr, "########0.00")
    End If

End If

End Sub
