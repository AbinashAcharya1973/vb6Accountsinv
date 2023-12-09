VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmprintjournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Journal"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdprint 
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
      Height          =   345
      Left            =   3195
      TabIndex        =   2
      Top             =   3090
      Width           =   1665
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4365
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "JournalDetails"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmprintjournal.frx":0000
      Height          =   2385
      Left            =   90
      OleObjectBlob   =   "frmprintjournal.frx":0014
      TabIndex        =   0
      Top             =   570
      Width           =   7965
   End
End
Attribute VB_Name = "frmprintjournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As Recordset

Private Sub cmdprint_Click()
    PDFprint Me.DBGrid1.Columns(0), Me.DBGrid1.Columns(1)

End Sub

Private Sub Form_Load()
    Data1.databasename = dbname
    Data1.RecordSource = "select * from journaldetails order by jdate,slno"
    Data1.Refresh
    Me.DBGrid1.Refresh
End Sub
Public Sub PDFprint(rcno, rcdate As Date)
    Dim objPDF As New clsPDFCreator
    Dim OFile As String
    Dim dht, dwdth As Double, cname As String
    Dim xbranch As String, tempamount As String
    Dim lineno As Single, caddress As String, cpin As String
    OFile = App.Path & "\journal.pdf"

    objPDF.Title = "Journal Voucher"
    objPDF.PaperSize = pdfA4
    'objPDF.PaperWidth = 595.2
    'objPDF.PaperHeight = 280.5

    objPDF.ScaleMode = pdfCentimeter
    objPDF.Margin = 0
    objPDF.Orientation = pdfPortrait

    objPDF.InitPDFFile OFile
    objPDF.LoadFont "Fnt1", "Times New Roman"
    objPDF.LoadFont "Fnt3", "Times New Roman", pdfBold
    objPDF.LoadFont "Fnt4", "Arial Black", pdfBold
    objPDF.LoadFont "Fnt5", "Times New Roman", pdfBoldItalic

    objPDF.BeginPage
    
    Set rec1 = db.OpenRecordset("select * from companymaster")
    If Not rec1.EOF Then
        cname = rec1("company")
        caddress = rec1("address") & "'," & rec1("address1")
        cpin = IIf(IsNull(rec1("pin")), "", rec1("pin"))
    End If
    
    Set rec1 = db.OpenRecordset("select * from journaldetails where slno=" & rcno & " and jdate=#" & Format(rcdate, "mm/dd/yyyy") & "#")
    If Not rec1.EOF Then
        tempaccname = rec1("accname")
        tempamount = rec1("amount")
    End If
    
    Set rec1 = db.OpenRecordset("select * from journalhead where slno=" & rcno & " and jdate=#" & Format(rcdate, "mm/dd/yyyy") & "#")
    If Not rec1.EOF Then
        tempnaration = rec1("narration")
    End If
    
    lineno = 28
    objPDF.DrawText 10.3, lineno, cname, "Fnt4", 14, pdfCenter
    lineno = lineno - 0.5
    objPDF.DrawText 10.3, lineno, caddress, "Fnt4", 10, pdfCenter
    lineno = lineno - 0.6
    objPDF.DrawText 10.3, lineno, "PIN:" & cpin, "Fnt4", 10, pdfCenter
    lineno = lineno - 0.5
    objPDF.DrawText 10.3, lineno, "JOURNAL VOUCHER", "Fnt4", 10, pdfCenter
    lineno = lineno - 1
    objPDF.DrawText 1, lineno, "Voucher No: " & rcno, "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 14, lineno, "Voucher Date: " & Format(rcdate, "dd/mm/yyyy"), "Fnt3", 10, pdfAlignLeft
    lineno = lineno - 1
    objPDF.DrawText 1, lineno, "DEBITED To " & tempaccname, "Fnt5", 12, pdfAlignLeft
    lineno = lineno - 0.7
    objPDF.DrawText 1, lineno, "Amount " & NumberToWord(tempamount), "Fnt5", 12, pdfAlignLeft
    lineno = lineno - 0.7
    objPDF.DrawText 1, lineno, "In Debit of " & tempnaration, "Fnt5", 12, pdfAlignLeft
    lineno = lineno - 1.3
    objPDF.DrawText 1, lineno, "Rs. " & Format(tempamount, "###,##,###"), "Fnt4", 14, pdfAlignLeft
    lineno = lineno - 0.6
    objPDF.DrawText 1, lineno, "Prepared By ", "Fnt4", 10, pdfAlignLeft
    objPDF.DrawText 16, lineno, "Debit Authorised Signature", "Fnt4", 10, pdfCenter

    objPDF.EndPage


    '##########################################################################

    objPDF.ClosePDFFile



    Call Shell("rundll32.exe url.dll,FileProtocolHandler " & (OFile), vbMaximizedFocus)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rec2 = Nothing
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)

    Me.Data1.RecordSource = "select * from journaldetails where accname like '" & Me.txtsearch.Text & "*' order by jdate,slno"
    Me.Data1.Refresh
    Me.DBGrid1.Refresh
End Sub

