VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmreceiptprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Voucher Printing"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8115
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ReceiptDetails"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmreceiptprint.frx":0000
      Height          =   2385
      Left            =   60
      OleObjectBlob   =   "frmreceiptprint.frx":0014
      TabIndex        =   2
      Top             =   570
      Width           =   8025
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
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   4365
   End
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
      Left            =   3225
      TabIndex        =   0
      Top             =   3090
      Width           =   1665
   End
End
Attribute VB_Name = "frmreceiptprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As Recordset

Private Sub cmdprint_Click()
    PDFprint Me.DBGrid1.Columns(0), Me.DBGrid1.Columns(1)
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()
    Data1.databasename = dbname
    Data1.RecordSource = "select * from receiptdetails order by receiptdate desc,receiptno desc"
    Data1.Refresh
    DBGrid1.Refresh
End Sub
Public Sub PDFprint(rcno, rcdate As Date)
    Dim objPDF As New clsPDFCreator
    Dim OFile As String
    Dim dht, dwdth As Double, cname As String
    Dim xbranch As String, tempamount As String
    Dim lineno As Single, caddress As String, cpin As String
    OFile = App.Path & "\receipt.pdf"

    objPDF.Title = "Money Receipt"
    'objPDF.PaperSize = pdfA4
    objPDF.PaperWidth = 595.2
    objPDF.PaperHeight = 280.5

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
    
    Set rec1 = db.OpenRecordset("select * from receiptdetails where receiptno=" & rcno & " and receiptdate=#" & Format(rcdate, "mm/dd/yyyy") & "#")
    If Not rec1.EOF Then
        tempaccname = rec1("accname")
        tempamount = rec1("amount")
    End If
    
    Set rec1 = db.OpenRecordset("select * from receipthead where receiptno=" & rcno & " and receiptdate=#" & Format(rcdate, "mm/dd/yyyy") & "#")
    If Not rec1.EOF Then
        tempnaration = rec1("narration")
    End If

    

    
    lineno = 8.5
    objPDF.DrawText 10.3, lineno, cname, "Fnt4", 14, pdfCenter
    lineno = lineno - 0.5
    objPDF.DrawText 10.3, lineno, caddress, "Fnt4", 10, pdfCenter
    lineno = lineno - 0.6
    objPDF.DrawText 10.3, lineno, "PIN:" & cpin, "Fnt4", 10, pdfCenter
    lineno = lineno - 0.5
    objPDF.DrawText 10.3, lineno, "RECEIPT", "Fnt4", 10, pdfCenter
    lineno = lineno - 1
    objPDF.DrawText 1, lineno, "Receipt No: " & rcno, "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 14, lineno, "Receipt Date: " & Format(rcdate, "dd/mm/yyyy"), "Fnt3", 10, pdfAlignLeft
    lineno = lineno - 1
    objPDF.DrawText 1, lineno, "Received From " & tempaccname, "Fnt5", 12, pdfAlignLeft
    lineno = lineno - 0.7
    objPDF.DrawText 1, lineno, "Amount " & NumberToWord(tempamount), "Fnt5", 12, pdfAlignLeft
    lineno = lineno - 0.7
    objPDF.DrawText 1, lineno, "For Payment of " & tempnaration, "Fnt5", 12, pdfAlignLeft
    lineno = lineno - 1.3
    objPDF.DrawText 1, lineno, "Rs. " & Format(tempamount, "###,##,###"), "Fnt4", 14, pdfAlignLeft
    lineno = lineno - 0.5
    objPDF.DrawText 16, lineno, "Receiver's Signature", "Fnt4", 10, pdfCenter

    objPDF.EndPage


    '##########################################################################

    objPDF.ClosePDFFile



    Call Shell("rundll32.exe url.dll,FileProtocolHandler " & (OFile), vbMaximizedFocus)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rec2 = Nothing
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    Data1.RecordSource = "select * from receiptdetails where accname like '" & Me.txtsearch.Text & "%' and userid=" & userid & " order by receiptdate desc,receiptno desc"
    Data1.Refresh
    DBGrid1.Refresh
End Sub
