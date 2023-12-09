VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frminvoicewiseprofit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice-Wise Profit"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   825
      Width           =   1815
   End
   Begin MSMask.MaskEdBox txtfrom 
      Height          =   375
      Left            =   900
      TabIndex        =   0
      Top             =   225
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtto 
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   225
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Height          =   375
      Left            =   3300
      TabIndex        =   3
      Top             =   225
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Left            =   300
      TabIndex        =   2
      Top             =   225
      Width           =   735
   End
End
Attribute VB_Name = "frminvoicewiseprofit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As DAO.Recordset, rec3 As DAO.Recordset
Private Sub cmdprint_Click()
    Dim clpdf As New clsPDFCreator
    Dim strFile As String
    strFile = App.Path & "\InvoiceProfit.pdf"
    clpdf.Title = "Invoice"       ' Titolo
    clpdf.ScaleMode = pdfCentimeter          ' Unità di misura
    clpdf.PaperSize = pdfA4                  ' Formato pagina
    clpdf.Margin = 0                         ' Margine
    clpdf.Orientation = pdfPortrait          ' Orientamento

    'clpdf.EncodeASCII85 = (chkASCII85.Value = Checked)

    clpdf.InitPDFFile strFile                ' inizializza il file

    ' Definisce le risorse relative ai font
    clpdf.LoadFont "Fnt1", "Times New Roman"                       ' Tipo TrueType
    clpdf.LoadFont "Fnt2", "Arial", pdfItalic                      ' Tipo TrueType
    clpdf.LoadFont "Fnt3", "Courier New", pdfBold                          ' Tipo TrueType
    clpdf.LoadFontStandard "Fnt4", "Courier New", pdfBold    ' Tipo Type1
    clpdf.LoadFont "Fnt5", "Courier New"
    clpdf.LoadFont "Fnt6", "Times New Roman"

    ' Definisce una risorsa comune da stampare solo sulle pagine pari

    from_dt = Split(Me.txtfrom.Text, "/")
    to_dt = Split(Me.txtto.Text, "/")
    '     Inizializza la prima pagina
    clpdf.BeginPage

    clpdf.DrawText 2, 1, "Page: " & Trim(CStr(clpdf.Pages)), "Fnt1", 8, pdfAlignRight
    clpdf.DrawObject "Footers"
    clpdf.DrawText 10.5, 28.5, "Invoice Wise Profit", "Fnt1", 24, pdfCenter
    clpdf.DrawLine 0, 27.4, 21, 27.4, Stroked, 1

    clpdf.DrawText 1, 27, "Invoice Date", "Fnt3", 10, pdfAlignLeft
    clpdf.DrawText 4, 27, "Invoice No", "Fnt3", 10, pdfAlignLeft
    clpdf.DrawText 6.5, 27, "Party Name", "Fnt3", 10, pdfAlignLeft
    clpdf.DrawText 14, 27, "Gross Sales", "Fnt3", 10, pdfAlignLeft
    clpdf.DrawText 17.5, 27, "Gross Margin", "Fnt3", 10, pdfAlignLeft
    clpdf.DrawLine 0, 26.8, 21, 26.8, Stroked, 1
    Dim lineno As Single
    lineno = 26
    total_grossmargin = 0
    Set rec1 = db.OpenRecordset("select * from invoicehead where InvoiceHead.InvDate between #" & from_dt(1) & "/" & from_dt(0) & "/" & from_dt(2) & "# and #" & to_dt(1) & "/" & to_dt(0) & "/" & to_dt(2) & "#")
    While Not rec1.EOF
        Set rec2 = db.OpenRecordset("select * from invoicedetails where invno=" & rec1("invno"))
        gross_purchase = 0
        gross_sale = 0
        While Not rec2.EOF
            Set rec3 = db.OpenRecordset("select max(slno) as m_slno from purchasedetails where productcode=" & rec2("productcode"))
            If Not IsNull(rec3!m_slno) Then
                temp_purchaseslno = rec3!m_slno
            Else
                temp_purchaseslno = 0
            End If
            If temp_purchaseslno <> 0 Then
                Set rec3 = db.OpenRecordset("select * from purchasedetails where slno=" & temp_purchaseslno & " and productcode=" & rec2("productcode"))
                If Not rec3.EOF Then
                    gross_purchase = gross_purchase + ((rec3("prrate") - (rec3("prrate") * (rec3("discount") / 100))) * rec2("qty"))
                    gross_sale = gross_sale + (rec2("gross") - rec2("discountamount"))
                End If
            Else
                Set rec3 = db.OpenRecordset("select * from itemmaster where productcode=" & rec2("productcode"))
                If Not rec3.EOF Then
                    gross_purchase = gross_purchase + (rec3("purchaserate") * rec2("qty"))
                    gross_sale = gross_sale + (rec2("gross") - rec2("discountamount"))
                End If
            End If
            rec2.MoveNext
        Wend
        If lineno <= 1.5 Then
            clpdf.EndPage
            clpdf.BeginPage
            clpdf.DrawLine 0, 27.4, 21, 27.4, Stroked, 1

            clpdf.DrawText 1, 27, "Invoice Date", "Fnt3", 10, pdfAlignLeft
            clpdf.DrawText 4, 27, "Invoice No", "Fnt3", 10, pdfAlignLeft
            clpdf.DrawText 6.5, 27, "Party Name", "Fnt3", 10, pdfAlignLeft
            clpdf.DrawText 14, 27, "Gross Sales", "Fnt3", 10, pdfAlignLeft
            clpdf.DrawText 17.5, 27, "Gross Margin", "Fnt3", 10, pdfAlignLeft
            clpdf.DrawLine 0, 26.8, 21, 26.8, Stroked, 1
            lineno = 26
        End If
        clpdf.DrawText 1, lineno, rec1("invdate"), "Fnt1", 10, pdfAlignLeft
        clpdf.DrawText 4, lineno, rec1("invno"), "Fnt1", 10, pdfAlignLeft
        clpdf.DrawText 6.5, lineno, rec1("party"), "Fnt1", 10, pdfAlignLeft
        clpdf.DrawText 6.5, lineno, rec1("party"), "Fnt1", 10, pdfAlignLeft
        clpdf.DrawText 14, lineno, str(gross_sale), "Fnt1", 10, pdfAlignLeft
        clpdf.DrawText 17.5, lineno, gross_sale - gross_purchase, "Fnt1", 10, pdfAlignLeft
        total_grossmargin = total_grossmargin + (gross_sale - gross_purchase)
        lineno = lineno - 0.5
        rec1.MoveNext
    Wend
    clpdf.DrawLine 0, 1.5, 21, 1.5, Stroked, 1
    clpdf.DrawText 17.5, 1, str(total_grossmargin), "Fnt3", 10, pdfAlignLeft
    clpdf.EndPage
    clpdf.ClosePDFFile
    Call Shell("rundll32.exe url.dll,FileProtocolHandler " & (strFile), vbMaximizedFocus)
End Sub

