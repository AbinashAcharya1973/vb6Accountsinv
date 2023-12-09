VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmInvPrint 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Print"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdpushinv 
      Caption         =   "Push Inv"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "InvoiceHead"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7395
      Begin VB.CommandButton cmdsupply 
         Caption         =   "Bill of Supply"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5220
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdmail 
         Caption         =   "Mail Inv"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   7
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdgst 
         Caption         =   "Print Inv"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdtear 
         Caption         =   "Tear Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6780
         TabIndex        =   5
         Top             =   2430
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmInvoicePrint.frx":0000
         Height          =   1695
         Left            =   120
         OleObjectBlob   =   "FrmInvoicePrint.frx":0014
         TabIndex        =   4
         Top             =   600
         Width           =   7155
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5280
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         PrinterName     =   "Microsoft Print to PDF"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PRINT"
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
         Left            =   -690
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CboInvType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Inv Type"
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
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmInvPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinHttpReq As WinHttp.WinHttpRequest
      Private Const LOCALE_SSHORTDATE = &H1F
      Private Const WM_SETTINGCHANGE = &H1A
      'same as the old WM_WININICHANGE
      Private Const HWND_BROADCAST = &HFFFF&

      Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
          "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
          Long, ByVal lpLCData As String) As Boolean
      Private Declare Function PostMessage Lib "user32" Alias _
          "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function GetSystemDefaultLCID Lib "kernel32" _
          () As Long

Dim rec1 As DAO.Recordset, db2 As DAO.Database, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset
Private Sub cboInvType_Change()
    Data1.RecordSource = "select * from InvoiceHead where invtype='" & Me.CboInvType.Text & "' order by invno desc"
    Me.Data1.Refresh
End Sub
Private Sub cboInvType_Click()
    cboInvType_Change
End Sub

Private Sub cboprinter_Click()
    If SelectPrinter(cboprinter.Text) Then
        MsgBox "Printer not found"
    End If
End Sub

Private Sub cmdgst_Click()
    'MsgBox tempemailid
    'Me.CommonDialog1.Action = 5
    'db.Execute "UPDATE INVOICEHEAD SET INVPF='" + Format(Me.DBGrid1.Columns(0), "0000") & "/ABSPVL/" & ACYEAR & "' WHERE INVNO=" & Me.DBGrid1.Columns(0)
    'frmprinter.Show vbModal
    Me.CrystalReport1.ReportFileName = App.Path & "\Invoicegst.rpt"
    'Me.CrystalReport1.PrinterName = Printer.DeviceName
    'Me.CrystalReport1.PrinterDriver = Printer.DriverName
    'Me.CrystalReport1.PrinterPort = Printer.Port
    
    CrystalReport1.SelectionFormula = "{InvoiceHead.InvNo}=" & Val(Me.DBGrid1.Columns(0)) & " and {InvoiceHead.InvType}='" & Me.CboInvType.Text & "'"
    CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.PrintReport
End Sub

Private Sub cmdmail_Click()
On Error GoTo errtrap
    Dim objPDF As New clsPDFCreator

    Set rec1 = db.OpenRecordset("select * from invoicehead where invno=" & Val(Me.DBGrid1.Columns(0)))
    Set rec2 = db.OpenRecordset("select * from ledgermaster where accid=" & rec1("accid"))
    Set rec3 = db.OpenRecordset("select * from InvoiceDetails where invno=" & Val(Me.DBGrid1.Columns(0)) & " order by slno")
    Set rec4 = db.OpenRecordset("select * from companymaster")
    objPDF.Title = "INVOICE"
    objPDF.PaperSize = pdfA4
    objPDF.ScaleMode = pdfCentimeter
    objPDF.Margin = 0
    objPDF.Orientation = pdfPortrait
    Dim attachments As New Collection
    objPDF.InitPDFFile App.Path & "\invoice.pdf"
    objPDF.LoadFont "Fnt1", "Times New Roman"
    objPDF.LoadFont "Fnt3", "Times New Roman", pdfBold

    objPDF.LoadFont "Fnt4", "Arial Black", pdfBold
    objPDF.BeginPage
    objPDF.DrawText 20, 29, "Credit ", "Fnt3", 15, pdfAlignRight
    objPDF.DrawLine 1, 28.5, 20, 28.5, Filled, pdfCenter    '1
    objPDF.DrawLine 1, 25.5, 20, 25.5, Filled, pdfCenter    '2
    objPDF.DrawLine 1, 22.5, 20, 22.5, Filled, pdfCenter    '3
    objPDF.DrawLine 1, 21, 20, 21, Filled, pdfCenter    '4
    objPDF.DrawLine 1, 20.3, 20, 20.3, Filled, pdfCenter    '5
    objPDF.DrawLine 1, 4.7, 20, 4.7, Filled, pdfCenter    '6
    objPDF.DrawLine 1, 5.3, 20, 5.3, Filled, pdfCenter    '7
    objPDF.DrawLine 1, 1.7, 20, 1.7, Filled, pdfCenter    '8
    objPDF.DrawLine 1, 1.3, 20, 1.3, Filled, pdfCenter    '9
    objPDF.DrawLine 1, 28.5, 1, 22.5, Filled, pdfCenter    '1
    objPDF.DrawLine 20, 28.5, 20, 22.5, Filled, pdfCenter    '2
    objPDF.DrawLine 10, 28.5, 10, 22.5, Filled, pdfCenter    '3


    objPDF.DrawText 1.3, 28, rec4("company"), "Fnt3", 15, pdfAlignLeft
    objPDF.DrawText 1.2, 27.4, rec4("address"), "Fnt3", 7.5, pdfAlignLeft
    objPDF.DrawText 1.2, 27, "PHONE : : " & rec4("phone"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 1.2, 26.3, "EMAIL: " & rec4("email"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 1.2, 25.7, "GSTIN" & rec4("taxno"), "Fnt3", 10, pdfAlignLeft

    objPDF.DrawText 10.5, 28, "Deals in :", "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 10.5, 27.5, "Bank Details:" & rec4("bankname"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 10.5, 26.5, "", "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 10.5, 25.9, "A/c No.:" & rec4("bankacno") & ", IFSC: " & rec4("ifsc"), "Fnt3", 10, pdfAlignLeft


    objPDF.DrawText 1.3, 25, "TO  :" & rec2("AccName"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 1.3, 24.5, "  " & rec2("Address1"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 1.3, 23.8, " " & rec2("Address2"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 1.3, 23.3, " " & rec2("Tin"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 1.3, 22.9, "STATE CODE  :" & rec2("StateCode"), "Fnt3", 10, pdfAlignLeft

    objPDF.DrawText 10.5, 25, "TAX      INVOICE", "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 10.5, 24.5, "Inv No  :" & rec1("invno"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 10.5, 23.8, "Inv Date :" & rec1("invdate"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 10.5, 23.3, "Ch No & Dt  :" & rec1("ChalanDate") & rec1("ChalanNo"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 10.5, 22.9, "L.R No" & rec1("LrNo"), "Fnt3", 10, pdfAlignLeft
    objPDF.DrawText 1, 20.7, "Sl.No", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 1.7, 20.7, "Product Name", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 4.2, 20.7, "HSN", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 6.5, 20.7, "Qty", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 7.5, 20.7, "Rate", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 9, 20.7, "Gross", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 10.6, 20.7, "CD", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 11.6, 20.7, "Dsc%", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 12.9, 20.7, "SGST", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 14.4, 20.7, "CGST", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 15.9, 20.7, "Tax Amt", "Fnt3", 8, pdfAlignLeft
    objPDF.DrawText 17.9, 20.7, "Net Amt", "Fnt3", 8, pdfAlignLeft

    Dim TROW As Single
    TROW = 20
    While Not rec3.EOF
        If TROW < 4 Then
            objPDF.BeginPage
            TROW = 28
            objPDF.DrawText 1, TROW, "Sl.No   Product Name       HSN          Quantity                 Rate            MRP           CD               Dsc%        SGST          CGST       Tax Amt     Net Amt", "Fnt3", 8, pdfAlignLeft
            TROW = TROW - 0.5
        End If
        objPDF.DrawText 1.3, TROW, " " & rec3("Slno"), "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 1.7, TROW, " " & rec3("Itemname"), "Fnt3", 8, pdfAlignLeft

        X = 0


        objPDF.DrawText 6.5, TROW, " " & (rec3("Qty") + rec3("Free_Qty")), "Fnt3", 8, pdfAlignLeft
        'objPDF.DrawText 6.6, TROW, " " & rec3("Qty"), "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 7.5, TROW, " " & rec3("SaleRate"), "Fnt3", 8, pdfAlignLeft
        X = X + rec3("SaleRate")
        'x = sum + (rec3("Qty") + rec3("Free_Qty"))
        objPDF.DrawText 9, TROW, " " & Round(rec3("qty") * rec3("salerate"), 2), "Fnt3", 8, pdfAlignLeft
        'objPDF.DrawText 7.6, 20, " " & rec3("Tradediscount"), "Fnt3", 8, pdfAlignLeft
        'objPDF.DrawText 9, 20, " " & rec3("SpecialDiscount"), "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 10.6, TROW, " " & rec3("Tradediscount"), "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 11.6, TROW, " " & rec3("SpecialDiscount"), "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 12.9, TROW, " " & rec3("Vat") / 2, "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 14.4, TROW, " " & rec3("Vat") / 2, "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 15.9, TROW, " " & rec3("VatAmount"), "Fnt3", 8, pdfAlignLeft
        objPDF.DrawText 17.9, TROW, " " & rec3("Net"), "Fnt3", 8, pdfAlignLeft
        sum1 = 0
        sum1 = sum1 + rec3("SpecialDiscount")
        netsum = 0
        netsum = netsum + rec3("Net")
        sgst = 0
        sgst = sgst + rec3("Vat") / 2
        cgst = 0
        cgst = cgst + rec3("Vat") / 2
        TROW = TROW - 0.5
        rec3.MoveNext
    Wend

    objPDF.DrawText 1.3, 4.9, "TOTAL", "Fnt3", 9, pdfAlignLeft

    objPDF.DrawText 6.5, 4.9, rec1("totalqty"), "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 9, 4.9, rec1("totalgross"), "Fnt3", 9, pdfAlignLeft

    objPDF.DrawText 15.9, 4.9, rec1("vatamount"), "Fnt3", 9, pdfAlignLeft
    'objPDF.DrawText 11.6, 4.9, " " & sum1, "Fnt3", 9, pdfAlignLeft
    'objPDF.DrawText 14.7, 4.9, " " & sgst, "Fnt3", 9, pdfAlignLeft
    'objPDF.DrawText 16, 4.9, " " & cgst, "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 17.4, 4.9, " " & rec1("net"), "Fnt3", 9, pdfAlignLeft
    'objPDF.DrawText 18.5, 4.9, " " & netsum, "Fnt3", 9, pdfAlignLeft

    objPDF.DrawText 14.7, 4.3, "Taxable", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 14.7, 3.8, "OGST", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 14.7, 3.3, "CGST", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 14.7, 2.8, "IGST", "Fnt3", 9, pdfAlignLeft

    objPDF.DrawText 18.5, 4.3, rec1("totalgross"), "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 18.5, 3.8, rec1("vatamount") / 2, "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 18.5, 3.3, rec1("vatamount") / 2, "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 18.5, 2.8, " 0.00", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 14.2, 2.6, " --------------------------------------------------------", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 18.5, 2.4, " " & rec1("Freight"), "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 14.5, 2.1, " Round off", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 14.5, 1.8, " Invoice Amount", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 18.5, 2.1, " " & rec1("RndUp"), "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 18.5, 1.8, " " & rec1("GrandTotal"), "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 3.5, 1.4, rec1("amountintext"), "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 2, 0.3, "All The Disputes sebject to " & rec4("jurisdiction") & " Jurisdiction", "Fnt3", 9, pdfAlignLeft
    objPDF.DrawText 14.5, 0.5, "For " & rec4("company"), "Fnt3", 11, pdfAlignLeft
    objPDF.ClosePDFFile
    
    Set rec1 = db.OpenRecordset("select * from partydr where accid=" & rec1("accid"))
    If Not rec1.EOF Then
        partyemailid = LCase(rec1("email"))
    End If
    partyemailid = frminstmsg.GetEmailID("Email ID", "Instant Mail ID", partyemailid)

    strFile = App.Path & "\Invoice.pdf"
    attachments.Add strFile
    Me.MousePointer = vbHourglass
    'Call Shell("rundll32.exe url.dll,FileProtocolHandler " & App.Path & "\INVOICE.PDF", vbMaximizedFocus)
    X = SendEmail("souravch30@gmail.com", partyemailid, "Invoice No-" & Me.DBGrid1.Columns(1), "Kindly Download the Attached Invoice", "", "", attachments)
    Me.MousePointer = vbDefault
    MsgBox "Email has been Sent ", vbOKOnly
    Exit Sub
errtrap:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdpushinv_Click()
    MsgBox "Connect Internet to Push the Invoice to Global Database", vbOKOnly
    Dim PartyRegistered As Boolean, fromPartyGID, ToPartyGID
    Dim p As Object, strRes As String, JSONrecordsH As String, JSONrecordsD As String
    PartyRegistered = False
    Set rec1 = db.OpenRecordset("select * from companymaster")
    If rec1("gid") = 0 Then
        WinHttpReq.Open "GET", _
                        "http://techspark.xp3.biz/enlite/getgid.php?Mobile=" & rec1("phone") & "&gstno=" & rec1("taxno"), False
        WinHttpReq.Send
        If WinHttpReq.ResponseText Like "*Not Found*" Then
            MsgBox "Not Found", vbCritical
        Else
            strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
            Set p = JSON.parse(strRes)
            db.Execute "update companymaster set gid=" & p.Item("CompanyID")
            fromPartyGID = p.Item("CompanyID")
            MsgBox "Global ID Updated", vbOKOnly
        End If
    Else
        fromPartyGID = rec1("gid")
    End If
    Set rec1 = db.OpenRecordset("select * from invoicehead inner join partydr on invoicehead.accid=partydr.accid where invoicehead.invno=" & Me.DBGrid1.Columns(0))
    If Not rec1.EOF Then
        If rec1("gid") = 0 Then
            WinHttpReq.Open "GET", _
                            "http://techspark.xp3.biz/enlite/getgid.php?Mobile=" & rec1("phone") & "&gstno=" & rec1("tin"), False
            WinHttpReq.Send
            If WinHttpReq.ResponseText Like "*Not Found*" Then
                MsgBox "The Party is Not Registered", vbCritical
            Else
                strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
                Set p = JSON.parse(strRes)
                db.Execute "update partydr set gid=" & p.Item("CompanyID") & " where accid=" & rec1("invoicehead.accid")
                MsgBox "Global ID Updated", vbOKOnly
                ToPartyGID = p.Item("CompanyID")
                PartyRegistered = True
            End If
        Else
            PartyRegistered = True
            ToPartyGID = rec1("gid")
        End If
    End If
    If PartyRegistered = True Then
        Set rec1 = db.OpenRecordset("select * from invoicehead where invno=" & Me.DBGrid1.Columns(0))
        JSONrecordsH = JSON.RStoJSON(rec1)
        MsgBox JSONrecordsH, vbOKOnly
        Set rec1 = db.OpenRecordset("select InvNo,InvType,itemmaster.ProductType as ProductType,itemmaster.ItemType as ItemType,itemmaster.Brand as Brandname,Itemname,itemmaster.Size as Size,Units,itemmaster.MRP as MRP,invoicedetails.SaleRate as SaleRate,Qty,Gross,SpecialDiscount,Tradediscount,DiscountAmount,Vat,VatAmount,Net,invoicedetails.ProductCode as ProductCode,Free_Qty,invoicedetails.Tax_type as Tax_type,invoicedetails.Pack as Pack,Slno,itemmaster.HSN as HSN,MfgDate,ExpDate,BatchNo,adapterslno,batteryslno from invoicedetails inner join itemmaster on invoicedetails.productcode=itemmaster.productcode where invoicedetails.invno=" & Me.DBGrid1.Columns(0))
        JSONrecordsD = JSON.RStoJSON(rec1)
        MsgBox JSONrecordsD, vbOKOnly
        WinHttpReq.Open "GET", _
                        "http://techspark.xp3.biz/enlite/receiveinv.php?invhead=" & JSONrecordsH & "&invdetails=" & JSONrecordsD & "&fromP=" & fromPartyGID & "&toP=" & ToPartyGID, False
        WinHttpReq.Send
        resultx = WinHttpReq.ResponseText
        'Debug.Print JSONrecords
        'db.Execute ("update invoicehead set pushed='Y' where invno=" & Me.DBGrid1.Columns(0))
        Me.DBGrid1.Columns(4) = "Y"
        Me.Data1.UpdateRecord
    End If
End Sub

Private Sub cmdsupply_Click()
    Me.CrystalReport1.ReportFileName = App.Path & "\billofsupply.rpt"
    CrystalReport1.SelectionFormula = "{InvoiceHead.InvNo}=" & Val(Me.DBGrid1.Columns(0)) & " and {InvoiceHead.InvType}='" & Me.CboInvType.Text & "'"
    CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.PrintReport
End Sub

Private Sub cmdtear_Click()
  Me.CrystalReport1.ReportFileName = App.Path & "\tearinv.rpt"
    CrystalReport1.SelectionFormula = "{InvoiceHead.InvNo}=" & Val(Me.DBGrid1.Columns(0)) & " and {InvoiceHead.InvType}='" & Me.CboInvType.Text & "'"
    CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.PrintReport
End Sub

Private Sub Command1_Click()
'    Me.CrystalReport1.ReportFileName = App.Path & "\Invoice.rpt"
'    CrystalReport1.SelectionFormula = "{InvoiceHead.InvNo}=" & Val(Me.DBGrid1.Columns(0)) & " and {InvoiceHead.InvType}='" & Me.CboInvType.Text & "'"
'    CrystalReport1.WindowState = crptMaximized
'    Me.CrystalReport1.PrintReport
Dim iFileNo As Integer
        iFileNo = FreeFile
        Open "D:\Test.json" For Output As #iFileNo
        Set rec1 = db.OpenRecordset("select * from invoicehead where invno=" & Me.DBGrid1.Columns(0))
        JSONrecordsH = JSON.RStoJSON(rec1)
        Print #iFileNo, JSONrecordsH
        Close #iFileNo
End Sub


Private Sub Form_Load()
On Error GoTo errtrap
    Dim i As Integer
    Me.Top = 3300
    Me.Left = 2500
    Me.Data1.databasename = dbname
    Set db2 = OpenDatabase(dbname)
    Set rec1 = db.OpenRecordset("select Distinct InvType From Invoicehead")
    While Not rec1.EOF
        Me.CboInvType.AddItem (rec1("InvType"))
        rec1.MoveNext
    Wend
    If Me.CboInvType.ListCount > 0 Then
        Me.CboInvType.ListIndex = 0
    End If
    For i = 0 To Printers.Count - 1
        cboprinter.AddItem Printers(i).DeviceName
    Next i
    Set WinHttpReq = New WinHttpRequest
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Function SelectPrinter(ByVal printer_name As String) As Boolean
    Dim i As Integer
 
    SelectPrinter = True
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = printer_name Then
            Set Printer = Printers(i)
            SelectPrinter = False
            Exit For
        End If
    Next i
End Function
