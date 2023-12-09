VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGSTReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GST Reports"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdcreditnote 
      BackColor       =   &H0080C0FF&
      Caption         =   "Credit Note"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1380
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2925
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdpurchase 
      BackColor       =   &H000080FF&
      Caption         =   "Purchase Reg."
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "B2C Report"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdb2b 
      BackColor       =   &H0080C0FF&
      Caption         =   "B2B Report"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin MSMask.MaskEdBox txtfrom 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
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
      Left            =   4080
      TabIndex        =   3
      Top             =   240
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
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   735
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
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmGSTReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec2 As DAO.Recordset
Private Sub cmdb2b_Click()
    Dim strLine As String
    Dim fso As New FileSystemObject
    Dim fsoStream As TextStream

    Me.MousePointer = vbHourglass
    Dim from_dt() As String
    Dim to_dt() As String
    Dim tempfilename
    'Me.CommonDialog1.Filter = "Apps (*.csv)"
    'Me.CommonDialog1.DefaultExt = "csv"
    'Me.CommonDialog1.DialogTitle = "Save CSV file"
    'Me.CommonDialog1.InitDir = App.Path & "\inputfile"
    'Me.CommonDialog1.ShowSave
    'tempfilename = Me.CommonDialog1.FileName
    'Set fsoStream = fso.CreateTextFile(tempfilename, True)
    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelWS.Cells(1, 1).Value = "GSTIN/UIN of Recipient"
    excelWS.Cells(1, 2).Value = "Recipient"
    excelWS.Cells(1, 3).Value = "Address"
    excelWS.Cells(1, 4).Value = "InvoiceNo"
    excelWS.Cells(1, 5).Value = "Invoice Date"
    excelWS.Cells(1, 6).Value = "Invoice Value"
    excelWS.Cells(1, 7).Value = "Place Of Supply"
    excelWS.Cells(1, 8).Value = "Reverse Charge"
    excelWS.Cells(1, 9).Value = "Invoice Type"
    excelWS.Cells(1, 10).Value = "E-Commerce GSTIN"
    excelWS.Cells(1, 11).Value = "Rate"
    excelWS.Cells(1, 12).Value = "Taxable Value"
    excelWS.Cells(1, 13).Value = "Tax Amount"
    strLine = "GSTIN/UIN of Recipient,Recipent,Address,Invoice No,Invoice Date,Invoice Value,Place Of Supply,Reverse Charge,Invoice type,E-CommerceGSTIN,Rate,Taxable Value,Tax Amount"
    'fsoStream.WriteLine strLine

    from_dt = Split(Me.txtfrom.Text, "/")
    to_dt = Split(Me.txtto.Text, "/")
    'Set rec1 = db.OpenRecordset("select * from invoicehead where InvNo=" & Val(Me.DBGrid1.Columns(0)) & " and InvType='" & Me.CboInvType.Text & "'")
    Set rec1 = db.OpenRecordset("select * from invoicehead inner join Ledgermaster on InvoiceHead.AccId=LedgerMaster.AccId where InvoiceHead.InvDate between #" & from_dt(1) & "/" & from_dt(0) & "/" & from_dt(2) & "# and #" & to_dt(1) & "/" & to_dt(0) & "/" & to_dt(2) & "#")
    RowCount = 2
    While Not rec1.EOF
        Set rec2 = db.OpenRecordset("select * from PartyDr where Accid=" & rec1("Invoicehead.Accid"))
        If Not rec2.EOF Then
            If rec2("Tin") <> "" Then
                If Not rec2.EOF Then
                    excelWS.Cells(RowCount, 1).Value = rec2("Tin")
                    strLine = IIf(IsNull(rec2("tin")), " ", rec2("tin")) & ","
                    Set rec2 = db.OpenRecordset("select * from statecode where Stcode=" & rec2("statecode"))
                    excelWS.Cells(RowCount, 7).Value = rec2("stcode") & "-" & rec2("statename")
                    temp_state = rec2("stcode") & "-" & rec2("statename")
                End If
                excelWS.Cells(RowCount, 2).Value = rec1("AccName")
                strLine = strLine & rec1("AccName") & ","
                excelWS.Cells(RowCount, 3).Value = rec1("Address1")
                strLine = strLine & rec1("Address1") & ","
                excelWS.Cells(RowCount, 4).Value = rec1("InvNO")
                strLine = strLine & rec1("InvNO") & ","
                excelWS.Cells(RowCount, 5).Value = "'" & str(rec1("InvDate"))
                strLine = strLine & str(rec1("InvDate")) & ","
                excelWS.Cells(RowCount, 6).Value = rec1("Grandtotal")
                strLine = strLine & rec1("Grandtotal") & ","
                strLine = strLine & temp_state & ","
                excelWS.Cells(RowCount, 8).Value = "0"
                strLine = strLine & "0,"
                excelWS.Cells(RowCount, 9).Value = "Tax Invoice"
                strLine = strLine & "Tax Invoice,"
                Set rec2 = db.OpenRecordset("select vat as TRate,sum(gross-(discountamount)) as Taxable,sum(vatamount) as taxamount from invoicedetails where Invno=" & rec1("InvNO") & " Group by vat")
                R = 0
                While Not rec2.EOF
                    strLine = strLine & ","
                    excelWS.Cells(RowCount, 11).Value = rec2!TRate
                    strLine = strLine & rec2!TRate & ","
                    'excelWS.Cells(RowCount, 12).Value = rec2!Taxable - rec2!TAXAMOUNT
                    excelWS.Cells(RowCount, 12).Value = rec2!Taxable
                    strLine = strLine & rec2!Taxable & ","
                    excelWS.Cells(RowCount, 13).Value = rec2!TAXAMOUNT
                    strLine = strLine & rec2!TAXAMOUNT & ","
                    If R = 0 Then
                        strLine = strLine
                    Else
                        strLine = strLine
                    End If
                    'fsoStream.WriteLine strLine
                    strLine = ""
                    RowCount = RowCount + 1
                    rec2.MoveNext
                    R = R + 1
                Wend
                RowCount = RowCount + 1

            End If
        End If
        rec1.MoveNext
    Wend
    'fsoStream.Close
    'Set fsoStream = Nothing
    'Set fso = Nothing

    excelApp.Visible = True
    Me.MousePointer = 0
    'excelApp.Workbooks.Close

    Set ExcelSheet = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdcreditnote_Click()
Dim strLine As String

    Me.MousePointer = vbHourglass
    Dim from_dt() As String
    Dim to_dt() As String
    Dim tempfilename
    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelWS.Cells(1, 1).Value = "GSTIN/UIN of Recipient"
    excelWS.Cells(1, 2).Value = "Recipient"
    excelWS.Cells(1, 3).Value = "Address"
    excelWS.Cells(1, 4).Value = "Note No"
    excelWS.Cells(1, 5).Value = "Note Date"
    excelWS.Cells(1, 6).Value = "Note Type"
    excelWS.Cells(1, 7).Value = "Place Of Supply"
    excelWS.Cells(1, 8).Value = "Reverse Charge"
    excelWS.Cells(1, 9).Value = "Note Supply Type"
    excelWS.Cells(1, 10).Value = "Note Value"
    excelWS.Cells(1, 11).Value = "Applicable % Tax"
    excelWS.Cells(1, 12).Value = "Taxable Value"
    excelWS.Cells(1, 13).Value = "Tax Amount"

    from_dt = Split(Me.txtfrom.Text, "/")
    to_dt = Split(Me.txtto.Text, "/")
    'Set rec1 = db.OpenRecordset("select * from invoicehead where InvNo=" & Val(Me.DBGrid1.Columns(0)) & " and InvType='" & Me.CboInvType.Text & "'")
    Set rec1 = db.OpenRecordset("select * from salesretunhead inner join Ledgermaster on salesreturnHead.AccId=LedgerMaster.AccId where salesreturnHead.InvDate between #" & from_dt(1) & "/" & from_dt(0) & "/" & from_dt(2) & "# and #" & to_dt(1) & "/" & to_dt(0) & "/" & to_dt(2) & "#")
    RowCount = 2
    While Not rec1.EOF
        Set rec2 = db.OpenRecordset("select * from PartyDr where Accid=" & rec1("Invoicehead.Accid"))
        If Not rec2.EOF Then
            If rec2("Tin") <> "" Then
                If Not rec2.EOF Then
                    excelWS.Cells(RowCount, 1).Value = rec2("Tin")
                    Set rec2 = db.OpenRecordset("select * from statecode where Stcode=" & rec2("statecode"))
                    excelWS.Cells(RowCount, 7).Value = rec2("stcode") & "-" & rec2("statename")
                    temp_state = rec2("stcode") & "-" & rec2("statename")
                End If
                excelWS.Cells(RowCount, 2).Value = rec1("AccName")
                excelWS.Cells(RowCount, 3).Value = rec1("Address1")
                excelWS.Cells(RowCount, 4).Value = rec1("InvNO")
                excelWS.Cells(RowCount, 5).Value = "'" & str(rec1("InvDate"))
                excelWS.Cells(RowCount, 6).Value = "C"
                excelWS.Cells(RowCount, 8).Value = "0"
                excelWS.Cells(RowCount, 9).Value = "Regular B2B"
                excelWS.Cells(RowCount, 10).Value = rec1("Grandtotal")
                Set rec2 = db.OpenRecordset("select vat as TRate,sum(gross-(discountamount)) as Taxable,sum(vatamount) as taxamount from salesreturndetails where Invno=" & rec1("InvNO") & " Group by vat")
                R = 0
                While Not rec2.EOF
                    excelWS.Cells(RowCount, 11).Value = rec2!TRate
                    'excelWS.Cells(RowCount, 12).Value = rec2!Taxable - rec2!TAXAMOUNT
                    excelWS.Cells(RowCount, 12).Value = rec2!Taxable
                    excelWS.Cells(RowCount, 13).Value = rec2!TAXAMOUNT
                    'fsoStream.WriteLine strLine
                    RowCount = RowCount + 1
                    rec2.MoveNext
                    R = R + 1
                Wend
                RowCount = RowCount + 1

            End If
        End If
        rec1.MoveNext
    Wend
    'fsoStream.Close
    'Set fsoStream = Nothing
    'Set fso = Nothing

    excelApp.Visible = True
    Me.MousePointer = 0
    'excelApp.Workbooks.Close

    Set ExcelSheet = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing
End Sub

Private Sub cmdpurchase_Click()
Me.MousePointer = vbHourglass
Dim from_dt() As String
Dim to_dt() As String
Set excelApp = CreateObject("Excel.Application")
Set excelWB = excelApp.Workbooks.Add
Set excelWS = excelWB.Worksheets(1)
excelWS.Cells(1, 1).Value = "GSTIN/UIN of Supplier"
excelWS.Cells(1, 2).Value = "Supplier Name"
excelWS.Cells(1, 3).Value = "Address"
excelWS.Cells(1, 4).Value = "Invoice No"
excelWS.Cells(1, 5).Value = "Invoice Date"
excelWS.Cells(1, 6).Value = "Invoice Value"
excelWS.Cells(1, 7).Value = "Place Of Supply"
excelWS.Cells(1, 8).Value = "Reverse Charge"
excelWS.Cells(1, 9).Value = "Invoice Type"
excelWS.Cells(1, 10).Value = "E-Commerce GSTIN"
excelWS.Cells(1, 11).Value = "Rate"
excelWS.Cells(1, 12).Value = "Taxable Value"
excelWS.Cells(1, 13).Value = "Tax Amount"


from_dt = Split(Me.txtfrom.Text, "/")
to_dt = Split(Me.txtto.Text, "/")
'Set rec1 = db.OpenRecordset("select * from invoicehead where InvNo=" & Val(Me.DBGrid1.Columns(0)) & " and InvType='" & Me.CboInvType.Text & "'")
Set rec1 = db.OpenRecordset("select * from purchasehead inner join Ledgermaster on PurchaseHead.AccId=LedgerMaster.AccId where PurchaseHead.PurchaseDate between #" & from_dt(1) & "/" & from_dt(0) & "/" & from_dt(2) & "# and #" & to_dt(1) & "/" & to_dt(0) & "/" & to_dt(2) & "#")
RowCount = 2
While Not rec1.EOF
    'excelWS.Cells(RowCount, 1).Value = REC1("GSTN") & "/" & REC1("ARN")
    Set rec2 = db.OpenRecordset("select * from PartyCr where Accid=" & rec1("PurchaseHead.Accid"))
    If Not rec2.EOF Then
        excelWS.Cells(RowCount, 1).Value = rec2("Tin")
        Set rec2 = db.OpenRecordset("select * from statecode where Stcode=" & rec2("statecode"))
        excelWS.Cells(RowCount, 7).Value = rec2("stcode") & "-" & rec2("statename")
    End If
    excelWS.Cells(RowCount, 2).Value = rec1("AccName")
    excelWS.Cells(RowCount, 3).Value = rec1("Address1")
    excelWS.Cells(RowCount, 4).Value = rec1("InvNO")
    excelWS.Cells(RowCount, 5).Value = "'" & str(rec1("InvDate"))
    excelWS.Cells(RowCount, 6).Value = rec1("Grandtotal")
    'Set rec2 = db.OpenRecordset("select * from statecode where Stcode=" & REC1("statecode"))
    'excelWS.Cells(RowCount, 7).Value = rec2("stcode") & "-" & rec2("statename")
    excelWS.Cells(RowCount, 8).Value = "0"
    excelWS.Cells(RowCount, 9).Value = "Tax Invoice"
    Set rec2 = db.OpenRecordset("select vat as TRate,sum((amount-discount_amount)) as Taxable,sum(vatamount) as taxamt from purchasedetails where slno=" & rec1("slno") & " Group by vat")
    While Not rec2.EOF
        excelWS.Cells(RowCount, 11).Value = rec2!TRate
        excelWS.Cells(RowCount, 12).Value = rec2!Taxable
        excelWS.Cells(RowCount, 13).Value = rec2!Taxamt
        RowCount = RowCount + 1
        rec2.MoveNext
    Wend
    RowCount = RowCount + 1
    rec1.MoveNext
Wend
excelApp.Visible = True
Me.MousePointer = 0
'excelApp.Workbooks.Close

    Set ExcelSheet = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing

End Sub

Private Sub Command1_Click()
    Dim strLine As String
    'Dim fso As New FileSystemObject
    'Dim fsoStream As TextStream

    Me.MousePointer = vbHourglass
    Dim from_dt() As String
    Dim to_dt() As String
    'Dim tempfilename
    'Me.CommonDialog1.Filter = "Apps (*.csv)"
    'Me.CommonDialog1.DefaultExt = "csv"
    'Me.CommonDialog1.DialogTitle = "Save CSV file"
    'Me.CommonDialog1.InitDir = App.Path & "\inputfile"
    'Me.CommonDialog1.ShowSave
    'tempfilename = Me.CommonDialog1.FileName
    'Set fsoStream = fso.CreateTextFile(tempfilename, True)
    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelWS.Cells(1, 1).Value = "GSTIN/UIN of Recipient"
    excelWS.Cells(1, 2).Value = "Recipient"
    excelWS.Cells(1, 3).Value = "Address"
    excelWS.Cells(1, 4).Value = "InvoiceNo"
    excelWS.Cells(1, 5).Value = "Invoice Date"
    excelWS.Cells(1, 6).Value = "Invoice Value"
    excelWS.Cells(1, 7).Value = "Place Of Supply"
    excelWS.Cells(1, 8).Value = "Reverse Charge"
    excelWS.Cells(1, 9).Value = "Invoice Type"
    excelWS.Cells(1, 10).Value = "E-Commerce GSTIN"
    excelWS.Cells(1, 11).Value = "Rate"
    excelWS.Cells(1, 12).Value = "Taxable Value"
    excelWS.Cells(1, 13).Value = "Tax Amount"
    strLine = "GSTIN/UIN of Recipient,Recipent,Address,Invoice No,Invoice Date,Invoice Value,Place Of Supply,Reverse Charge,Invoice type,E-CommerceGSTIN,Rate,Taxable Value,Tax Amount"
    'fsoStream.WriteLine strLine

    from_dt = Split(Me.txtfrom.Text, "/")
    to_dt = Split(Me.txtto.Text, "/")
    'Set rec1 = db.OpenRecordset("select * from invoicehead where InvNo=" & Val(Me.DBGrid1.Columns(0)) & " and InvType='" & Me.CboInvType.Text & "'")
    Set rec1 = db.OpenRecordset("select * from invoicehead inner join Ledgermaster on InvoiceHead.AccId=LedgerMaster.AccId where InvoiceHead.InvDate between #" & from_dt(1) & "/" & from_dt(0) & "/" & from_dt(2) & "# and #" & to_dt(1) & "/" & to_dt(0) & "/" & to_dt(2) & "#")
    RowCount = 2
    While Not rec1.EOF
        'Set rec2 = db.OpenRecordset("select * from PartyDr where Accid=" & rec1("Invoicehead.Accid"))
        'If Not rec2.EOF Then
            If rec1("Tin") = "" Or IsNull(rec1("tin")) Then
                'If Not rec2.EOF Then
                    'excelWS.Cells(RowCount, 1).Value = rec2("Tin")
                    'strLine = IIf(IsNull(rec2("tin")), " ", rec2("tin")) & ","
                    'Set rec2 = db.OpenRecordset("select * from statecode where Stcode=" & rec1("statecode"))
                    'excelWS.Cells(RowCount, 7).Value = rec2("stcode") & "-" & rec2("statename")
                    'temp_state = rec2("stcode") & "-" & rec2("statename")
                'End If
                excelWS.Cells(RowCount, 2).Value = rec1("AccName")
                strLine = strLine & rec1("AccName") & ","
                excelWS.Cells(RowCount, 3).Value = rec1("Address1")
                strLine = strLine & rec1("Address1") & ","
                excelWS.Cells(RowCount, 4).Value = rec1("InvNO")
                strLine = strLine & rec1("InvNO") & ","
                excelWS.Cells(RowCount, 5).Value = "'" & str(rec1("InvDate"))
                strLine = strLine & str(rec1("InvDate")) & ","
                excelWS.Cells(RowCount, 6).Value = rec1("Grandtotal")
                strLine = strLine & rec1("Grandtotal") & ","
                strLine = strLine & temp_state & ","
                excelWS.Cells(RowCount, 8).Value = "0"
                strLine = strLine & "0,"
                excelWS.Cells(RowCount, 9).Value = "Tax Invoice"
                strLine = strLine & "Tax Invoice,"
                Set rec2 = db.OpenRecordset("select vat as TRate,sum(gross-(discountamount)) as Taxable,sum(vatamount) as taxamount from invoicedetails where Invno=" & rec1("InvNO") & " Group by vat")
                R = 0
                While Not rec2.EOF
                    strLine = strLine & ","
                    excelWS.Cells(RowCount, 11).Value = rec2!TRate
                    strLine = strLine & rec2!TRate & ","
                    'excelWS.Cells(RowCount, 12).Value = rec2!Taxable - rec2!TAXAMOUNT
                    excelWS.Cells(RowCount, 12).Value = rec2!Taxable
                    strLine = strLine & rec2!Taxable & ","
                    excelWS.Cells(RowCount, 13).Value = rec2!TAXAMOUNT
                    strLine = strLine & rec2!TAXAMOUNT & ","
                    If R = 0 Then
                        strLine = strLine
                    Else
                        strLine = strLine
                    End If
                    'fsoStream.WriteLine strLine
                    strLine = ""
                    RowCount = RowCount + 1
                    rec2.MoveNext
                    R = R + 1
                Wend
                RowCount = RowCount + 1

            End If
       ' End If
        rec1.MoveNext
    Wend
    'fsoStream.Close
    'Set fsoStream = Nothing
    'Set fso = Nothing

    excelApp.Visible = True
    Me.MousePointer = 0
    'excelApp.Workbooks.Close

    Set ExcelSheet = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing
End Sub

