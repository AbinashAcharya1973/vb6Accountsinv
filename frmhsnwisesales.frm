VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmhsnwisesales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HSN-Wise Sales"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Generate Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   975
      Width           =   1665
   End
   Begin MSMask.MaskEdBox txtto 
      Height          =   375
      Left            =   3150
      TabIndex        =   1
      Top             =   225
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
   Begin MSMask.MaskEdBox txtfrom 
      Height          =   375
      Left            =   630
      TabIndex        =   2
      Top             =   225
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
      Left            =   75
      TabIndex        =   4
      Top             =   225
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
      Left            =   2790
      TabIndex        =   3
      Top             =   225
      Width           =   375
   End
End
Attribute VB_Name = "frmhsnwisesales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSave_Click()
Me.MousePointer = vbHourglass
Dim from_dt() As String
Dim to_dt() As String
Set excelApp = CreateObject("Excel.Application")
Set excelWB = excelApp.Workbooks.Add
Set excelWS = excelWB.Worksheets(1)
excelWS.Cells(1, 1).Value = "GSTIN/UIN of Recipient"
excelWS.Cells(1, 2).Value = "Recipient"
excelWS.Cells(1, 3).Value = "Address"
excelWS.Cells(1, 4).Value = "InvoiceNo"
excelWS.Cells(1, 5).Value = "Invoice Date"
excelWS.Cells(1, 6).Value = "Place Of Supply"
excelWS.Cells(1, 7).Value = "HSN"
excelWS.Cells(1, 8).Value = "Taxable Amount"
excelWS.Cells(1, 9).Value = "Tax Amount"

from_dt = Split(Me.txtfrom.Text, "/")
to_dt = Split(Me.txtto.Text, "/")
'Set rec1 = db.OpenRecordset("select * from invoicehead where InvNo=" & Val(Me.DBGrid1.Columns(0)) & " and InvType='" & Me.CboInvType.Text & "'")
''Set REC1 = db.OpenRecordset("select * from invoicehead inner join Ledgermaster on InvoiceHead.AccId=LedgerMaster.AccId where InvoiceHead.InvDate between #" & from_dt(1) & "/" & from_dt(0) & "/" & from_dt(2) & "# and #" & to_dt(1) & "/" & to_dt(0) & "/" & to_dt(2) & "#")

Set rec1 = db.OpenRecordset("select * from invoicehead where InvoiceHead.InvDate between #" & from_dt(1) & "/" & from_dt(0) & "/" & from_dt(2) & "# and #" & to_dt(1) & "/" & to_dt(0) & "/" & to_dt(2) & "#")
RowCount = 2
While Not rec1.EOF
    Set rec2 = db.OpenRecordset("select * from invoicehead inner join ledgermaster on invoicehead.accid=ledgermaster.accid where invoicehead.invno=" & rec1("invno"))
    If Not rec2.EOF Then
    excelWS.Cells(RowCount, 1).Value = rec2("tin")
    excelWS.Cells(RowCount, 2).Value = rec2("accname")
    excelWS.Cells(RowCount, 3).Value = rec2("Address1")
    excelWS.Cells(RowCount, 4).Value = rec2("InvNO")
    excelWS.Cells(RowCount, 5).Value = "'" & str(rec2("InvDate"))
    
    Set rec2 = db.OpenRecordset("select * from statecode where Stcode=" & rec2("statecode"))
    excelWS.Cells(RowCount, 6).Value = rec2("stcode") & "-" & rec2("statename")
    End If
    Set rec2 = db.OpenRecordset("select invoicedetails.invno,itemmaster.hsn,sum(invoicedetails.qty) as tqty,sum(gross-discountamount) as taxable,sum(invoicedetails.vatamount) as taxamount from (invoicedetails inner join itemmaster on invoicedetails.productcode=itemmaster.productcode) inner join invoicehead on invoicedetails.invno=InvoiceHead.invno where InvoiceHead.Invno=" & rec1("invno") & " and invoicehead.invtype='" & rec1("invtype") & "' group by invoicedetails.invno,itemmaster.hsn")
    While Not rec2.EOF
        excelWS.Cells(RowCount, 7).Value = rec2("hsn")
        excelWS.Cells(RowCount, 8).Value = rec2("taxable")
        excelWS.Cells(RowCount, 9).Value = rec2("taxamount")
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

