VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmnewstockst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Statement"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdprint 
      Caption         =   "Generate Report"
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
      Left            =   2183
      TabIndex        =   6
      Top             =   1620
      Width           =   1725
   End
   Begin VB.ComboBox cbocategory 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1568
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   930
      Width           =   3915
   End
   Begin MSMask.MaskEdBox txtfrom 
      Height          =   375
      Left            =   1238
      TabIndex        =   0
      Top             =   270
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
      Left            =   3998
      TabIndex        =   1
      Top             =   270
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
   Begin VB.Label Label3 
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   608
      TabIndex        =   4
      Top             =   960
      Width           =   1215
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
      Left            =   3645
      TabIndex        =   3
      Top             =   300
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
      Left            =   630
      TabIndex        =   2
      Top             =   300
      Width           =   735
   End
End
Attribute VB_Name = "frmnewstockst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec1 As Recordset

Private Sub cmdprint_Click()
    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)

    excelWS.Cells(1, 1).Value = "STOCK REPORT FROM " & Me.txtfrom.Text & " TO " & Me.txtto.Text
    excelWS.Cells(2, 2).Value = "PRODUCT NAME"
    excelWS.Cells(2, 3).Value = "OPENING"
    excelWS.Cells(2, 4).Value = "STOCK IN"
    excelWS.Cells(2, 5).Value = "STOCK OUT"
    excelWS.Cells(2, 6).Value = "CLOSING"

    Set rec1 = db.OpenRecordset("select min(purchasedate) as min_pdate from purchasehead")
    If Not IsNull(rec1!min_pdate) Then
        minpdate = rec1!min_pdate
    End If
    Set rec1 = db.OpenRecordset("select min(invdate) as min_cdate from invoicehead")
    If Not IsNull(rec1!min_cdate) Then
        mincdate = rec1!min_cdate
    End If

    from_dt = Split(Me.txtfrom.Text, "/")
    to_dt = Split(Me.txtto.Text, "/")
    
    Set rec1 = db.OpenRecordset("SELECt * from itemmaster where producttype='" & Me.cbocategory.Text & "'")
    'Set rec1 = db.OpenRecordset("select * from itemmaster where productcode=1059")
    RowCount = 3
    While Not rec1.EOF
        tempopstock = rec1("openingstock")
        temppcode = rec1("productcode")
        Set rec = db.OpenRecordset("select sum(purchasedetails.qty) as tpqty from purchasehead inner join purchasedetails on purchasehead.slno=purchasedetails.slno where purchasedetails.productcode=" & temppcode & " and purchasehead.purchasedate <#" & Format(Me.txtfrom.Text, "mm/dd/yyyy") & "#")
        If Not IsNull(rec!tpqty) Then
            temppqty = rec!tpqty
        Else
            temppqty = 0
        End If
        Set rec = db.OpenRecordset("select sum(outwardchallandetails.qty) as tpqty from outwardchallanhead inner join outwardchallandetails on outwardchallanhead.challanno=outwardchallandetails.challanno where outwardchallandetails.productcode=" & temppcode & " and outwardchallanhead.challandaate < #" & Format(Me.txtfrom.Text, "mm/dd/yyyy") & "#")
        If Not IsNull(rec!tpqty) Then
            temcpqty = rec!tpqty
        Else
            tempcqty = 0
        End If
        Set rec = db.OpenRecordset("select sum(invoicedetails.qty) as tpqty from invoicehead inner join invoicedetails on invoicehead.invno=invoicedetails.invno where invoicedetails.productcode=" & temppcode & " and invoicehead.invdate < #" & Format(Me.txtfrom.Text, "mm/dd/yyyy") & "#")
        If Not IsNull(rec!tpqty) Then
            tempsqty = rec!tpqty
        Else
            tempsqty = 0
        End If
        openingqty = (tempopstock + temppqty) - (tempcqty + tempsqty)
        '---------------------------------------------------------------
        Set rec = db.OpenRecordset("select sum(purchasedetails.qty) as tpqty from purchasehead inner join purchasedetails on purchasehead.slno=purchasedetails.slno where purchasedetails.productcode=" & temppcode & " and purchasehead.purchasedate between #" & Format(Me.txtfrom.Text, "mm/dd/yyyy") & "# and #" & Format(Me.txtto.Text, "mm/dd/yyyy") & "#")
        If Not IsNull(rec!tpqty) Then
            temppurchaseqty = rec!tpqty
        Else
            temppurchaseqty = 0
        End If
        Set rec = db.OpenRecordset("select sum(outwardchallandetails.qty) as tpqty from outwardchallanhead inner join outwardchallandetails on outwardchallanhead.challanno=outwardchallandetails.challanno where outwardchallandetails.productcode=" & temppcode & " and outwardchallanhead.challandaate between #" & Format(Me.txtfrom.Text, "mm/dd/yyyy") & "# and #" & Format(Me.txtto.Text, "mm/dd/yyyy") & "#")
        If Not IsNull(rec!tpqty) Then
            tempoutqty = rec!tpqty
        Else
            tempoutqty = 0
        End If
        Set rec = db.OpenRecordset("select sum(invoicedetails.qty) as tpqty from invoicehead inner join invoicedetails on invoicehead.invno=invoicedetails.invno where invoicedetails.productcode=" & temppcode & " and invoicehead.invdate between #" & Format(Me.txtfrom.Text, "mm/dd/yyyy") & "# and #" & Format(Me.txtto.Text, "mm/dd/yyyy") & "#")
        If Not IsNull(rec!tpqty) Then
            tempsaleqty = rec!tpqty
        Else
            tempsaleqty = 0
        End If
        excelWS.Cells(RowCount, 2).Value = rec1("item")
        excelWS.Cells(RowCount, 3).Value = openingqty
        excelWS.Cells(RowCount, 4).Value = temppurchaseqty
        excelWS.Cells(RowCount, 5).Value = tempoutqty + tempsaleqty
        excelWS.Cells(RowCount, 6).Value = (openingqty + temppurchaseqty) - (tempoutqty + tempsaleqty)
        rec1.MoveNext
        RowCount = RowCount + 1
    Wend


    excelApp.Visible = True
    Me.MousePointer = 0
    'excelApp.Workbooks.Close

    Set ExcelSheet = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing
End Sub

Private Sub Form_Load()
Me.txtfrom.Text = Format(Date, "dd/mm/yyyy")
Me.txtto.Text = Format(Date, "dd/mm/yyyy")
Set rec1 = db.OpenRecordset("select distinct producttype from itemmaster")
While Not rec1.EOF
    Me.cbocategory.AddItem rec1("producttype")
    rec1.MoveNext
Wend
If Me.cbocategory.ListCount > 0 Then
    Me.cbocategory.ListIndex = 0
End If
End Sub
