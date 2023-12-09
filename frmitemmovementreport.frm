VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmitemmovementreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Movement Report"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   12630
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7560
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
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
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtto 
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSMask.MaskEdBox txtfrom 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmitemmovementreport.frx":0000
      Height          =   5535
      Left            =   120
      OleObjectBlob   =   "frmitemmovementreport.frx":0014
      TabIndex        =   4
      Top             =   1560
      Width           =   12375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ItemWiseStockStatement"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboitemname 
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
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.ComboBox cboitemtype 
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
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Left            =   3480
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Category"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmitemmovementreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432

Private Sub cboitemtype_Change()
    Set rec2 = db.OpenRecordset("select * from ItemMaster where CategoryId=" & Me.cboitemtype.ItemData(Me.cboitemtype.ListIndex))
    Me.cboitemname.Clear
    If Not rec2.EOF Then
        While Not rec2.EOF
            Me.cboitemname.AddItem (rec2("Item"))
            Me.cboitemname.ItemData(Me.cboitemname.NewIndex) = rec2("ProductCode")
            rec2.MoveNext
        Wend
        If Me.cboitemname.ListCount > 0 Then
            Me.cboitemname.ListIndex = 0
        End If
    End If
End Sub

Private Sub cboitemtype_Click()
    cboitemtype_Change
End Sub

Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    Me.CrystalReport1.ReportFileName = App.Path & "\ItemMovement.rpt"
    Me.CrystalReport1.PrinterName = Printer.DeviceName
    Me.CrystalReport1.PrinterDriver = Printer.DriverName
    Me.CrystalReport1.PrinterPort = Printer.Port
    Me.CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Me.Data1.databasename = db.Name
    Me.txtfrom.Text = Format(Date, "dd/mm/yyyy")
    Me.txtto.Text = Format(Date, "dd/mm/yyyy")

    db.Execute ("Delete * from ItemWiseStockStatement")
    Me.Data1.Refresh
    
        Set rec1 = db.OpenRecordset("select * from Product")
        While Not rec1.EOF
            Me.cboitemtype.AddItem (rec1("Productname"))
            Me.cboitemtype.ItemData(Me.cboitemtype.NewIndex) = rec1("Pid")
            rec1.MoveNext
        Wend
        If Me.cboitemtype.ListCount > 0 Then
            Me.cboitemtype.ListIndex = 0
        End If
        Exit Sub
errtrap:
        MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtFrom_GotFocus()
    Me.txtfrom.SelStart = 0
    Me.txtfrom.SelLength = Len(Me.txtfrom.Text)
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtto.SetFocus
    End If
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        fromdate = Mid(Me.txtfrom.Text, 4, 2) & "/" & Mid(Me.txtfrom.Text, 1, 2) & "/" & Mid(Me.txtfrom.Text, 7, 4)
        todate = Mid(Me.txtto.Text, 4, 2) & "/" & Mid(Me.txtto.Text, 1, 2) & "/" & Mid(Me.txtto.Text, 7, 4)

        db.Execute ("Delete * from ItemWiseStockStatement")

        Set rec1 = db.OpenRecordset("select * From ItemMaster where ProductCode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
        If Not rec1.EOF Then
            OpeningStock = rec1("Openingstock")
            Closing_qty = rec1("Openingstock")
            Purchaserate = rec1("Purchaserate")
            db.Execute ("Insert into ItemWiseStockStatement (PartiCulars,ClosingQty,ClosingValue) values('Opening Stock'," & OpeningStock & "," & OpeningStock * rec1("Purchaserate") & ")")
        End If
        '    'Find opening stock as from date
        '    Set rec1 = db.OpenRecordset("select InvoiceHead.InvDate,InvoiceDetails.ProductCode,InvoiceDetails.Qty,InvoiceDetails.Free_Qty from invoicehead inner join InvoiceDetails on InvoiceHead.InvNo=InvoiceDetails.InvNo where InvoiceHead.InvDate  between #" & fromdate & "# and #" & todate & "# and InvoiceDetails.ProductCode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
        '    While Not rec1.EOF
        '               Op_Inward_qty = Op_Inward_qty + (rec1("qty") + rec1("free_qty"))
        '        rec1.MoveNext
        '    Wend

        'Purchase Transaction between Date range
        Set rec1 = db.OpenRecordset("select PurchaseHead.Slno,PurchaseHead.Purchasedate,PurchaseDetails.Qty,PurchaseHead.Supplier,PurchaseDetails.Free_Qty,PurchaseDetails.Amount,PurchaseDetails.ProductCode from PurchaseHead inner join PurchaseDetails on PurchaseHead.SlNo=PurchaseDetails.SlNo where Purchasehead.Purchasedate between #" & fromdate & "# and #" & todate & "# and PurchaseDetails.ProductCode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
        While Not rec1.EOF
            db.Execute ("Insert into ItemWiseStockStatement (Tdate,PartiCulars,VchType,VchNo,InwardQty,InwardValue) values('" & rec1("Purchasedate") & "','" & rec1("Supplier") & "','Purchase'," & rec1("Slno") & "," & rec1("Qty") + rec1("Free_qty") & "," & rec1("Amount") & ")")
            rec1.MoveNext
        Wend
        'Invoice Transaction between Date range
        Set rec1 = db.OpenRecordset("select InvoiceHead.InvNo,InvoiceHead.InvDate,InvoiceHead.InvType,InvoiceHead.Party,InvoiceDetails.ProductCode,InvoiceDetails.Qty,InvoiceDetails.Free_Qty,InvoiceDetails.Gross from invoicehead inner join InvoiceDetails on InvoiceHead.InvNo=InvoiceDetails.InvNo where InvoiceHead.InvDate  between #" & fromdate & "# and #" & todate & "# and InvoiceDetails.ProductCode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
        While Not rec1.EOF
            db.Execute ("Insert into ItemWiseStockStatement (Tdate,PartiCulars,VchType,VchNo,OutwardQty,OutwardValue) values('" & rec1("InvDate") & "','" & rec1("Party") & "','" & rec1("InvType") & "'," & rec1("InvNo") & "," & rec1("Qty") + rec1("Free_qty") & "," & rec1("Gross") & ")")
            rec1.MoveNext
        Wend
        'Sales Return between Date Range
        Set rec1 = db.OpenRecordset("select Salesreturnhead.InvNo,Salesreturnhead.InvDate,Salesreturnhead.Party,Salesreturndetails.ProductCode,Salesreturndetails.Qty,Salesreturndetails.Free_Qty,Salesreturndetails.Gross from Salesreturnhead inner join Salesreturndetails on Salesreturnhead.InvNo=Salesreturndetails.InvNo where Salesreturnhead.InvDate  between #" & fromdate & "# and #" & todate & "# and Salesreturndetails.ProductCode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
        While Not rec1.EOF
            db.Execute ("Insert into ItemWiseStockStatement (Tdate,PartiCulars,VchType,VchNo,InwardQty,InwardValue) values('" & rec1("InvDate") & "','" & rec1("Party") & "','Sales Return'," & rec1("InvNo") & "," & rec1("Qty") + rec1("Free_qty") & "," & rec1("Gross") & ")")
            rec1.MoveNext
        Wend
        'Purchase Return between date Range
        Set rec1 = db.OpenRecordset("select PurchaseReturnHead.Slno,PurchaseReturnHead.Purchasedate,PurchaseReturnHead.Supplier,PurchaseReturnDetails.ProductCode,PurchaseReturnDetails.Qty,PurchaseReturnDetails.Free_Qty,PurchaseReturnDetails.Amount from PurchaseReturnHead inner join PurchaseReturnDetails on PurchaseReturnHead.Slno=PurchaseReturnDetails.slno where PurchaseReturnHead.Purchasedate  between #" & fromdate & "# and #" & todate & "# and PurchaseReturnDetails.ProductCode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
        While Not rec1.EOF
            db.Execute ("Insert into ItemWiseStockStatement (Tdate,PartiCulars,VchType,VchNo,OutwardQty,OutwardValue) values('" & rec1("Purchasedate") & "','" & rec1("Supplier") & "','Purchase Return'," & rec1("Slno") & "," & rec1("Qty") + rec1("Free_qty") & "," & rec1("Amount") & ")")
            rec1.MoveNext
        Wend

        'Calculate closing balance & sort date wise
        Set rec1 = db.OpenRecordset("Select * from ItemWiseStockStatement order by Tdate")
        While Not rec1.EOF
            Closing_qty = Closing_qty + rec1("InwardQty") - rec1("OutwardQty")
            Closing_value = Format(Closing_qty * Purchaserate, "######0.00")
            db.Execute ("Update ItemWiseStockStatement set ClosingQty=" & Closing_qty & ",ClosingValue=" & Closing_value & " where VchType='" & rec1("VchType") & "' and VchNo=" & rec1("VchNo"))
            rec1.MoveNext
        Wend
        db.Execute ("Update ItemWiseStockStatement set fromdt='" & Me.txtfrom.Text & "',todt='" & Me.txtto.Text & "',ItemName='" & Me.cboitemname.Text & "'")
        Me.Data1.RecordSource = "Select * from ItemWiseStockStatement order by tdate"
        Me.Data1.Refresh
    End If
End Sub
