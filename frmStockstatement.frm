VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmStockstatement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Statement"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5655
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cboitemtype 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cboitemname 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   3855
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4200
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.TextBox txtfrom_dt_d 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtfrom_dt_m 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtfrom_dt_y 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtto_dt_d 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtto_dt_m 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtto_dt_y 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label9 
      Caption         =   "Item Type"
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Item Name"
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
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1335
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
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   855
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
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "frmStockstatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As DAO.Recordset, REC2 As DAO.Recordset
Attribute REC2.VB_VarUserMemId = 1073938432

Private Sub cboitemtype_Change()
    Set REC1 = db.OpenRecordset("select distinct Item from ItemMaster where ItemType='" & Me.cboitemtype.Text & "'")
    Me.cboitemname.Clear
    If Not REC1.EOF Then
        While Not REC1.EOF
            Me.cboitemname.AddItem (REC1("Item"))
            REC1.MoveNext
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
    db.Execute ("delete * from StockStatement")
    db.Execute ("insert into StockStatement(ItemType,Items,Size,PurchaseRate,Qty) select itemtype,itemname,size,prate,Qty from stock")

    '-------------------------Finding Issue Qty FromDate--------------------
    Set REC1 = db.OpenRecordset("select InvoiceHead.InvDate,InvoiceDetails.ItemType,InvoiceDetails.Itemname,InvoiceDetails.size,InvoiceDetails.PrRate,InvoiceDetails.Qty from invoicehead inner join InvoiceDetails on InvoiceHead.InvNo=InvoiceDetails.InvNo where InvoiceHead.InvDate  between #" & Me.txtfrom_dt_m.Text & "/" & Me.txtfrom_dt_d.Text & "/" & Me.txtfrom_dt_y.Text & "# and #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement Set FromDtSales=FromDtSales + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and items='" & REC1("ItemName") & "' and size='" & REC1("Size") & "' and PurchaseRate=" & REC1("PrRate"))
        Debug.Print REC1("invdate")
        REC1.MoveNext
    Wend
    '--------------------------Finding Stockin FromDate-----------------------
    Set REC1 = db.OpenRecordset("select PurchaseHead.Purchasedate,PurchaseDetails.ItemType,PurchaseDetails.ItemName,PurchaseDetails.Size,PurchaseDetails.Qty,PurchaseDetails.PrRate from PurchaseHead inner join PurchaseDetails on PurchaseHead.SlNo=PurchaseDetails.SlNo where Purchasehead.Purchasedate between #" & Me.txtfrom_dt_m.Text & "/" & Me.txtfrom_dt_d.Text & "/" & Me.txtfrom_dt_y.Text & "# and #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set FromDtPurchase=FromDtPurchase + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    '---------------------------Finding IssueReturn FromDate-------------------------
    'Set Rec1 = Db.OpenRecordset("SELECT * FROM SALESRETURNDETAILS")
    Set REC1 = db.OpenRecordset("select Salesreturnhead.ReturnDate,SALESRETURNDETAILS.ItemType,SALESRETURNDETAILS.ItemName,SALESRETURNDETAILS.ItemName,SALESRETURNDETAILS.Size,SALESRETURNDETAILS.Qty,SALESRETURNDETAILS.PrRate from Salesreturnhead inner join SALESRETURNDETAILS on Salesreturnhead.SlNo=SALESRETURNDETAILS.SlNo where Salesreturnhead.ReturnDate between #" & Me.txtfrom_dt_m.Text & "/" & Me.txtfrom_dt_d.Text & "/" & Me.txtfrom_dt_y.Text & "# and #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set FromDtSalesreturn=FromDtSalesreturn + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and size='" & REC1("Size") & "' and PurchaseRate=" & REC1("PrRate"))
        REC1.MoveNext
        
        
    Wend
    '******************************For StockReturn********************************************
    Set REC1 = db.OpenRecordset("select PurchaseReturnHead.Rdate,PurchaseReturnDetails.ItemType,PurchaseReturnDetails.ItemName,PurchaseReturnDetails.Size,PurchaseReturnDetails.Qty,PurchaseReturnDetails.PrRate from PurchaseReturnHead inner join PurchaseReturnDetails on PurchaseReturnHead.SlNo=PurchaseReturnDetails.SlNo where PurchaseReturnHead.RDate between #" & Me.txtfrom_dt_m.Text & "/" & Me.txtfrom_dt_d.Text & "/" & Me.txtfrom_dt_y.Text & "# and #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set FromDtPurchasereturn=FromDtPurchasereturn + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and size='" & REC1("size") & "' and PurchaseRate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    '*****************************************************************************************************************************************************


    '--------------------Finding opening stock of To Date-----------------------------
    Set REC1 = db.OpenRecordset("select InvoiceHead.InvDate,InvoiceDetails.ItemType,InvoiceDetails.Itemname,InvoiceDetails.size,InvoiceDetails.PrRate,InvoiceDetails.Qty from invoicehead inner join InvoiceDetails on InvoiceHead.InvNo=InvoiceDetails.InvNo where InvoiceHead.InvDate  between #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "# and #" & Format(Date, "mm") & "/" & Format(Date, "dd") & "/" & Format(Date, "yyyy") & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToCrDt_Sales=ToCrDt_Sales + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and items='" & REC1("ItemName") & "' and size='" & REC1("Size") & "' and PurchaseRate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend

    Set REC1 = db.OpenRecordset("select PurchaseHead.Purchasedate,PurchaseDetails.ItemType,PurchaseDetails.ItemName,PurchaseDetails.Size,PurchaseDetails.Qty,PurchaseDetails.PrRate from PurchaseHead inner join PurchaseDetails on PurchaseHead.SlNo=PurchaseDetails.SlNo where PurchaseHead.Purchasedate between #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "# and #" & Format(Date, "mm") & "/" & Format(Date, "dd") & "/" & Format(Date, "yyyy") & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToCrDt_purchase=ToCrDt_purchase + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    Set REC1 = db.OpenRecordset("select Salesreturnhead.ReturnDate,SALESRETURNDETAILS.ItemType,SALESRETURNDETAILS.ItemName,SALESRETURNDETAILS.ItemName,SALESRETURNDETAILS.Size,SALESRETURNDETAILS.Qty,SALESRETURNDETAILS.PrRate from Salesreturnhead inner join SALESRETURNDETAILS on Salesreturnhead.SlNo=SALESRETURNDETAILS.SlNo where Salesreturnhead.ReturnDate between #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "# and #" & Format(Date, "mm") & "/" & Format(Date, "dd") & "/" & Format(Date, "yyyy") & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToCrDt_Salesreturn=ToCrDt_Salesreturn+ " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and size='" & REC1("Size") & "' and PurchaseRate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    '********************************For StockReturn**********************************
    Set REC1 = db.OpenRecordset("select PurchaseReturnHead.Rdate,PurchaseReturnDetails.ItemType,PurchaseReturnDetails.ItemName,PurchaseReturnDetails.Size,PurchaseReturnDetails.Qty,PurchaseReturnDetails.PrRate from PurchaseReturnHead inner join PurchaseReturnDetails on PurchaseReturnHead.SlNo=PurchaseReturnDetails.SlNo where PurchaseReturnHead.RDate  between #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "# and #" & Format(Date, "mm") & "/" & Format(Date, "dd") & "/" & Format(Date, "yyyy") & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToCrDt_Preturn=ToCrDt_Preturn + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    '*********************************************************************************

    '-------------------Opening Qty Of ToDate--------------------
    Set REC1 = db.OpenRecordset("select * from StockStatement")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToDtOpBalance=(Qty + ToCrDt_Sales+ToCrDt_Preturn) - (ToCrDt_purchase + ToCrDt_Salesreturn) where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("Items") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("Purchaserate"))
        REC1.MoveNext
    Wend
    '---------------------Finding Closing Stock of To Date------------------
    Set REC1 = db.OpenRecordset("select InvoiceHead.InvDate,InvoiceDetails.ItemType,InvoiceDetails.Itemname,InvoiceDetails.Size,InvoiceDetails.PrRate,InvoiceDetails.Qty from invoicehead inner join InvoiceDetails on InvoiceHead.InvNo=InvoiceDetails.InvNo where InvoiceHead.InvDate = #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToDt_Sales=ToDt_Sales + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and Size='" & REC1("Size") & "' and Purchaserate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    Set REC1 = db.OpenRecordset("select PurchaseHead.Purchasedate,PurchaseDetails.ItemType,PurchaseDetails.ItemName,PurchaseDetails.Size,PurchaseDetails.Qty,PurchaseDetails.PrRate from PurchaseHead inner join PurchaseDetails on PurchaseHead.SlNo=PurchaseDetails.SlNo where Purchasehead.Purchasedate = #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToDt_Purchase=ToDt_Purchase + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and Size='" & REC1("Size") & "' and Purchaserate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    Set REC1 = db.OpenRecordset("select Salesreturnhead.ReturnDate,SALESRETURNDETAILS.ItemType,SALESRETURNDETAILS.ItemName,SALESRETURNDETAILS.Qty,SALESRETURNDETAILS.PrRate,SALESRETURNDETAILS.size from Salesreturnhead inner join SALESRETURNDETAILS on Salesreturnhead.SlNo=SALESRETURNDETAILS.SlNo where Salesreturnhead.ReturnDate = #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToDt_SalesReturn=ToDt_SalesReturn + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and Size='" & REC1("Size") & "' and Purchaserate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    '*************************For Stockreturn******************************
    Set REC1 = db.OpenRecordset("select PurchaseReturnHead.Rdate,PurchaseReturnDetails.ItemType,PurchaseReturnDetails.ItemName,PurchaseReturnDetails.Size,PurchaseReturnDetails.Qty,PurchaseReturnDetails.PrRate from PurchaseReturnHead inner join PurchaseReturnDetails on PurchaseReturnHead.SlNo=PurchaseReturnDetails.SlNo where PurchaseReturnHead.RDate= #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ToDt_PReturn=ToDt_PReturn + " & REC1("Qty") & " where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("ItemName") & "' and Size='" & REC1("Size") & "' and Purchaserate=" & REC1("PrRate"))
        REC1.MoveNext
    Wend
    '**********************************************************************
    Set REC1 = db.OpenRecordset("select  * from StockStatement")
    While Not REC1.EOF
        db.Execute ("update StockStatement set ClosingBalance=(ToDtOpBalance - ToDt_Sales+ToDt_PReturn) + (ToDt_Purchase + ToDt_SalesReturn) where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("Items") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("Purchaserate"))
        REC1.MoveNext
    Wend

    '-------------------------------Finding Opening Balance Of From Date---------------------------------------
    Set REC1 = db.OpenRecordset("select  * from StockStatement")
    While Not REC1.EOF
        db.Execute ("update StockStatement set OpBalance=ClosingBalance + FromDtSales + FromDtPurchasereturn where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("Items") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("Purchaserate"))
        db.Execute ("update StockStatement set OPBalance=OPBalance - FromDtPurchase where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("Items") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("Purchaserate"))
        db.Execute ("update StockStatement set OPBalance=OPBalance - FromDtSalesreturn where ItemType='" & REC1("ItemType") & "' and Items='" & REC1("Items") & "' and size='" & REC1("Size") & "' and Purchaserate=" & REC1("Purchaserate"))

        REC1.MoveNext
    Wend

    db.Execute ("update StockStatement set Fromdate='" & Me.txtfrom_dt_d.Text & "/" & Me.txtfrom_dt_m.Text & "/" & Me.txtfrom_dt_y.Text & "',Todate='" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_y.Text & "'")
    db.Close
    MsgBox "Complete", vbOKOnly
    Set db = OpenDatabase(dbname)


End Sub

Private Sub Command1_Click()
    Me.CrystalReport1.SelectionFormula = "{StockStatement.Items} = '" & Me.cboitemname.Text & "' and {StockStatement.ItemType} = '" & Me.cboitemtype.Text & "'"
    CrystalReport1.PrintReport

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.CrystalReport1.ReportFileName = App.Path & "\stockstatement.rpt"
    Me.txtfrom_dt_d.Text = Format(Date, "dd")
    Me.txtfrom_dt_m.Text = Format(Date, "mm")
    Me.txtfrom_dt_y.Text = Format(Date, "yyyy")
    Me.txtto_dt_d.Text = Format(Date, "dd")
    Me.txtto_dt_m.Text = Format(Date, "mm")
    Me.txtto_dt_y.Text = Format(Date, "yyyy")
    Set REC1 = db.OpenRecordset("select distinct ItemType from ItemMaster ")
    If Not REC1.EOF Then
        While Not REC1.EOF
            Me.cboitemtype.AddItem (REC1("ItemType"))
            REC1.MoveNext
        Wend
        If Me.cboitemtype.ListCount > 0 Then
            Me.cboitemtype.ListIndex = 0
        End If
    End If

End Sub


Private Sub txtfrom_dt_d_GotFocus()
    Me.txtfrom_dt_d.SelStart = 0
    Me.txtfrom_dt_d.SelLength = Len(Me.txtfrom_dt_d.Text)
End Sub

Private Sub txtfrom_dt_m_GotFocus()
    Me.txtfrom_dt_m.SelStart = 0
    Me.txtfrom_dt_m.SelLength = Len(Me.txtfrom_dt_m.Text)
End Sub

Private Sub txtfrom_dt_y_GotFocus()
    Me.txtfrom_dt_y.SelStart = 0
    Me.txtfrom_dt_y.SelLength = Len(Me.txtfrom_dt_y.Text)
End Sub

Private Sub txtto_dt_d_GotFocus()
    Me.txtto_dt_d.SelStart = 0
    Me.txtto_dt_d.SelLength = Len(Me.txtto_dt_d.Text)
End Sub

Private Sub txtto_dt_m_GotFocus()
    Me.txtto_dt_m.SelStart = 0
    Me.txtto_dt_m.SelLength = Len(Me.txtto_dt_m.Text)
End Sub

Private Sub txtto_dt_y_GotFocus()
    Me.txtto_dt_y.SelStart = 0
    Me.txtto_dt_y.SelLength = Len(Me.txtfrom_dt_y.Text)
End Sub

