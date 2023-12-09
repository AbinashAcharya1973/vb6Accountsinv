VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmutility 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   525
      Left            =   1710
      TabIndex        =   7
      Top             =   1800
      Width           =   1245
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   525
      Left            =   1710
      TabIndex        =   6
      Top             =   1170
      Width           =   1245
   End
   Begin VB.CommandButton Command5 
      Caption         =   "test tran"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmutility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432

Private Sub Command1_Click()
Set rec1 = db.OpenRecordset("select * from PurchaseHead")
ItemSlno = 0
While Not rec1.EOF
    ItemSlno = 0
    Set rec2 = db.OpenRecordset("select * from PurchaseDetails where Slno=" & rec1("Slno"))
    While Not rec2.EOF
        ItemSlno = ItemSlno + 1
        db.Execute ("Update PurchaseDetails set Item_Slno=" & ItemSlno & " where Slno=" & rec1("Slno") & " and ProductCode=" & rec2("ProductCode"))
    rec2.MoveNext
    Wend
rec1.MoveNext
Wend






'Set REC1 = db.OpenRecordset("select * from partydr")
'While Not REC1.EOF
'    db.Execute ("update LedgerMaster set Tin='" & REC1("Tin") & "',Phone='" & REC1("Phone") & "' where Accid=" & REC1("AccId"))
'REC1.MoveNext
'Wend

'Dim Db1 As DAO.Database
'Set Db1 = OpenDatabase(App.Path & "\ac.mdb", False, False, ";PWD=suhana")
''PartCr
'Set REC1 = db.OpenRecordset("select max(AccId) as max_id from LedgerMaster")
'If Not IsNull(REC1!max_id) Then
'    AccID = REC1!max_id + 1
'Else
'    AccID = 1
'End If
'Set REC1 = db1.OpenRecordset("select * from ac where acsub_cd>=24")
'While Not REC1.EOF
'    '/get area name
'    Set REC2 = db1.OpenRecordset("select * from ACSUB where ACSUB_CD=" & REC1("acsub_cd"))
'    If Not REC2.EOF Then
'        db.Execute ("Insert into PartyDr (Party,Address,Phone,TIN,AccId,Zone) Values('" & REC1("ac_nm") & "','" & REC1("address") & "','" & REC1("Ph") & "','TIN-" & REC1("stno") & "'," & AccID & ",'" & REC2("ACSUB_NM") & "')")
'        db.Execute ("Insert into LedgerMaster (AccID,AccName,GroupID,Groupname,Address1,TIN) Values(" & AccID & ",'" & REC1("ac_nm") & "',17,'Sundry Debtor','" & REC1("Address") & "','TIN-" & REC1("stno") & "')")
'    End If
'    AccID = AccID + 1
'REC1.MoveNext
'Wend
'Slno = 1
'Set REC1 = db.OpenRecordset("select * from ZoneMaster")
'While Not REC1.EOF
'    db.Execute ("Update PartyDr set ZoneCode=" & REC1("Slno") & " where Zone='" & REC1("ZoneName") & "'")
'    Slno = Slno + 1
'REC1.MoveNext
'Wend
'Set REC1 = db.OpenRecordset("select * from ItemMaster")
'While Not REC1.EOF
'    db.Execute ("Insert into Stock (ProductCode,itemname,MRP,PRate,UniyType) values(" & REC1("ProductCode") & ",'" & REC1("item") & "'," & REC1("MRP") & "," & REC1("Purchaserate") & ",'" & REC1("UnitType") & "')")
'REC1.MoveNext
'Wend

End Sub

Private Sub Command2_Click()
Set rec1 = db.OpenRecordset("select * from Otheroutstanding where Dr>0")
While Not rec1.EOF
    db.Execute ("Update LedgerMaster set OBalance=" & rec1("Dr") & ",BalanceType='Dr' where AccId=" & rec1("AccId"))
rec1.MoveNext
Wend
Set rec1 = db.OpenRecordset("select * from Otheroutstanding where Cr>0")
While Not rec1.EOF
    db.Execute ("Update LedgerMaster set OBalance=" & rec1("Cr") & ",BalanceType='Cr' where AccId=" & rec1("AccId"))
rec1.MoveNext
Wend
End Sub

Private Sub Command3_Click()
  'On Error Resume Next

    

    db.Execute ("Update Stock set qty=0")

    Set rec1 = db.OpenRecordset("select count(*) as total_record from ItemMaster") ' where CategoryId=" & Me.cboitemtype.ItemData(Me.cboitemtype.ListIndex))
    If Not IsNull(rec1("total_record")) Then
        pr_value = 100 / rec1("total_record")
    End If

    '//Process progress bar
    For i = 0 To 4000000
        ' Do nothing, but wait
        ' To show up the progress bar proceeding
    Next i


    Set rec1 = db.OpenRecordset("select * from ItemMaster") 'where CategoryId=" & Me.cboitemtype.ItemData(Me.cboitemtype.ListIndex) & " order by ProductCode")
    While Not rec1.EOF
        Slno = 1
        If Not IsNull(rec1!OpeningStock) Then
            opstock = rec1!OpeningStock
        Else
            opstock = 0
        End If
        '/calculate purchase
        Set rec2 = db.OpenRecordset("select sum(qty+Free_Qty) as total_qty from PurchaseDetails where ProductCode=" & rec1("ProductCode"))
        If Not rec2.EOF Then
            If Not IsNull(rec2!Total_qty) Then
                purchase = rec2!Total_qty
            Else
                purchase = 0
            End If
        End If
        'sales return
        Set rec2 = db.OpenRecordset("Select sum(qty+Free_Qty) as total_qty from Salesreturndetails where Productcode=" & rec1("Productcode"))
        If Not rec2.EOF Then
            If Not IsNull(rec2!Total_qty) Then
                sale_return = rec2!Total_qty
            Else
                sale_return = 0
            End If
        End If
        'calculate sales
        Set rec2 = db.OpenRecordset("Select sum(qty+Free_Qty) as total_qty from InvoiceDetails where ProductCode=" & rec1("Productcode"))
        If Not rec2.EOF Then
            If Not IsNull(rec2!Total_qty) Then
                Sales_qty = rec2!Total_qty
            Else
                Sales_qty = 0
            End If
        End If
        'calculate purchasereturn
        Set rec2 = db.OpenRecordset("select sum(qty+free_qty) as total_qty from PurchaseReturnDetails where ProductCode=" & rec1("ProductCode"))
        If Not rec2.EOF Then
            If Not IsNull(rec2!Total_qty) Then
                Pur_return = rec2!Total_qty
            Else
                Pur_return = 0
            End If
        End If

        stock = (opstock + purchase + sale_return) - (Sales_qty + Pur_return)

        db.Execute ("Update Stock set Qty=" & stock & " where ProductCode=" & rec1("ProductCode"))
        rec1.MoveNext
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + pr_value
    Wend
    Me.ProgressBar1.Value = 0
    
    'Me.Data1.RecordSource = "select * from Stock INNER JOIN itemmaster on itemmaster.productcode=stock.productcode where itemmaster.categoryid=" & Me.cboitemtype.ItemData(Me.cboitemtype.ListIndex)
    'Me.Data1.Refresh
End Sub

Private Sub Command4_Click()
Set rec1 = db.OpenRecordset("select * from stock")
While Not rec1.EOF
    db.Execute ("update itemmaster set Openingstock =" & rec1("qty") & " where productcode=" & rec1("productcode"))
rec1.MoveNext
Wend
End Sub

Private Sub Command5_Click()
Dim dbSession As DAO.Workspace
Dim cDb As DAO.Database
Set dbSession = DBEngine(0)
Set cDb = dbSession.OpenDatabase(db.Name)
dbSession.BeginTrans
cDb.Execute "update companymaster set company='XXx'"
ans = MsgBox("Commit the change?", vbYesNo)
If ans = 6 Then
    dbSession.CommitTrans
Else
    dbSession.Rollback
End If
End Sub

Private Sub Command6_Click()
Set rec1 = db.OpenRecordset("select * from stock")
While Not rec1.EOF
    Set rec2 = db.OpenRecordset("select * from stockdetails where batchno='Opening Stock' and productcode=" & rec1("productcode"))
    If rec2.EOF Then
        db.Execute ("insert into Stockdetails (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,mfgdate,batchno,expdate,hsn) values('" & rec1("ProductType") & "','" & rec1("itemtype") & "','" & rec1("itemname") & "','" & rec1("brand") & "','" & rec1("barcode") & "','" & rec1("size") & "'," & rec1("MRP") & "," & rec1("PRate") & "," & rec1("Qty") & "," & rec1("ProductCode") & "," & rec1("Vat") & "," & rec1("Lose") & ",'" & rec1("UniyType") & "'," & rec1("SaleRate") & ",'" & Format(Date, "dd/mm/yyyy") & "','Opening Stock',' ','" & rec1("hsn") & "')")
    End If
    
    rec1.MoveNext
Wend
End Sub

Private Sub Command7_Click()
Set rec1 = db.OpenRecordset("select * from itemmaster")
While Not rec1.EOF
    Set rec2 = db.OpenRecordset("select * from stock where productcode=" & rec1("productcode"))
    If rec2.EOF Then
        db.Execute ("insert into Stock (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,hsn) values('" & rec1("ProductType") & "','" & rec1("itemtype") & "','" & rec1("itemname") & "','" & rec1("brand") & "','" & rec1("barcode") & "','" & rec1("size") & "'," & rec1("MRP") & "," & rec1("PRate") & "," & rec1("Qty") & "," & rec1("ProductCode") & "," & rec1("Vat") & "," & rec1("Lose") & ",'" & rec1("UniyType") & "'," & rec1("SaleRate") & ",'" & rec1("hsn") & "')")
        db.Execute ("insert into Stockdetails (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,mfgdate,batchno,expdate,hsn) values('" & rec1("ProductType") & "','" & rec1("itemtype") & "','" & rec1("itemname") & "','" & rec1("brand") & "','" & rec1("barcode") & "','" & rec1("size") & "'," & rec1("MRP") & "," & rec1("PRate") & "," & rec1("Qty") & "," & rec1("ProductCode") & "," & rec1("Vat") & "," & rec1("Lose") & ",'" & rec1("UniyType") & "'," & rec1("SaleRate") & ",'" & Format(Date, "dd/mm/yyyy") & "','Opening Stock',' ','" & rec1("hsn") & "')")
    End If
    
    rec1.MoveNext
Wend

End Sub

Private Sub Form_Load()
'    Set REC1 = db.OpenRecordset("Select * from ItemMaster")
'    While Not REC1.EOF
'        db.Execute ("update stock set qty=" & REC1("Openingstock") & " where ProductCode=" & REC1("ProductCode"))
'        REC1.MoveNext
'    Wend
End Sub
