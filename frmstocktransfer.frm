VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmstocktransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Transfer"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   8475
   Begin VB.TextBox txttotalqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6750
      TabIndex        =   20
      Text            =   "0"
      Top             =   7530
      Width           =   1515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\videotutorial\EnLite\EnliteFinal\DATA\2022-2023\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tempstocktran"
      Top             =   7620
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5370
      TabIndex        =   18
      Top             =   7980
      Width           =   1065
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   17
      Top             =   7980
      Width           =   1065
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3150
      TabIndex        =   16
      Top             =   7980
      Width           =   1065
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   15
      Top             =   7980
      Width           =   1065
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmstocktransfer.frx":0000
      Height          =   4455
      Left            =   90
      OleObjectBlob   =   "frmstocktransfer.frx":0014
      TabIndex        =   14
      Top             =   2970
      Width           =   8325
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6780
      TabIndex        =   12
      Text            =   "0"
      Top             =   2340
      Width           =   1575
   End
   Begin VB.ComboBox cbobatchno 
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
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2310
      Width           =   3075
   End
   Begin VB.ComboBox cboproduct 
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
      Left            =   1020
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1860
      Width           =   7335
   End
   Begin VB.ComboBox cbotostock 
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
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1170
      Width           =   7335
   End
   Begin VB.ComboBox cbofromstock 
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
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   660
      Width           =   7335
   End
   Begin VB.TextBox txtslno 
      Appearance      =   0  'Flat
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
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1575
   End
   Begin MSMask.MaskEdBox txtdate 
      Height          =   315
      Left            =   6660
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   21
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label lblqty 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   4110
      TabIndex        =   19
      Top             =   2370
      Width           =   1485
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6210
      TabIndex        =   13
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6060
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmstocktransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec1 As Recordset, rec2 As Recordset, deletef As Boolean

Private Sub cbobatchno_Change()
Set rec2 = db.OpenRecordset("select * from " & Me.cbofromstock.Text & "details where productcode=" & Me.cboproduct.ItemData(Me.cboproduct.ListIndex) & " and batchno='" & Me.cbobatchno.Text & "'")
If Not rec2.EOF Then
    Me.lblqty.Caption = rec2("qty")
End If
End Sub

Private Sub cbobatchno_Click()
cbobatchno_Change
End Sub

Private Sub cbobatchno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtqty.SetFocus
End If
End Sub

Private Sub cbofromstock_Change()
Set rec1 = db.OpenRecordset("select * from " & Me.cbofromstock.Text)
While Not rec1.EOF
    Me.cboproduct.AddItem rec1("itemname")
    Me.cboproduct.ItemData(Me.cboproduct.NewIndex) = rec1("productcode")
    rec1.MoveNext
Wend
If Me.cboproduct.ListCount > 0 Then
    Me.cboproduct.ListIndex = 0
End If
End Sub

Private Sub cbofromstock_Click()
cbofromstock_Change
End Sub

Private Sub cbofromstock_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cbotostock.SetFocus
End If
End Sub

Private Sub cboproduct_Change()
Set rec2 = db.OpenRecordset("select * from " & Me.cbofromstock.Text & "details where productcode=" & Me.cboproduct.ItemData(Me.cboproduct.ListIndex))
If Me.cbobatchno.ListCount > 0 Then
    Me.cbobatchno.Clear
End If
While Not rec2.EOF
    Me.cbobatchno.AddItem rec2("batchno")
    'Me.lblqty.Caption = rec2("qty")
    rec2.MoveNext
Wend
If Me.cbobatchno.ListCount > 0 Then
    Me.cbobatchno.ListIndex = 0
End If
End Sub

Private Sub cboproduct_Click()
cboproduct_Change
End Sub

Private Sub cboproduct_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cbobatchno.SetFocus
End If
If KeyCode = 27 Then
    Me.cmdsave.SetFocus
End If
End Sub

Private Sub cbotostock_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cboproduct.SetFocus
End If
End Sub

Private Sub CmdDelete_Click()
deletef = True
Me.txtslno.SetFocus
Me.txtslno.Locked = False
End Sub

Private Sub CmdEdit_Click()
Me.txtslno.Locked = False
Me.txtslno.SetFocus
deletef = False
End Sub

Private Sub cmdprint_Click()
frmprintstocktran.Show 0
End Sub

Private Sub CmdSave_Click()
    ans = MsgBox("Save the Stock Transfer", vbYesNo)
    If ans = 6 Then
        Set rec = db.OpenRecordset("select * from stocktranhead where challanno=" & Me.txtslno.Text)
        If Not rec.EOF Then
            Set rec1 = db.OpenRecordset("select * from stocktrandetails where challanno=" & Me.txtslno.Text)
            While Not rec1.EOF
                db.Execute "update " & rec("fromstock") & " set qty=qty+" & rec1("qty") & " where productcode=" & rec1("productcode")
                db.Execute "update " & rec("fromstock") & "details set qty=qty+" & rec1("qty") & " where productcode=" & rec1("productcode") & " and batchno='" & rec1("batchno") & "'"

                db.Execute "update " & rec("tostock") & " set qty=qty-" & rec1("qty") & " where productcode=" & rec1("productcode")
                db.Execute "update " & rec("tostock") & "details set qty=qty-" & rec1("qty") & " where productcode=" & rec1("productcode") & " and batchno='" & rec1("batchno") & "'"

                rec1.MoveNext
            Wend
            db.Execute ("delete * from stocktranhead where challanno=" & Me.txtslno.Text)
            db.Execute ("delete * from stocktrandetails where challanno=" & Me.txtslno.Text)
        End If
        fromstockid = Me.cbofromstock.ItemData(Me.cbofromstock.ListIndex)
        tostockid = Me.cbotostock.ItemData(Me.cbotostock.ListIndex)
        db.Execute "insert into stocktranhead (ChallanNO,ChallanDaate,FromStock,FromStockCode,ToStock,ToStockCode,TotalQty) values(" & Me.txtslno.Text & ",'" & Me.txtdate.Text & "','" & Me.cbofromstock.Text & "'," & fromstockid & ",'" & Me.cbotostock.Text & "'," & tostockid & "," & Me.txttotalqty.Text & ")"

        Set rec = db.OpenRecordset("select * from tempstocktran")
        While Not rec.EOF
            db.Execute "insert into stocktrandetails (ChallanNo,ProductCode,ProductName,Qty,Batchno) values(" & Me.txtslno.Text & "," & rec("productcode") & ",'" & rec("ProductName") & "'," & rec("Qty") & ",'" & rec("Batchno") & "')"
            db.Execute "update " & Me.cbofromstock.Text & " set qty=qty-" & rec("qty") & " where productcode=" & rec("productcode")
            db.Execute "update " & Me.cbofromstock.Text & "details set qty=qty-" & rec("qty") & " where productcode=" & rec("productcode") & " and batchno='" & rec("batchno") & "'"
            Set rec1 = db.OpenRecordset("select * from " & Me.cbotostock.Text & " where productcode=" & rec("productcode"))
            If Not rec1.EOF Then
                db.Execute "update " & Me.cbotostock.Text & " set qty=qty+" & rec("qty") & " where productcode=" & rec("productcode")

            Else
                Set rec1 = db.OpenRecordset("select * from itemmaster where productcode=" & rec("productcode"))
                db.Execute ("insert into " & Me.cbotostock.Text & " (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,hsn) values('" & rec1("ProductType") & "','" & rec1("itemtype") & "','" & rec1("item") & "','" & rec1("brand") & "','" & rec1("barcode") & "','" & rec1("size") & "'," & rec1("MRP") & "," & rec1("PurchaseRate") & "," & rec("Qty") & "," & rec("ProductCode") & "," & rec1("Tax") & "," & rec1("Lose") & ",'" & rec1("UnitType") & "'," & rec1("SaleRate") & ",'" & rec1("hsn") & "')")
            End If
            Set rec1 = db.OpenRecordset("select * from " & Me.cbotostock & "details where productcode=" & rec("productcode") & " and batchno='" & rec("batchno") & "'")
            If Not rec1.EOF Then
                db.Execute "update " & Me.cbotostock.Text & "details set qty=qty+" & rec("qty") & " where productcode=" & rec("productcode") & " and batchno='" & rec("batchno") & "'"
            Else
                Set rec2 = db.OpenRecordset("select * from " & Me.cbofromstock.Text & "details where productcode=" & rec("productcode") & " and batchno='" & rec("batchno") & "'")
                db.Execute ("insert into " & Me.cbotostock.Text & "details (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,mfgdate,batchno,expdate,hsn) values('" & rec2("ProductType") & "','" & rec2("itemtype") & "','" & rec2("itemname") & "','" & rec2("brand") & "','" & rec2("barcode") & "','" & rec2("size") & "'," & rec2("MRP") & "," & rec2("PRate") & "," & rec("Qty") & "," & rec("ProductCode") & "," & rec2("Vat") & "," & rec2("Lose") & ",'" & rec2("UniyType") & "'," & rec2("SaleRate") & ",'" & rec2("mfgdate") & "','" & rec("batchno") & "','" & rec2("expdate") & "','" & rec2("hsn") & "')")
            End If

            rec.MoveNext
        Wend
        db.Execute "delete * from tempstocktran"
        Me.txttotalqty.Text = 0
        Me.Data1.Refresh
        Me.DBGrid1.Refresh
        Me.txtslno.Text = Val(Me.txtslno.Text) + 1
        deletef = False
        Me.txtdate.SetFocus
    End If
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
Me.txttotalqty.Text = Val(Me.txttotalqty.Text) - Val(Me.DBGrid1.Columns(3))
End Sub

Private Sub Form_Load()
Me.Data1.databasename = db.Name
Me.txtdate.Text = Format(Date, "dd/mm/yyyy")
Set rec = db.OpenRecordset("select max(challanno) as maxslno from stocktranhead")
If Not IsNull(rec!maxslno) Then
    Me.txtslno.Text = rec!maxslno + 1
Else
    Me.txtslno.Text = 1
End If
Me.cbofromstock.AddItem "Stock"
Me.cbofromstock.ItemData(Me.cbofromstock.NewIndex) = 0
    Set rec1 = db.OpenRecordset("select * from stockpoints")
    While Not rec1.EOF
        Me.cbofromstock.AddItem rec1("stockpoint")
        Me.cbofromstock.ItemData(Me.cbofromstock.NewIndex) = rec1("stockpointid")
        rec1.MoveNext
    Wend
    Me.cbofromstock.ListIndex = 0
Me.cbotostock.AddItem "Stock"
Me.cbotostock.ItemData(Me.cbotostock.NewIndex) = 0
    Set rec1 = db.OpenRecordset("select * from stockpoints")
    While Not rec1.EOF
        Me.cbotostock.AddItem rec1("stockpoint")
        Me.cbotostock.ItemData(Me.cbotostock.NewIndex) = rec1("stockpointid")
        rec1.MoveNext
    Wend
    Me.cbotostock.ListIndex = 0
  deletef = False
End Sub

Private Sub frmproduct_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
db.Execute "delete * from tempstocktran"
End Sub

Private Sub txtdate_GotFocus()
Me.txtdate.SelStart = 0
Me.txtdate.SelLength = Len(Me.txtdate.Text)
End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cbofromstock.SetFocus
End If
End Sub

Private Sub txtqty_GotFocus()
Me.txtqty.SelStart = 0
Me.txtqty.SelLength = Len(Me.txtqty.Text)
End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    tempqty = Val(Me.lblqty.Caption)
    tempinqty = Val(Me.txtqty.Text)
    If tempinqty <= tempqty Then
    db.Execute "insert into tempstocktran (ProductCode,ProductName,Batchno,qty) values(" & Me.cboproduct.ItemData(Me.cboproduct.ListIndex) & ",'" & Me.cboproduct.Text & "','" & Me.cbobatchno.Text & "'," & Me.txtqty.Text & ")"
    Me.Data1.Refresh
    Me.DBGrid1.Refresh
    Me.txttotalqty.Text = Val(Me.txttotalqty.Text) + Val(Me.txtqty.Text)
    Me.cboproduct.SetFocus
    Else
        MsgBox "Stock Qty is Less", vbCritical
    End If
End If
End Sub

Private Sub txtslno_GotFocus()
Me.txtslno.SelStart = 0
Me.txtslno.SelLength = Len(Me.txtslno.Text)
End Sub

Private Sub txtSlno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec = db.OpenRecordset("select * from stocktranhead where challanno=" & Me.txtslno.Text)
        If rec.EOF Then
            MsgBox "Challan No not Found", vbCritical
        Else
            i = 0
            Me.txtdate.Text = Format(rec("ChallanDaate"), "dd/mm/yyyy")
            Me.txttotalqty.Text = rec("totalqty")
            Me.txtdate.SetFocus
            Do While i < Me.cbofromstock.ListCount
                If Me.cbofromstock.List(i) = rec("fromstock") Then

                    Me.cbofromstock.ListIndex = i
                    Exit Do
                Else
                    i = i + 1
                End If
            Loop
            i = 0
            Do While i < Me.cbotostock.ListCount
                If Me.cbotostock.List(i) = rec("tostock") Then
                    Me.cbotostock.ListIndex = i
                    Exit Do
                Else
                    i = i + 1
                End If
            Loop
            Set rec1 = db.OpenRecordset("select * from stocktrandetails where challanno=" & Me.txtslno.Text)
            While Not rec1.EOF
                db.Execute "insert into tempstocktran (ProductCode,ProductName,Qty,Batchno) values(" & rec1("ProductCode") & ",'" & rec1("ProductName") & "'," & rec1("Qty") & ",'" & rec1("Batchno") & "')"
                rec1.MoveNext
            Wend
            Me.Data1.Refresh
            Me.DBGrid1.Refresh
            If deletef = True Then
                ans = MsgBox("Delete the Challan?", vbYesNo)
                If ans = 6 Then
                    Set rec1 = db.OpenRecordset("select * from stocktrandetails where challanno=" & Me.txtslno.Text)
                    While Not rec1.EOF
                        db.Execute "update " & rec("fromstock") & " set qty=qty+" & rec1("qty") & " where productcode=" & rec1("productcode")
                        db.Execute "update " & rec("fromstock") & "details set qty=qty+" & rec1("qty") & " where productcode=" & rec1("productcode") & " and batchno='" & rec1("batchno") & "'"

                        db.Execute "update " & rec("tostock") & " set qty=qty-" & rec1("qty") & " where productcode=" & rec1("productcode")
                        db.Execute "update " & rec("tostock") & "details set qty=qty-" & rec1("qty") & " where productcode=" & rec1("productcode") & " and batchno='" & rec1("batchno") & "'"

                        rec1.MoveNext
                    Wend
                    db.Execute ("delete * from stocktranhead where challanno=" & Me.txtslno.Text)
                    db.Execute ("delete * from stocktrandetails where challanno=" & Me.txtslno.Text)
                    db.Execute ("delete * from tempstocktran")
                    Me.Data1.Refresh
                    Me.DBGrid1.Refresh
                    Me.txtdate.SetFocus
                    Set rec2 = db.OpenRecordset("select max(challanno) as maxslno from stocktranhead")
                    If Not IsNull(rec2!maxslno) Then
                        Me.txtslno.Text = rec2!maxslno + 1
                    Else
                        Me.txtslno.Text = 1
                    End If
                    deletef = False
                End If
            End If
        End If
    End If
End Sub
