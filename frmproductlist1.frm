VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmproductlist1 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Product List"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmproductlist1.frx":0000
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "frmproductlist1.frx":0014
      TabIndex        =   1
      Top             =   480
      Width           =   11115
   End
   Begin VB.TextBox txtproductname 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   9795
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
      RecordSource    =   "ItemMaster"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmproductlist1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 13 Then
        If FORMNAME = "Purchase" Then
            frmStockin.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmStockin.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmStockin.txtmrp.Text = Me.DBGrid1.Columns(4)
            frmStockin.TxtVat.Text = Me.DBGrid1.Columns(5)
            frmStockin.cbounit.Text = Me.DBGrid1.Columns(8)
            frmStockin.txtPrate.Text = Me.DBGrid1.Columns(3)
            frmStockin.txttaxtype.Text = Me.DBGrid1.Columns(6)
            frmStockin.txthsn1.Text = Me.DBGrid1.Columns(7)
            Unload Me
            frmStockin.txthsn1.SetFocus
        End If
        If FORMNAME = "Purchase_Barcode" Then
            frmStockint.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmStockint.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmStockint.txtmrp.Text = Me.DBGrid1.Columns(4)
            frmStockint.TxtVat.Text = Me.DBGrid1.Columns(5)
            frmStockint.cbounit.Text = Me.DBGrid1.Columns(8)
            frmStockint.txtPrate.Text = Me.DBGrid1.Columns(3)
            frmStockint.txttaxtype.Text = Me.DBGrid1.Columns(6)
            'frmStockint.txthsn1.Text = Me.DBGrid1.Columns(7)
            Unload Me
            'frmStockint.txtpack.SetFocus
        End If
        If FORMNAME = "Invoice" Then
            frmInvoice.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmInvoice.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmInvoice.txtmrp.Text = Me.DBGrid1.Columns(3)
            frmInvoice.txtsalerate.Text = Me.DBGrid1.Columns(4)
            frmInvoice.TxtVat.Text = Me.DBGrid1.Columns(6)
            frmInvoice.cbounit.Text = Me.DBGrid1.Columns(7)
            frmInvoice.txttaxtype.Text = Me.DBGrid1.Columns(8)
            Unload Me
            frmInvoice.txtpack.SetFocus
        End If
        If FORMNAME = "itemslab" Then
            frmitemslabmaster.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmitemslabmaster.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmitemslabmaster.cbomrp.Text = Me.DBGrid1.Columns(3)
            Unload Me
            frmitemslabmaster.cbounit.SetFocus
        End If
        If FORMNAME = "damage" Then
            frmdamageentry.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmdamageentry.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmdamageentry.txtmrp.Text = Me.DBGrid1.Columns(3)
            frmdamageentry.TxtVat.Text = Me.DBGrid1.Columns(6)
            frmdamageentry.cbounit.Text = Me.DBGrid1.Columns(7)
            Unload Me
            frmdamageentry.cbounit.SetFocus
        End If
        If FORMNAME = "SalesReturn" Then
            frmSalesReturn.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmSalesReturn.cbobatch.Text = Trim(Me.DBGrid1.Columns(4))
            frmSalesReturn.txtmfgdate.Text = Trim(Me.DBGrid1.Columns(2))
            frmSalesReturn.txtexpdate.Text = Trim(Me.DBGrid1.Columns(3))
            frmSalesReturn.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmSalesReturn.txtmrp.Text = Me.DBGrid1.Columns(7)
            frmSalesReturn.txtsalerate.Text = Me.DBGrid1.Columns(8)
            frmSalesReturn.TxtVat.Text = Me.DBGrid1.Columns(10)
            'frmSalesReturn.cbounit.Text = Me.DBGrid1.Columns(7)
            frmSalesReturn.txttaxtype.Text = "SALES"
            Unload Me
            frmSalesReturn.txtpack.SetFocus
        End If
        If FORMNAME = "PurchaseReturn" Then
            frmpurchasereturn.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmpurchasereturn.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmpurchasereturn.cbobatch.Text = Trim(Me.DBGrid1.Columns(4))
            frmpurchasereturn.txtmfgdate.Text = Trim(Me.DBGrid1.Columns(2))
            frmpurchasereturn.txtexpdate.Text = Trim(Me.DBGrid1.Columns(3))
            frmpurchasereturn.txtmrp.Text = Me.DBGrid1.Columns(7)
            frmpurchasereturn.TxtVat.Text = Me.DBGrid1.Columns(10)
            'frmpurchasereturn.cbounit.Text = Me.DBGrid1.Columns(7)
            frmpurchasereturn.txttaxtype.Text = "SALES"
            Unload Me
            frmpurchasereturn.txtpack.SetFocus
        End If
        If FORMNAME = "DamageEntry" Then
            frmdamageentry.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmdamageentry.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmdamageentry.cbobatch.Text = Trim(Me.DBGrid1.Columns(4))
            frmdamageentry.txtmfgdate.Text = Trim(Me.DBGrid1.Columns(2))
            frmdamageentry.txtexpdate.Text = Trim(Me.DBGrid1.Columns(3))
            frmdamageentry.txtmrp.Text = Me.DBGrid1.Columns(7)
            frmdamageentry.TxtVat.Text = Me.DBGrid1.Columns(10)
            frmdamageentry.txtPrate.Text = Me.DBGrid1.Columns(13)
            'frmpurchasereturn.cbounit.Text = Me.DBGrid1.Columns(7)
            frmdamageentry.txttaxtype.Text = "SALES"
            Unload Me
            frmdamageentry.txtpack.SetFocus
        End If
    End If
End Sub

Private Sub DBGrid2_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errtrap
    If KeyCode = 27 Then
        Unload Me
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Form_Load()
    Me.Top = 3000
    Me.Left = 1740
    Me.Data1.databasename = dbname
    'Me.Data1.RecordSource = "select * from Stock where itemname like '" & SEARCHWORD & "*' order by itemname"
    'Me.Data1.RecordSource = "Select Stock.ProductCode,Stock.itemname,Stock.Lose,Stock.MRP,Stock.SaleRate,Stock.Qty,ItemMaster.Tax,Stock.UniyType,ItemMaster.tax_type from Stock inner Join ItemMaster on Stock.ProductCode=ItemMaster.ProductCode where Stock.ItemName Like '" & SEARCHWORD & "*'"
   
    Me.Data1.RecordSource = "select * from itemmaster where Item like '" & Replace(SEARCHWORD, "'", "''") & "*' order by Item"
    Me.Data1.Refresh
   
End Sub

Private Sub txtproductname_Change()
'Me.Data1.RecordSource = "Select Stock.ProductCode,Stock.itemname,Stock.Lose,Stock.MRP,Stock.SaleRate,Stock.Qty,ItemMaster.Tax,Stock.UniyType,ItemMaster.tax_type from Stock inner Join ItemMaster on Stock.ProductCode=ItemMaster.ProductCode where Stock.ItemName Like '" & Me.txtproductname.Text & "*'"
  
        Me.Data1.RecordSource = "select * from itemmaster where Item like '*" & Replace(Me.txtproductname.Text, "'", "''") & "*' order by Item"
        Me.Data1.Refresh
  
End Sub


Private Sub txtproductname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.DBGrid1.SetFocus
    End If
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 40 Then
        Me.DBGrid1.SetFocus
    End If
End Sub

Private Sub txtproductname_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
