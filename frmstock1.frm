VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmstock1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock View"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10740
   Begin VB.ComboBox cbostock 
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
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   3075
   End
   Begin VB.CommandButton cmdexport 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export[xls]"
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
      Left            =   5580
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Stock"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "StockDetails"
      Top             =   6300
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "All Stock View"
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      Height          =   315
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3660
      Width           =   1215
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   3660
      Width           =   1335
   End
   Begin VB.TextBox txtitemname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Top             =   120
      Width           =   4305
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Short Expiry"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmstock1.frx":0000
      Height          =   3555
      Left            =   180
      OleObjectBlob   =   "frmstock1.frx":0014
      TabIndex        =   1
      Top             =   4140
      Width           =   10395
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmstock1.frx":244B
      Height          =   3015
      Left            =   180
      OleObjectBlob   =   "frmstock1.frx":245F
      TabIndex        =   6
      Top             =   540
      Width           =   10455
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3420
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblqty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   9180
      TabIndex        =   9
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   7500
      TabIndex        =   8
      Top             =   3660
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "frmstock1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset

Private Sub cbostock_Change()
Data1.RecordSource = "select * from " & Me.cbostock
Data1.Refresh
End Sub

Private Sub cbostock_Click()
cbostock_Change
End Sub

Private Sub cmdexport_Click()
    Me.MousePointer = vbHourglass
    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelWS.Cells(1, 1).Value = "Product code"
    excelWS.Cells(1, 2).Value = "Category"
    excelWS.Cells(1, 3).Value = "SubCategory"
    excelWS.Cells(1, 4).Value = "Brand"
    excelWS.Cells(1, 5).Value = "Item Name"
    excelWS.Cells(1, 6).Value = "HSN"
    excelWS.Cells(1, 7).Value = "Tax %"
    excelWS.Cells(1, 8).Value = "Purchase Price"
    excelWS.Cells(1, 9).Value = "Sale Price"
    excelWS.Cells(1, 10).Value = "MRP"
    excelWS.Cells(1, 11).Value = "Qty"
    excelWS.Cells(1, 12).Value = "Stock Value"
    Set rec1 = db.OpenRecordset("select * from stock")
    RowCount = 2
    While Not rec1.EOF
        excelWS.Cells(1, 1).Value = rec1("productcode")
        excelWS.Cells(1, 2).Value = rec1("producttype")
        excelWS.Cells(1, 3).Value = rec1("itemtype")
        excelWS.Cells(1, 4).Value = rec1("brand")
        excelWS.Cells(1, 5).Value = rec1("itemname")
        excelWS.Cells(1, 6).Value = rec1("HSN")
        excelWS.Cells(1, 7).Value = rec1("vat")
        excelWS.Cells(1, 8).Value = rec1("prate")
        excelWS.Cells(1, 9).Value = rec1("salerate")
        excelWS.Cells(1, 10).Value = rec1("mrp")
        excelWS.Cells(1, 11).Value = rec1("qty")
        excelWS.Cells(1, 12).Value = rec1("qty") * rec1("prate")
        RowCount = RowCount + 1
        rec1.MoveNext
    Wend
    excelApp.Visible = True
    Me.MousePointer = 0

End Sub

Private Sub cmdprint_Click()
Me.CrystalReport1.ReportFileName = App.Path & "\itemwisestock.rpt"
Me.CrystalReport1.PrintReport
End Sub

Private Sub Command1_Click()
Me.Data1.RecordSource = "select * from Stock order by ItemName"
Me.Data1.Refresh
End Sub

Private Sub Command3_Click()

End Sub

'Private Sub Command2_Click()
'Me.Data2.RecordSource = "SELECT * FROM STOCKdetails where  datediff('m',date(),FORMAT('01/'+EXPdate,'MM/YYYY')) <=6 and qty>0"
'Data2.Refresh
'End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    Me.DBGrid1.AllowDelete = True
    Me.DBGrid1.AllowUpdate = True
End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
Data2.RecordSource = "SELECT * FROM " & Me.cbostock.Text & "details WHERE PRODUCTCODE=" & Me.DBGrid1.Columns(0)
Data2.Refresh
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Data2.databasename = dbname
    Me.cbostock.AddItem "Stock"
    Set rec1 = db.OpenRecordset("select * from stockpoints")
    While Not rec1.EOF
        Me.cbostock.AddItem rec1("stockpoint")
        rec1.MoveNext
    Wend
    Me.cbostock.ListIndex = 0
    
'    Set rec1 = db.OpenRecordset("select sum(prate*qty) as totalamount from stock")
'    If Not IsNull(rec1!totalamount) Then
'        'Me.txtamount.Text = Format(rec1!totalamount, "############0.00")
'    End If
'    Set rec1 = db.OpenRecordset("select sum(qty) as totalqty from stock")
'    If Not IsNull(rec1!totalqty) Then
'        'Me.lblqty.Caption = Format(rec1!totalqty, "############0.00")
'    End If
End Sub

Private Sub txtitemname_Change()
Me.Data1.RecordSource = "select * from " & Me.cbostock.Text & " where ItemName Like '" & Replace(Me.txtitemname.Text, "'", "''") & "*'"
Me.Data1.Refresh
End Sub

