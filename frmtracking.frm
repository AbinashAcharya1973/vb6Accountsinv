VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmtracking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Tracker"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   18600
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   14910
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6570
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frmtracking.frx":0000
      Height          =   6705
      Left            =   14490
      OleObjectBlob   =   "frmtracking.frx":0014
      TabIndex        =   12
      Top             =   810
      Width           =   4095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ComboBox cboitemname 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   4320
   End
   Begin VB.TextBox txtproductcode 
      Appearance      =   0  'Flat
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
      Left            =   4500
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmtracking.frx":09E7
      Height          =   6735
      Left            =   120
      OleObjectBlob   =   "frmtracking.frx":09FB
      TabIndex        =   0
      Top             =   780
      Width           =   7095
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmtracking.frx":13CA
      Height          =   6735
      Left            =   7320
      OleObjectBlob   =   "frmtracking.frx":13DE
      TabIndex        =   1
      Top             =   780
      Width           =   7095
   End
   Begin VB.Label Label3 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14550
      TabIndex        =   13
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label lblsavg 
      Caption         =   "Average Price:0.00"
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
      Left            =   10560
      TabIndex        =   11
      Top             =   7740
      Width           =   3195
   End
   Begin VB.Label lblsmax 
      Caption         =   "Maximum Price:0.00"
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
      Left            =   7380
      TabIndex        =   10
      Top             =   8160
      Width           =   2835
   End
   Begin VB.Label lblsmin 
      Caption         =   "Minimum Price:0.00"
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
      Left            =   7380
      TabIndex        =   9
      Top             =   7740
      Width           =   2775
   End
   Begin VB.Label lblpavg 
      Caption         =   "Average Price:0.00"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   7740
      Width           =   3075
   End
   Begin VB.Label lblpmax 
      Caption         =   "Maximum Price:0.00"
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
      TabIndex        =   7
      Top             =   8160
      Width           =   3075
   End
   Begin VB.Label lblpmin 
      Caption         =   "Minimum Price:0.00"
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
      TabIndex        =   6
      Top             =   7740
      Width           =   2955
   End
   Begin VB.Label Label2 
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   3
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "frmtracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboitemname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.Data3.RecordSource = "select Batchno,PurchaseInfo from stockdetails where productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex)
    Data3.Refresh
    Me.Data1.RecordSource = "select p.slno,p.InvNo,p.PurchaseDate,pd.PrRate,pd.Qty,pd.BatchNo from purchasehead  `p` inner join purchasedetails `pd` on p.slno=pd.slno where pd.productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex)
    Data1.Refresh
    Me.Data2.RecordSource = "select p.invno,p.invdate,pd.salerate,pd.Qty,pd.Batchno from invoicehead  `p` inner join invoicedetails `pd` on p.invno=pd.invno where pd.productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex)
    Data2.Refresh
    Set rec = db.OpenRecordset("select min(prrate) as minprate from purchasedetails where productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
    If Not IsNull(rec!minprate) Then
        Me.lblpmin.Caption = "Minimum Price:" & rec!minprate
    Else
        Me.lblpmin.Caption = "Minimum Price:0.00"
    End If
    Set rec = db.OpenRecordset("select max(prrate) as maxprate from purchasedetails where productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
    If Not IsNull(rec!maxprate) Then
        Me.lblpmax.Caption = "Maximum Price:" & rec!maxprate
    Else
        Me.lblpmax.Caption = "Maximum Price:0.00"
    End If
    Set rec = db.OpenRecordset("select avg(prrate) as avgprate from purchasedetails where productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
    If Not IsNull(rec!avgprate) Then
        Me.lblpavg.Caption = "Average Price:" & rec!avgprate
    Else
        Me.lblpavg.Caption = "Average Price:0.00"
    End If
    
    Set rec = db.OpenRecordset("select min(salerate) as minprate from invoicedetails where productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
    If Not IsNull(rec!minprate) Then
        Me.lblsmin.Caption = "Minimum Price:" & rec!minprate
    Else
        Me.lblsmin.Caption = "Minimum Price:0.00"
    End If
    Set rec = db.OpenRecordset("select max(salerate) as maxprate from invoicedetails where productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
    If Not IsNull(rec!maxprate) Then
        Me.lblsmax.Caption = "Maximum Price:" & rec!maxprate
    Else
        Me.lblsmax.Caption = "Maximum Price:0.00"
    End If
    Set rec = db.OpenRecordset("select avg(salerate) as avgprate from invoicedetails where productcode=" & Me.cboitemname.ItemData(Me.cboitemname.ListIndex))
    If Not IsNull(rec!avgprate) Then
        Me.lblsavg.Caption = "Average Price:" & rec!avgprate
    Else
        Me.lblsavg.Caption = "Average Price:0.00"
    End If
    
    
End If
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
Me.Data1.databasename = db.Name
Me.Data2.databasename = db.Name
Me.Data3.databasename = db.Name
Set rec = db.OpenRecordset("select * from stock")
    While Not rec.EOF
        Me.cboitemname.AddItem rec("itemname")
        Me.cboitemname.ItemData(Me.cboitemname.NewIndex) = rec("productcode")
        rec.MoveNext
    Wend
    If Me.cboitemname.ListCount > 0 Then
        Me.cboitemname.ListIndex = 0
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

