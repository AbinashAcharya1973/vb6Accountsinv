VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmstock 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock View"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmstock.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   10755
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
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "StockDetails"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmstock.frx":D4E3
      Height          =   3555
      Left            =   120
      OleObjectBlob   =   "frmstock.frx":D4F7
      TabIndex        =   9
      Top             =   4200
      Width           =   10395
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
      Left            =   1440
      TabIndex        =   8
      Top             =   180
      Width           =   4155
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3720
      Width           =   1335
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3360
      Top             =   3420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Stock"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmstock.frx":F92E
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "frmstock.frx":F942
      TabIndex        =   2
      Top             =   600
      Width           =   10455
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3780
      Width           =   1575
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
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
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
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432

Private Sub cmdprint_Click()
Me.CrystalReport1.ReportFileName = App.Path & "\itemwisestock.rpt"
Me.CrystalReport1.PrintReport
End Sub

Private Sub Command1_Click()
Me.Data1.RecordSource = "select * from Stock order by ItemName"
Me.Data1.Refresh
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
Data2.RecordSource = "SELECT * FROM STOCKDETAILS WHERE PRODUCTCODE=" & Me.DBGrid1.Columns(0)
Data2.Refresh
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Data2.databasename = dbname
    
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
Me.Data1.RecordSource = "select * from Stock where ItemName Like '" & Me.txtitemname.Text & "*'"
Me.Data1.Refresh
End Sub
