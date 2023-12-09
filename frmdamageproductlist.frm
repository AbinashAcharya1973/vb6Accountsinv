VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmdamageproductlist 
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
      Bindings        =   "frmdamageproductlist.frx":0000
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "frmdamageproductlist.frx":0014
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
      Left            =   1380
      TabIndex        =   0
      Top             =   120
      Width           =   9855
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
      RecordSource    =   "DamageStock"
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
Attribute VB_Name = "frmdamageproductlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 13 Then
        
        
            frmdamagereturn.cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
            frmdamagereturn.txtproductcode.Text = Trim(Me.DBGrid1.Columns(0))
            frmdamagereturn.cbobatch.Text = Trim(Me.DBGrid1.Columns(4))
            frmdamagereturn.txtmfgdate.Text = Trim(Me.DBGrid1.Columns(2))
            frmdamagereturn.txtexpdate.Text = Trim(Me.DBGrid1.Columns(3))
            frmdamagereturn.txtmrp.Text = Me.DBGrid1.Columns(7)
            frmdamagereturn.TxtVat.Text = Me.DBGrid1.Columns(10)
            'frmpurchasereturn.cbounit.Text = Me.DBGrid1.Columns(7)
            frmdamagereturn.txttaxtype.Text = "SALES"
            Unload Me
            frmdamagereturn.txtpack.SetFocus
        
        
    End If
End Sub

Private Sub DBGrid2_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 3000
    Me.Left = 1740
    Me.Data1.databasename = dbname
    'Me.Data1.RecordSource = "select * from Stock where itemname like '" & SEARCHWORD & "*' order by itemname"
    'Me.Data1.RecordSource = "Select Stock.ProductCode,Stock.itemname,Stock.Lose,Stock.MRP,Stock.SaleRate,Stock.Qty,ItemMaster.Tax,Stock.UniyType,ItemMaster.tax_type from Stock inner Join ItemMaster on Stock.ProductCode=ItemMaster.ProductCode where Stock.ItemName Like '" & SEARCHWORD & "*'"
    Me.Data1.RecordSource = "select * from damagestock where Itemname like '" & SEARCHWORD & "*' order by Itemname"
    Me.Data1.Refresh
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtproductname_Change()
    'Me.Data1.RecordSource = "Select Stock.ProductCode,Stock.itemname,Stock.Lose,Stock.MRP,Stock.SaleRate,Stock.Qty,ItemMaster.Tax,Stock.UniyType,ItemMaster.tax_type from Stock inner Join ItemMaster on Stock.ProductCode=ItemMaster.ProductCode where Stock.ItemName Like '" & Me.txtproductname.Text & "*'"
    Me.Data1.RecordSource = "select * from damagestock where Itemname like '" & Me.txtproductname.Text & "*' order by Itemname"
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
