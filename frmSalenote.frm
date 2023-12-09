VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalenote 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Note "
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9345
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   9135
      Begin VB.CommandButton cmdPrint 
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
         Height          =   375
         Left            =   7320
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   5880
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtTotalamount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtTotalqty 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Total Amount"
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
         Left            =   2760
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Total Qty"
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
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   9135
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmSalenote.frx":0000
         Height          =   2655
         Left            =   120
         OleObjectBlob   =   "frmSalenote.frx":0014
         TabIndex        =   34
         Top             =   120
         Width           =   8895
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   9135
      Begin VB.TextBox txtsl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   41
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtDp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtShadeQty 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboUnitcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtQuality 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Remarks"
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Del.Prd"
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
         Left            =   6000
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Amount"
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
         Left            =   3600
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Shades"
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
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Rate"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "U.Code"
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
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Quality"
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
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   9135
      Begin MSMask.MaskEdBox txtorderdate 
         Height          =   315
         Left            =   3480
         TabIndex        =   42
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.TextBox txtPageno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtPaymode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtSaleDate 
         Height          =   315
         Left            =   7080
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.TextBox txtSalenoteno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtorderno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtsupplier 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtbuyeerno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtbuyeer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label8 
         Caption         =   "Page No"
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
         Left            =   6000
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Pay Mode"
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
         Left            =   6000
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Date"
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
         Left            =   6000
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Sale Note No"
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
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Order No"
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
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Supplier"
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
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Buyeers No"
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
         Left            =   6000
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Buyeers Name"
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Temp_Salenote"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmSalenote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset

Private Sub cboUnitcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtRate.SetFocus
End If
End Sub
Private Sub cmdSave_Click()
ans = MsgBox("seve This?", vbYesNo)
If ans = 6 Then
Set rec1 = Db.OpenRecordset("select * from SalenoteHead where Salenoteno='" & Me.txtSalenoteno.Text & "'")
If rec1.EOF Then
    Db.Execute ("insert into SalenoteHead(BuyerNo,BuyeersName,Supplier_Place,Orderno,Orderdate,Salenoteno,SalenoteDate,Pageno,Paymentmode,TotalQty,TotalAmount) values('" & Me.txtbuyeerno.Text & "','" & Me.txtbuyeer.Text & "','" & Me.txtsupplier.Text & "','" & Me.txtorderno.Text & "','" & Me.txtorderdate.Text & "','" & Me.txtSalenoteno.Text & "','" & Me.txtSaleDate.Text & "'," & Me.txtPageno.Text & ",'" & Me.txtPaymode.Text & "'," & Me.txtTotalqty.Text & "," & Me.txtTotalamount.Text & ")")
    Set rec2 = Db.OpenRecordset("select * from Temp_Salenote")
    If Not rec2.EOF Then
    While Not rec2.EOF
    Db.Execute ("insert into Salenote_Details(Salenoteno,Quality,Uitcode,Qty,Rate,Amount) values('" & Me.txtSalenoteno.Text & "','" & rec2("Quality") & "','" & rec2("UnitCode") & "'," & rec2("Qty") & "," & rec2("Rate") & "," & rec2("Amount") & ")")
    rec2.MoveNext
    Wend
    End If
    Set rec2 = Db.OpenRecordset("Select * from Temp_saleshade")
    While Not rec2.EOF
    Db.Execute ("insert into SaleShade(Salenoteno,Quality,UnitCode,Shades,Units,Rate) values('" & rec2("SaleNote_No") & "','" & rec2("Quality") & "','" & rec2("UnitCode") & "','" & rec2("Shades") & "'," & rec2("Units") & "," & rec2("Rate") & ")")
        '-----------------Order Note -----Update ---------------
        Set rec3 = Db.OpenRecordset("select * from OrderShade where OrderNo='" & Me.txtorderno.Text & "' and Quality='" & rec2("Quality") & "' and Shades='" & rec2("Shades") & "' and rate=" & rec2("Rate"))
        If Not rec3.EOF Then
        Db.Execute ("update OrderShade set Status='Yes' where OrderNo='" & Me.txtorderno.Text & "' and Quality='" & rec2("Quality") & "' and Shades='" & rec2("Shades") & "' and rate=" & rec2("Rate"))
        End If
    rec2.MoveNext
    Wend
Else
    Db.Execute ("update SalenoteHead set TotalQty=TotalQty+" & Me.txtTotalqty.Text & ",TotalAmount=TotalAmount+" & Me.txtTotalamount.Text & " where Salenoteno='" & Me.txtSalenoteno.Text & "'")
    
    Set rec2 = Db.OpenRecordset("select * from Temp_Salenote")
    If Not rec2.EOF Then
    While Not rec2.EOF
    Db.Execute ("insert into Salenote_Details(Salenoteno,Quality,Uitcode,Qty,Rate,Amount) values('" & Me.txtSalenoteno.Text & "','" & rec2("Quality") & "','" & rec2("UnitCode") & "'," & rec2("Qty") & "," & rec2("Rate") & "," & rec2("Amount") & ")")
    rec2.MoveNext
    Wend
    End If
    Set rec2 = Db.OpenRecordset("Select * from Temp_saleshade")
    While Not rec2.EOF
    Db.Execute ("insert into SaleShade(Salenoteno,Quality,UnitCode,Shades,Units,Rate) values('" & rec2("SaleNote_No") & "','" & rec2("Quality") & "','" & rec2("UnitCode") & "','" & rec2("Shades") & "'," & rec2("Units") & "," & rec2("Rate") & ")")
    rec2.MoveNext
    Wend


End If
Db.Execute ("delete * from Temp_saleshade")
Db.Execute ("delete * from Temp_Salenote")
Data1.Refresh
Me.txtSalenoteno.Text = ""
Me.txtTotalqty.Text = 0

End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
formid = 2
Me.txtSaleDate.Text = Format(Date, "dd") & "/" & Format(Date, "mm") & "/" & Format(Date, "YYYY")
Set rec1 = Db.OpenRecordset("select * from unitcode")
While Not rec1.EOF
Me.cboUnitcode.AddItem (rec1("UnitName"))
rec1.MoveNext
Wend
Me.cboUnitcode.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Db.Execute ("Delete * from Temp_Salenote")
End Sub

Private Sub txtAmount_GotFocus()
Me.txtAmount.SelStart = 0
Me.txtAmount.SelLength = Len(Me.txtAmount.Text)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtDp.SetFocus
End If
End Sub

Private Sub txtbuyeer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtPageno.SetFocus
End If
End Sub

Private Sub txtbuyeerno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtsupplier.SetFocus
End If
End Sub

Private Sub txtDp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtRemarks.SetFocus
End If
End Sub

Private Sub txtorderno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtbuyeerno.SetFocus
End If
End Sub

Private Sub txtPageno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtQuality.SetFocus
End If
End Sub

Private Sub txtPaymode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtbuyeer.SetFocus
End If
End Sub

Private Sub txtQuality_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cboUnitcode.SetFocus
End If
End Sub

Private Sub txtRate_GotFocus()
Me.txtRate.SelStart = 0
Me.txtRate.SelLength = Len(Me.txtRate.Text)
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

Me.txtShadeQty.SetFocus
End If
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtsl.SetFocus
End If
End Sub

Private Sub txtSaleDate_GotFocus()
Me.txtSaleDate.SelStart = 0
Me.txtSaleDate.SelLength = Len(Me.txtSaleDate.Text)
End Sub

Private Sub txtSaleDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtorderno.SetFocus
End If
End Sub

Private Sub txtSalenoteno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtSaleDate.SetFocus
End If
End Sub

Private Sub txtShadeQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
frmSaleNoteShade.Show vbModal
End If
If KeyCode = 13 Then
If Me.txtShadeQty.Text <> "" Then
Me.txtAmount.Text = Format(Val(Me.txtRate.Text) * Val(Me.txtShadeQty.Text), "######0.00")
Me.txtAmount.SetFocus
End If
End If
End Sub

Private Sub txtsl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Set rec1 = Db.OpenRecordset("select * from Temp_Salenote where Quality='" & Me.txtQuality.Text & "'")
    If rec1.EOF Then
    Db.Execute ("insert into Temp_Salenote (Quality,UnitCode,Qty,Rate,Amount,DP,Remarks,Sl) values('" & Me.txtQuality.Text & "','" & Me.cboUnitcode.Text & "'," & Me.txtShadeQty.Text & "," & Me.txtRate.Text & "," & Me.txtAmount.Text & ",'" & Me.txtDp.Text & "','" & Me.txtRemarks.Text & "'," & Me.txtsl.Text & ")")
    Else
    Db.Execute ("update Temp_Salenote set Qty=Qty+" & Me.txtShadeQty.Text & ",Amount=Amount+" & (Val(Me.txtRate.Text) * Val(Me.txtShadeQty.Text)) & " where Quality='" & Me.txtQuality.Text & "'")
    End If
    Data1.Refresh
Me.txtTotalqty.Text = Format(Val(Me.txtTotalqty.Text) + Val(Me.txtShadeQty.Text), "######0.00")
Me.txtTotalamount.Text = Format(Val(Me.txtTotalamount.Text) + Val(Me.txtAmount.Text), "######0.00")
Me.txtQuality.Text = ""
Me.txtRate.Text = "0.00"
Me.txtAmount.Text = "0.00"
Me.txtDp.Text = ""
Me.txtRemarks.Text = ""
Me.txtsl.Text = ""
Me.txtShadeQty.Text = ""
Me.txtQuality.SetFocus
End If
End Sub

Private Sub txtsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtPaymode.SetFocus
End If
End Sub
