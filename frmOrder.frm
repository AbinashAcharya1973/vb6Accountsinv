VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOrder 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9600
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "temp_order"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   6600
      Width           =   9375
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtTotalAmount 
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
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtTotalQty 
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
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label9 
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
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   9375
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmOrder.frx":0000
         Height          =   3735
         Left            =   120
         OleObjectBlob   =   "frmOrder.frx":0014
         TabIndex        =   25
         Top             =   120
         Width           =   9135
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   9375
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
         Left            =   8520
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtQty 
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
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtRate 
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
         Left            =   5040
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboUnitCode 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   975
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
         Left            =   1200
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "D.P"
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
         Left            =   8160
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
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
         Left            =   6120
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "U.C"
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
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Quality "
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
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9375
      Begin VB.TextBox txtBuyerno 
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
         Left            =   7560
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboCompany 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtOrderdate 
         Height          =   315
         Left            =   7560
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.TextBox txtOrderno 
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
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Buyer No"
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
         Left            =   6600
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Company"
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
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Left            =   6600
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As DAO.Recordset, REC2 As Recordset
Private Sub cboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtBuyerno.SetFocus
End If
End Sub

Private Sub cboUnitcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtRate.SetFocus
End If
End Sub

Private Sub CmdSave_Click()
ans = MsgBox("save This?", vbYesNo)
If ans = 6 Then
Set REC1 = Db.OpenRecordset("select * from OrderHead where OrderNo=" & Me.txtOrderno.Text)
If REC1.EOF Then
    Db.Execute ("insert into OrderHead(OrderNo,OrderDate,Company,TotalQty,TotalValue,BuyerNo) values(" & Me.txtOrderno.Text & ",'" & Me.txtOrderdate.Text & "','" & Me.cboCompany.Text & "'," & Me.txtTotalQty.Text & "," & Me.txtTotalAmount.Text & ",'" & Me.txtBuyerno.Text & "')")
    Set REC2 = Db.OpenRecordset("select * from temp_order")
    If Not REC2.EOF Then
        While Not REC2.EOF
        Db.Execute ("insert into OrderDetails(OrderNo,Quality,UnitCode,Qty,Rate,DeliveryPeriod) values(" & Me.txtOrderno.Text & ",'" & REC2("Quality") & "','" & REC2("UnitCode") & "'," & REC2("Qty") & "," & REC2("Rate") & ",'" & REC2("DeliveryPeriod") & "')")
        REC2.MoveNext
        Wend
    End If
    Set REC2 = Db.OpenRecordset("select * from Temp_Ordershade")
    If Not REC2.EOF Then
    While Not REC2.EOF
    Db.Execute ("insert into OrderShade(OrderNo,Quality,UnitCode,Shades,Units,Rate,Status) values(" & REC2("OrderNo") & ",'" & REC2("Quality") & "','" & REC2("UnitCode") & "','" & REC2("Shades") & "'," & REC2("Units") & "," & REC2("Rate") & ",'No')")
    REC2.MoveNext
    Wend
    End If
Else
Db.Execute ("update OrderHead set TotalQty=TotalQty+" & Val(Me.txtTotalQty.Text) & ",TotalValue=TotalValue+" & Me.txtTotalAmount.Text) & " where OrderNo=" & Me.txtOrderno.Text & ")"
    Set REC2 = Db.OpenRecordset("select * from temp_order")
    If Not REC2.EOF Then
        While Not REC2.EOF
        Db.Execute ("insert into OrderDetails(OrderNo,Quality,UnitCode,Qty,Rate,DeliveryPeriod) values(" & Me.txtOrderno.Text & ",'" & REC2("Quality") & "','" & REC2("UnitCode") & "'," & REC2("Qty") & "," & REC2("Rate") & ",'" & REC2("DeliveryPeriod") & "')")
        REC2.MoveNext
        Wend
    End If
    Set REC2 = Db.OpenRecordset("select * from Temp_Ordershade")
    If Not REC2.EOF Then
    While Not REC2.EOF
    Db.Execute ("insert into OrderShade(OrderNo,Quality,UnitCode,Shades,Units,Rate,Status) values(" & REC2("OrderNo") & ",'" & REC2("Quality") & "','" & REC2("UnitCode") & "','" & REC2("Shades") & "'," & REC2("Units") & "," & REC2("Rate") & ",'No')")
    REC2.MoveNext
    Wend
    End If

End If
Db.Execute ("delete * from temp_order")
Db.Execute ("delete * from Temp_Ordershade")
Me.txtOrderno.Text = Val(Me.txtOrderno.Text) + 1
Me.txtBuyerno.Text = ""
Me.txtTotalAmount.Text = "0.00"
Me.txtTotalQty.Text = "0.00"




End If
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
Me.txtTotalAmount.Text = Format(Val(Me.txtTotalAmount.Text) - (Val(Me.DBGrid1.Columns(2)) * Val(Me.DBGrid1.Columns(3))), "######0.00")
Me.txtTotalQty.Text = Format(Val(Me.txtTotalQty.Text) - Val(Me.DBGrid1.Columns(2)), "######0.00")
Db.Execute ("delete * from Temp_Ordershade where Quality='" & Me.DBGrid1.Columns(1) & "'")
End Sub
Private Sub Form_Load()
formid = 1
Me.Top = 0
Me.Left = 0
Data1.databasename = App.Path & "\cuts.mdb"
Set REC1 = Db.OpenRecordset("Select max(OrderNo) as max_no from OrderHead")
If Not IsNull(REC1!max_no) Then
Me.txtOrderno.Text = REC1!max_no + 1
Else
Me.txtOrderno.Text = 1000001
End If
Set REC1 = Db.OpenRecordset("select * from UnitCode")
While Not REC1.EOF
Me.cboUnitCode.AddItem (REC1("UnitName"))
REC1.MoveNext
Wend
If Me.cboUnitCode.ListCount > 0 Then
Me.cboUnitCode.ListIndex = 0
End If
Me.txtOrderdate.Text = Format(Date, "DD") & "/" & Format(Date, "MM") & "/" & Format(Date, "YYYY")
Set REC1 = Db.OpenRecordset("select * from CompanyMaster")
While Not REC1.EOF
Me.cboCompany.AddItem (REC1("Company"))
REC1.MoveNext
Wend
If Me.cboCompany.ListCount > 0 Then
Me.cboCompany.ListIndex = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Db.Execute ("delete * from temp_order")
Db.Execute ("delete * from Temp_Ordershade")
End Sub

Private Sub txtBuyerno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtQuality.SetFocus
End If
End Sub

Private Sub txtDp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.txtDp.Text <> "" Then
    Set REC1 = Db.OpenRecordset("select * from temp_order where OrderNo=" & Me.txtOrderno.Text & " and Quality='" & Me.txtQuality.Text & "'")
        If Not REC1.EOF Then
            Db.Execute ("update temp_order set Qty=Qty+" & Me.txtQty.Text & " where OrderNo=" & Me.txtOrderno.Text & " and Quality='" & Me.txtQuality.Text & "'")
            Else
            Db.Execute ("insert into temp_order(OrderNo,Quality,UnitCode,Qty,Rate,DeliveryPeriod) values('" & Me.txtOrderno.Text & "','" & Me.txtQuality.Text & "','" & Me.cboUnitCode.Text & "'," & Me.txtQty.Text & "," & Me.txtRate.Text & "," & Me.txtDp.Text & ")")
        End If
        Data1.Refresh

End If
Me.txtTotalAmount.Text = Format(Val(Me.txtTotalAmount.Text) + (Val(Me.txtRate.Text) * Val(Me.txtQty.Text)), "######0.00")
Me.txtTotalQty.Text = Format(Val(Me.txtTotalQty.Text) + Val(Me.txtQty.Text), "######0.00")
Me.txtQuality.Text = ""
Me.txtQty.Text = 0
Me.txtDp.Text = 0
Me.txtRate.Text = "0.00"
Me.txtQuality.SetFocus
End If

End Sub
Private Sub txtOrderdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cboCompany.SetFocus
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
frmShades.Show vbModal
End If
If KeyCode = 13 Then
Me.txtDp.SetFocus
End If
End Sub

Private Sub txtQuality_GotFocus()
Me.txtQuality.SelStart = 0
Me.txtQuality.SelLength = Len(Me.txtQuality.Text)
End Sub

Private Sub txtQuality_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cboUnitCode.SetFocus
End If
End Sub

Private Sub txtRate_GotFocus()
Me.txtRate.SelStart = 0
Me.txtRate.SelLength = Len(Me.txtRate.Text)
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtQty.SetFocus
End If
End Sub
