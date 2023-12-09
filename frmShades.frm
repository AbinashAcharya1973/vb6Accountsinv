VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmShades 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Shades"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
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
      RecordSource    =   "Temp_Ordershade"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   6735
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmShades.frx":0000
         Height          =   2295
         Left            =   120
         OleObjectBlob   =   "frmShades.frx":0014
         TabIndex        =   12
         Top             =   120
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtTotal 
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
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtUnitCode 
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
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtUnits 
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
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtShade 
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
         TabIndex        =   0
         Top             =   720
         Width           =   1695
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtOderNo 
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
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   720
         Width           =   495
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
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Units"
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
         Left            =   2880
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Shade No"
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
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   735
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
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmShades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As DAO.Recordset, REC2 As Recordset
Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'before_unitcode = Me.DBGrid1.Columns(2)
'before_units = Me.DBGrid1.Columns(4)
'before_rate = Me.DBGrid1.Columns(5)
'Set rec1 = Db.OpenRecordset("select Qty from UnitCode where UnitName='" & Me.DBGrid1.Columns(2) & "'")
'If Not rec1.EOF Then
'temp_unit = before_units * rec1("Qty")
'
'Db.Execute ("update temp_order set Qty=Qty-" & Val(temp_unit) & " where Quality='" & Me.DBGrid1.Columns(1) & "'")
'frmOrder.Data1.Refresh
'frmOrder.DBGrid1.Refresh
'frmOrder.txtTotalamount.Text = Format(Val(frmOrder.txtTotalamount.Text) - (Val(temp_unit) * Val(before_rate)), "######0.00")
'frmOrder.txtTotalqty.Text = Format(Val(frmOrder.txtTotalqty.Text) - Val(temp_unit), "######0.00")
'End If
'before_unitcode = 0
'before_units = 0
'before_rate = 0

End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
delete_units = Me.DBGrid1.Columns(4)
delete_rate = Me.DBGrid1.Columns(5)
Set REC1 = Db.OpenRecordset("select Qty from UnitCode where UnitName='" & Me.DBGrid1.Columns(2) & "'")
If Not REC1.EOF Then
temp_unit = delete_units * REC1("Qty")
Db.Execute ("update temp_order set Qty=Qty-" & Val(temp_unit) & " where Quality='" & Me.DBGrid1.Columns(1) & "'")
frmOrder.txtTotalAmount.Text = Format(Val(frmOrder.txtTotalAmount.Text) - (Val(temp_unit) * Val(delete_rate)), "######0.00")
frmOrder.txtTotalqty.Text = Format(Val(frmOrder.txtTotalqty.Text) - Val(temp_unit), "######0.00")
End If
delete_units = 0
delete_rate = 0
frmOrder.Data1.Refresh
frmOrder.DBGrid1.Refresh
End Sub

Private Sub Form_Load()
Me.Top = 2880
Me.Left = 7092
Data1.databasename = App.Path & "\cuts.mdb"
Me.txtOderNo.Text = frmOrder.txtOrderno.Text
Me.txtUnitCode.Text = frmOrder.cboUnitCode.Text
Me.txtQuality.Text = frmOrder.txtQuality.Text
End Sub

Private Sub txtShade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If formid = 1 Then
    Set REC1 = Db.OpenRecordset("select Qty from UnitCode where UnitName='" & Me.txtUnitCode.Text & "'")
    If Not REC1.EOF Then
    'set rec2=db.OpenRecordset("select sum(Units) from Temp_Ordershade where
    frmOrder.txtQty.Text = Val(Me.txtTotal.Text) * Val(REC1("Qty"))
    End If
    Unload Me
    End If
    End If
If KeyCode = 13 Then
Me.txtUnits.SetFocus
End If



End Sub

Private Sub txtUnits_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.txtUnits.Text <> "" Then
Db.Execute ("insert into temp_OrderShade(OrderNo,Quality,UnitCode,Shades,Units,rate) values(" & Me.txtOderNo.Text & ",'" & Me.txtQuality.Text & "','" & Me.txtUnitCode.Text & "','" & Me.txtShade.Text & "'," & Me.txtUnits.Text & "," & frmOrder.txtRate.Text & ")")
Data1.Refresh
Me.txtTotal = Val(Me.txtTotal) + Val(Me.txtUnits.Text)
Me.txtShade.Text = ""
Me.txtUnits.Text = ""
Me.txtShade.SetFocus
End If
End If
End Sub
