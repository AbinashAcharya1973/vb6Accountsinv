VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSaleNoteShade 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Note Shade Entry"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
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
      RecordSource    =   "Temp_saleshade"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   7455
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmSaleNoteShade.frx":0000
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "frmSaleNoteShade.frx":0014
         TabIndex        =   14
         Top             =   120
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7455
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
         Left            =   5520
         TabIndex        =   2
         Top             =   720
         Width           =   1815
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
         Left            =   3600
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtShades 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtUnitcode 
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   975
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4680
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbl34 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "S Note No"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSaleNoteShade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
delete_rate = Me.DBGrid1.Columns(5)
delete_unit = Me.DBGrid1.Columns(4)
delete_uc = Me.DBGrid1.Columns(2)
Set rec1 = Db.OpenRecordset("select Qty from UnitCode where UnitName='" & delete_uc & "'")
If Not rec1.EOF Then
temp_unit = delete_unit * rec1("Qty")
Db.Execute ("update Temp_Salenote set Qty=Qty-" & temp_unit & ",Amount=Amount-" & Val((temp_unit * delete_rate)) & " where Quality='" & Me.DBGrid1.Columns(1) & "'")
frmSalenote.txtTotalqty.Text = Format(Val(frmSalenote.txtTotalqty.Text) - temp_unit, "######0.00")
frmSalenote.txtTotalamount.Text = Format(Val(frmSalenote.txtTotalamount.Text) - (temp_unit * delete_rate), "######0.00")
frmSalenote.Data1.Refresh
End If
delete_rate = 0
delete_unit = ""
delete_uc = ""
temp_unit = 0

End Sub
Private Sub Form_Load()
Me.Top = 3900
Me.Left = 1600
Me.txtSalenoteno.Text = frmSalenote.txtSalenoteno.Text
Me.txtQuality.Text = frmSalenote.txtQuality.Text
Me.txtUnitCode.Text = frmSalenote.cboUnitcode.Text

End Sub

Private Sub txtShades_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If formid = 2 Then
    Set rec1 = Db.OpenRecordset("select Qty from UnitCode where UnitName='" & Me.txtUnitCode.Text & "'")
    If Not rec1.EOF Then
        frmSalenote.txtShadeQty.Text = Format(Val(Me.txtTotal.Text) * Val(rec1("Qty")), "######0.00")
    End If
    Unload Me
    End If
    End If
    If KeyCode = 13 Then
    If Me.txtShades.Text <> "" Then
    Me.txtUnits.SetFocus
    End If
    End If
End Sub
Private Sub txtUnits_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.txtUnits.Text <> "" And Me.txtShades.Text <> "" Then
Db.Execute ("insert into Temp_saleshade(SaleNote_No,Quality,UnitCode,Shades,Units,Rate) values('" & Me.txtSalenoteno.Text & "','" & Me.txtQuality.Text & "','" & Me.txtUnitCode.Text & "','" & Me.txtShades.Text & "'," & Me.txtUnits.Text & "," & frmSalenote.txtRate.Text & ")")
Data1.Refresh
Me.txtTotal.Text = Val(Me.txtTotal.Text) + Val(Me.txtUnits.Text)
Me.txtShades.Text = ""
Me.txtUnits.Text = ""
Me.txtShades.SetFocus

End If
End If
End Sub
