VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmstockinshade 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock In Item/Shades"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6975
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
      RecordSource    =   "Temp_Stockinshade"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   6735
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmstockinshade.frx":0000
         Height          =   2295
         Left            =   120
         OleObjectBlob   =   "frmstockinshade.frx":0014
         TabIndex        =   14
         Top             =   120
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtslno 
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
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "SlNo."
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
         TabIndex        =   12
         Top             =   240
         Width           =   855
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
         TabIndex        =   11
         Top             =   240
         Width           =   735
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
         TabIndex        =   9
         Top             =   720
         Width           =   375
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
         TabIndex        =   8
         Top             =   240
         Width           =   375
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
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmstockinshade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset
Private Sub Form_Load()
Me.Top = 2880
Me.Left = 7092
Me.txtslno.Text = frmStockin.txtslno.Text
Me.txtQuality.Text = frmStockin.txtQuality.Text
Me.txtUnitCode.Text = frmStockin.cboUnitcode.Text
End Sub

Private Sub txtShade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If formid = 3 Then
    Set rec1 = Db.OpenRecordset("select Qty from UnitCode where UnitName='" & Me.txtUnitCode.Text & "'")
    If Not rec1.EOF Then
    'set rec2=db.OpenRecordset("select sum(Units) from Temp_Ordershade where
    frmStockin.txtShadeQty.Text = Val(Me.txtTotal.Text) * Val(rec1("Qty"))
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
Db.Execute ("insert into Temp_Stockinshade(Slno,Quality,UnitCode,Shades,Units,Rate) values(" & Me.txtslno.Text & ",'" & Me.txtQuality.Text & "','" & Me.txtUnitCode.Text & "','" & Me.txtShade.Text & "'," & Me.txtUnits.Text & "," & frmStockin.txtRate.Text & ")")
Data1.Refresh
Me.txtTotal = Val(Me.txtTotal) + Val(Me.txtUnits.Text)
Me.txtShade.Text = ""
Me.txtUnits.Text = ""
Me.txtShade.SetFocus
End If
End If
End Sub
