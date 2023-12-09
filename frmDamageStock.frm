VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDamageStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Damage Stock View"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   13245
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   5175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DamageStock"
      Top             =   3300
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDamageStock.frx":0000
      Height          =   5955
      Left            =   120
      OleObjectBlob   =   "frmDamageStock.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   12915
   End
End
Attribute VB_Name = "frmDamageStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Data1.databasename = db.Name
End Sub

Private Sub txtsearch_Change()
Me.Data1.RecordSource = "select * from damagestock where itemname like '" & Me.txtsearch.Text & "*'"
Me.Data1.Refresh
End Sub
