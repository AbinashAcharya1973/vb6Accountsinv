VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmStockinshadeview 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock In Shade View"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Stockinshade"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmStockinshadeview.frx":0000
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "frmStockinshadeview.frx":0014
         TabIndex        =   1
         Top             =   120
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmStockinshadeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Top = 1000
Me.Left = 3000
Data1.RecordSource = "select * from Stockinshade where Slno=" & frmStockinview.DBGrid2.Columns(0) & " and Quality='" & frmStockinview.DBGrid2.Columns(1) & "' and UnitCode='" & frmStockinview.DBGrid2.Columns(2) & "'"
Data1.Refresh
End Sub
