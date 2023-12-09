VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmStockShadeView 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Shade View"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Stock_Shade"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmStockShadeView.frx":0000
         Height          =   3855
         Left            =   120
         OleObjectBlob   =   "frmStockShadeView.frx":0014
         TabIndex        =   1
         Top             =   120
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmStockShadeView"
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
Me.Left = 3050

Data1.RecordSource = "select * from Stock_Shade where Quality='" & frmstock.DBGrid1.Columns(0) & "' and UnitCode='" & frmstock.DBGrid1.Columns(1) & "' and Rate=" & frmstock.DBGrid1.Columns(3) & ""
Data1.Refresh
End Sub
