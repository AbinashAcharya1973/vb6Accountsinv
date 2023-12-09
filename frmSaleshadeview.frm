VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSaleshadeview 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Note Shades"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SaleShade"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmSaleshadeview.frx":0000
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "frmSaleshadeview.frx":0014
         TabIndex        =   1
         Top             =   120
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmSaleshadeview"
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
Me.Top = 4000
Me.Left = 2000
Data1.RecordSource = "select * from SaleShade where Salenoteno='" & frmSalenoteview.DBGrid2.Columns(0) & "' and Quality='" & frmSalenoteview.DBGrid2.Columns(1) & "'"
Data1.Refresh
End Sub
