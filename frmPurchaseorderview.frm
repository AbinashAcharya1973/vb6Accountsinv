VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPurchaseorderview 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order View"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11175
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      RecordSource    =   "OrderDetails"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1140
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
      RecordSource    =   "OrderHead"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   10935
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmPurchaseorderview.frx":0000
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "frmPurchaseorderview.frx":0014
         TabIndex        =   3
         Top             =   120
         Width           =   10695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmPurchaseorderview.frx":1097
         Height          =   2775
         Left            =   120
         OleObjectBlob   =   "frmPurchaseorderview.frx":10AB
         TabIndex        =   2
         Top             =   120
         Width           =   10695
      End
   End
End
Attribute VB_Name = "frmPurchaseorderview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
Data2.RecordSource = "select * from OrderDetails where OrderNo=" & Me.DBGrid1.Columns(0)
Data2.Refresh
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
frmOrderShadeview.Show vbModal
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub
