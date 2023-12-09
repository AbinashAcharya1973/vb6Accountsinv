VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmOrderShadeview 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Shade View"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6945
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "OrderShade"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmOrderShadeview.frx":0000
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "frmOrderShadeview.frx":0014
         TabIndex        =   1
         Top             =   120
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmOrderShadeview"
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
Me.Left = 4000
Data1.databasename = App.Path & "\cuts.mdb"
Data1.RecordSource = "select * from OrderShade where OrderNo=" & frmPurchaseorderview.DBGrid2.Columns(0) & " and Quality='" & frmPurchaseorderview.DBGrid2.Columns(1) & "'"
Data1.Refresh
End Sub
