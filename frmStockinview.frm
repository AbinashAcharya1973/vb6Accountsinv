VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmStockinview 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock In / Purchase Register"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmStockinview.frx":0000
   ScaleHeight     =   7410
   ScaleWidth      =   10095
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PurchaseDetails"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PurchaseHead"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   9855
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmStockinview.frx":D4E3
         Height          =   3735
         Left            =   120
         OleObjectBlob   =   "frmStockinview.frx":D4F7
         TabIndex        =   3
         Top             =   120
         Width           =   9615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmStockinview.frx":E8A6
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "frmStockinview.frx":E8BA
         TabIndex        =   2
         Top             =   120
         Width           =   9615
      End
   End
End
Attribute VB_Name = "frmStockinview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Data2.RecordSource = "select * from Purchasedetails where Slno=" & Me.DBGrid1.Columns(0) & ""
    Data2.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Data2.databasename = dbname
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub


