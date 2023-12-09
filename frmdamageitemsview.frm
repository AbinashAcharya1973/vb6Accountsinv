VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmdamageitemsview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Damage Items"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmdamageitemsview.frx":0000
   ScaleHeight     =   7425
   ScaleWidth      =   10020
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\BISINABAR_FMCG\FMCG\DATA\2011-2012\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DamageHead"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   9735
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmdamageitemsview.frx":D4E3
         Height          =   3375
         Left            =   120
         OleObjectBlob   =   "frmdamageitemsview.frx":D4F7
         TabIndex        =   3
         Top             =   240
         Width           =   9495
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
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
         RecordSource    =   "DamageDetails"
         Top             =   720
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmdamageitemsview.frx":F5E2
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "frmdamageitemsview.frx":F5F6
         TabIndex        =   2
         Top             =   240
         Width           =   9495
      End
   End
End
Attribute VB_Name = "frmdamageitemsview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Me.Data2.RecordSource = "Select * from DamageDetails where Slno=" & Val(Me.DBGrid1.Columns(0))
    Me.Data2.Refresh
    Me.DBGrid2.Refresh
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub
