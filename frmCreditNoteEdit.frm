VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCreditNoteEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Note Or Sales Returns Edit"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10095
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   9855
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmCreditNoteEdit.frx":0000
         Height          =   3375
         Left            =   120
         OleObjectBlob   =   "frmCreditNoteEdit.frx":0014
         TabIndex        =   3
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmCreditNoteEdit.frx":1F8F
         Height          =   3135
         Left            =   0
         OleObjectBlob   =   "frmCreditNoteEdit.frx":1FA3
         TabIndex        =   2
         Top             =   120
         Width           =   9615
      End
   End
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
      RecordSource    =   "Salesreturndetails"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1260
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
      RecordSource    =   "Salesreturnhead"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmCreditNoteEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Data2.RecordSource = "select * from Salesreturndetails where InvNo=" & Me.DBGrid1.Columns(0) & ""
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
