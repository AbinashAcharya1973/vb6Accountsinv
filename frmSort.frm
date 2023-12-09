VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSort 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort No List"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmSort.frx":0000
         Height          =   5535
         Left            =   120
         OleObjectBlob   =   "frmSort.frx":0014
         TabIndex        =   1
         Top             =   120
         Width           =   2295
      End
   End
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
      RecordSource    =   "Sort_Quality"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
If KeyCode = 13 Then
If formname = "Invoice" Then
frmInvoice.txtSortNo.Text = Me.DBGrid1.Columns(0)
End If
If formname = "PReturn" Then
frmpurchasereturn.txtSortNo.Text = Me.DBGrid1.Columns(0)
End If
If formid = 3 Then
frmStockin.txtQuality.Text = Me.DBGrid1.Columns(0)
End If
If formid = 4 Then
frmSalesReturn.txtSortNo.Text = Me.DBGrid1.Columns(0)
End If
Unload Me
End If
End Sub
Private Sub Form_Load()
Me.Top = 1575
Me.Left = 375
Data1.databasename = App.Path & "\cuts.mdb"
If formname = "invoice" Then
Data1.RecordSource = "select * from Sort_Quality where SortNo like '" & frmInvoice.txtSortNo.Text & "*'"
Data1.Refresh
End If
If formname = "PReturn" Then
Data1.RecordSource = "select * from Sort_Quality where SortNo like '" & frmpurchasereturn.txtSortNo.Text & "*'"
Data1.Refresh
End If
If formid = 3 Then
Data1.RecordSource = "select * from Sort_Quality where SortNo like '" & frmStockin.txtQuality.Text & "*'"
Data1.Refresh
End If
If formid = 4 Then
Data1.RecordSource = "select * from Sort_Quality where SortNo like '" & frmSalesReturn.txtSortNo.Text & "*'"
Data1.Refresh
End If
End Sub
