VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmInvoiceEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Edit [Sales Register]"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmInvoiceEdit.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   10095
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12495
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmInvoiceEdit.frx":D4E3
         Height          =   2775
         Left            =   120
         OleObjectBlob   =   "frmInvoiceEdit.frx":D4F7
         TabIndex        =   5
         Top             =   120
         Width           =   9735
      End
   End
   Begin VB.ComboBox cboInvType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   12495
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmInvoiceEdit.frx":10516
         Height          =   3495
         Left            =   120
         OleObjectBlob   =   "frmInvoiceEdit.frx":1052A
         TabIndex        =   2
         Top             =   120
         Width           =   9735
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "InvoiceHead"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\BISINABAR_FMCG\FMCG\DATA\2011-2012\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "InvoiceDetails"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmInvoiceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432
Private Sub cboInvType_Change()
    If Me.CboInvType.Text = "TAX" Then
        Data1.RecordSource = "select * from invoicehead where invtype='TAX' order by InvNo asc "
        Data1.Refresh
    End If
    If Me.CboInvType.Text = "RETAIL" Then
        Data1.RecordSource = "select * from invoicehead where invtype='RETAIL' order by InvNo Asc"
        Data1.Refresh
    End If
    Data2.RecordSource = "select * from invoicedetails where invtype='" & Me.CboInvType.Text & "'"
    Data2.Refresh
End Sub
Private Sub cboInvType_Click()
    cboInvType_Change
End Sub
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Data2.RecordSource = "select * from invoicedetails where InvNo=" & Me.DBGrid1.Columns(0) & " and Invtype='" & Me.DBGrid1.Columns(16) & "'"
    Data2.Refresh
End Sub
Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        ans = MsgBox("Delete This?", vbYesNo)
        If ans = 6 Then
            Me.DBGrid2.AllowDelete = True
        End If
    Else
        Me.DBGrid2.AllowDelete = False
    End If
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Data2.databasename = dbname
    Me.CboInvType.AddItem ("TAX")
    Me.CboInvType.AddItem ("RETAIL")
    Me.CboInvType.ListIndex = 0
    Data2.RecordSource = "select * from Invoicedetails where InvType='" & Me.CboInvType.Text & "'"
    Data2.Refresh
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

