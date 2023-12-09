VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmItemEdit 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Edit / View"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmItemEdit.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6270
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ItemMaster"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmItemEdit.frx":D4E3
         Height          =   5535
         Left            =   120
         OleObjectBlob   =   "frmItemEdit.frx":D4F7
         TabIndex        =   1
         Top             =   120
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BEFOREITEM
Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    ans = MsgBox("Update This?", vbYesNo)
    If ans = 6 Then
        db.Execute ("update PurchaseDetails set Items='" & Trim(Me.DBGrid1.Columns(0)) & "' where Items='" & BEFOREITEM & "'")
        db.Execute ("update invoicedetails set Items='" & Trim(Me.DBGrid1.Columns(0)) & "' where Items='" & BEFOREITEM & "'")
        db.Execute ("UPDATE STOCK SET Items='" & Trim(Me.DBGrid1.Columns(0)) & "' WHERE Items='" & BEFOREITEM & "'")
        Me.DBGrid1.AllowUpdate = False
    End If
End Sub
Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    BEFOREITEM = Trim(Me.DBGrid1.Columns(0))
End Sub
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        Me.DBGrid1.AllowUpdate = True
    End If
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
End Sub
