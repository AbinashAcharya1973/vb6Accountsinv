VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmstockpoints 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create StockPoints/Branch"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6135
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4350
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "StockPoints"
      Top             =   630
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmstockpoints.frx":0000
      Height          =   2805
      Left            =   60
      OleObjectBlob   =   "frmstockpoints.frx":0014
      TabIndex        =   2
      Top             =   780
      Width           =   5985
   End
   Begin VB.TextBox txtstockpoint 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2580
      TabIndex        =   1
      Top             =   240
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "New StockPoint/Branch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   270
      Width           =   2415
   End
End
Attribute VB_Name = "frmstockpoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
ans = MsgBox("Do you want to Delete the Branch?", vbYesNo)
If ans = 6 Then
    db.Execute ("drop table " & Me.DBGrid1.Columns(1))
    db.Execute ("drop table " & Me.DBGrid1.Columns(1) & "Details")
Else
    Cancel = 1
End If
End Sub

Private Sub Form_Load()
Me.Data1.databasename = db.Name
End Sub

Private Sub txtstockpoint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ans = MsgBox("Do you want to Create New Branch?", vbYesNo)
    If ans = 6 Then
        Set rec1 = db.OpenRecordset("select * from stockpoints where StockPoint='" & Me.txtstockpoint.Text & "'")
        If rec1.EOF Then
            db.Execute ("insert into stockpoints (stockpoint) values('" & Me.txtstockpoint.Text & "')")
            db.Execute ("select * into " & Me.txtstockpoint.Text & " from stock")
            db.Execute ("select * into " & Me.txtstockpoint.Text & "Details from stockdetails")
            db.Execute ("update " & Me.txtstockpoint.Text & " set qty=0")
            db.Execute ("delete * from " & Me.txtstockpoint.Text & "Details")
            Me.Data1.Refresh
            Me.DBGrid1.Refresh
        Else
            MsgBox "Branch is already exists in this Name", vbCritical
        End If
    End If
End If
End Sub
