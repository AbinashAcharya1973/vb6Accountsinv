VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmTemp_TradingAccount 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temp Trading Account"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8190
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmTemp_TradingAccount.frx":0000
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "frmTemp_TradingAccount.frx":0014
         TabIndex        =   1
         Top             =   120
         Width           =   7695
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Temp_TradingAcc"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
   End
End
Attribute VB_Name = "frmTemp_TradingAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset
Private Sub Form_Load()
Db.Execute ("delete * from Temp_TradingAcc")
Me.Top = 0
Me.Left = 0
Dim opstock, clstock, purchase, sale, preturn, salereturn
Set rec1 = Db.OpenRecordset("select * from Stock")
While Not rec1.EOF
    opstock = opstock + (rec1("PurchaseRate") * rec1("OStock"))
    clstock = clstock + (rec1("PurchaseRate") * rec1("Qty"))
    rec1.MoveNext
Wend
Db.Execute ("insert into Temp_TradingAcc (Particulars,Dr) values('To Opening Stock'," & opstock & ")")
Db.Execute ("insert into Temp_TradingAcc (Particulars,Cr) values('By Closing Stock'," & clstock & ")")

Set rec1 = Db.OpenRecordset("select * from ledgermaster where accname like 'Purchase*'")
If Not rec1.EOF Then
    Set rec2 = Db.OpenRecordset("select * from LedgerTran where AccId=" & rec1("AccId") & " and SlNo=(Select max(Slno) from LedgerTran where AccId=" & rec1("AccId") & ")")
    If Not rec2.EOF Then
    Db.Execute ("insert into Temp_TradingAcc (Particulars,Dr) values('" & rec1("AccName") & "'," & rec2("Balance") & ")")
    End If
End If
Set rec1 = Db.OpenRecordset("select * from ledgermaster where accname like 'Sale*'")
If Not rec1.EOF Then
    Set rec2 = Db.OpenRecordset("select * from LedgerTran where AccId=" & rec1("AccId") & " and SlNo=(Select max(Slno) from LedgerTran where AccId=" & rec1("AccId") & ")")
    If Not rec2.EOF Then
    Db.Execute ("insert into Temp_TradingAcc (Particulars,Cr) values('" & rec1("AccName") & "'," & rec2("Balance") & ")")
    End If
End If
Set rec1 = Db.OpenRecordset("select Sum(Dr) as Total_Dr from Temp_TradingAcc")
If Not IsNull(rec1!Total_Dr) Then
Me.Label1.Caption = Format(rec1!Total_Dr, "#########0.00")
End If
Set rec1 = Db.OpenRecordset("select Sum(Cr) as Total_Cr from Temp_TradingAcc")
If Not IsNull(rec1!Total_Cr) Then
Me.Label2.Caption = Format(rec1!Total_Cr, "#########0.00")
End If

End Sub
