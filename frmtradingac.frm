VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmtradingac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trading Account"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9990
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmtradingac.frx":0000
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frmtradingac.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   9855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TradingAc"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmtradingac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As DAO.Recordset

Private Sub Form_Load()
Dim opstock, clstock, purchase, sale, preturn, salereturn
Set rec1 = Db.OpenRecordset("select * from sortmaster")
While Not rec1.EOF
    opstock = opstock + (rec1("purchaserate") * rec1("ostock"))
    rec1.MoveNext
Wend
Set rec1 = Db.OpenRecordset("select * from ledgermaster where accname like 'Purchase*'")
If Not rec1.EOF Then
    Set rec2 = Db.OpenRecordset("select sum(dr) as dr_total from ledgertran where accid=" & rec1("accid"))
    If Not IsNull(rec2!dr_total) Then
        purchase = rec1!dr_total
    End If
    Set rec2 = Db.OpenRecordset("select sum(cr) as cr_total from ledgertran where accid=" & rec1("accid"))
    If Not IsNull(rec2!cr_total) Then
        preturn = rec1!cr_total
    End If
End If
Set rec1 = Db.OpenRecordset("select * from ledgermaster where accname like 'Sale*'")
If Not rec1.EOF Then
    Set rec2 = Db.OpenRecordset("select sum(dr) as dr_total from ledgertran where accid=" & rec1("accid"))
    If Not IsNull(rec2!dr_total) Then
        saleretun = rec1!dr_total
    End If
    Set rec2 = Db.OpenRecordset("select sum(cr) as cr_total from ledgertran where accid=" & rec1("accid"))
    If Not IsNull(rec2!cr_total) Then
        sale = rec1!cr_total
    End If
End If

End Sub
