VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmopstockedit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Or Opening Stock Edit"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmopstockedit.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   12765
   Begin VB.TextBox txtBar 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7500
      TabIndex        =   5
      Top             =   180
      Width           =   3555
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      TabIndex        =   3
      Top             =   180
      Width           =   4455
   End
   Begin VB.CommandButton cmdallitem 
      BackColor       =   &H0000C000&
      Caption         =   "All Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmopstockedit.frx":D4E3
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmopstockedit.frx":D4F7
      TabIndex        =   0
      Top             =   720
      Width           =   12495
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
      RecordSource    =   "ItemMaster"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode Search"
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
      Left            =   5820
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frmopstockedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BEF_QTY, rec1 As DAO.Recordset, rec2 As Recordset, BEF_ITEM, BEF_MRP, BEF_PRATE, BEF_PRDCODE
Private Sub cboitemtype_Click()
    'cboitemtype_Change
End Sub

Private Sub cmdallitem_Click()
    Data1.RecordSource = "SELECT * FROM itemmaster order by item"
    Data1.Refresh

End Sub

Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
'Data1.Refresh
End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        Set rec1 = db.OpenRecordset("select * from itemmaster where item='" & BEF_ITEM & "' AND ProductCode=" & BEF_PRDCODE)
        If Not rec1.EOF Then
            db.Execute ("update itemmaster set item='" & Trim(Me.DBGrid1.Columns(2)) & "' where item='" & BEF_ITEM & "' and productcode=" & BEF_PRDCODE)
        End If

        Set rec1 = db.OpenRecordset("SELECT * FROM stock where itemname='" & BEF_ITEM & "' and productcode=" & BEF_PRDCODE)
        If Not rec1.EOF Then
            db.Execute ("update stock set itemname='" & Trim(Me.DBGrid1.Columns(2)) & "' where itemname='" & BEF_ITEM & "' and productcode=" & BEF_PRDCODE)
        End If

        Set rec1 = db.OpenRecordset("select * from InvoiceDetails where itemname='" & BEF_ITEM & "' and productcode=" & BEF_PRDCODE)
        If Not rec1.EOF Then
            db.Execute ("update InvoiceDetails set Itemname='" & Trim(Me.DBGrid1.Columns(2)) & "' where itemname='" & BEF_ITEM & "' and productcode=" & BEF_PRDCODE)
        End If

        Set rec1 = db.OpenRecordset("select * from PurchaseDetails where ItemName='" & BEF_ITEM & "'")
        If Not rec1.EOF Then
            db.Execute ("update PurchaseDetails set ItemName='" & Trim(Me.DBGrid1.Columns(2)) & "' where ItemName='" & BEF_ITEM & "'")
        End If
        Data1.Refresh
    End If

    If ColIndex = 9 Then
        Set rec1 = db.OpenRecordset("select * from stock where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        If Not rec1.EOF Then
            difqty = Val(BEF_QTY) - Val(Me.DBGrid1.Columns(9))
            db.Execute ("update stock set qty=qty - " & difqty & " where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        End If
    End If
    If ColIndex = 8 Then
        Set rec1 = db.OpenRecordset("select * from stock where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        If Not rec1.EOF Then
            db.Execute ("update stock set Salerate=" & Val(Me.DBGrid1.Columns(8)) & " where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        End If
    End If

    If ColIndex = 5 Then
        Set rec1 = db.OpenRecordset("select * from Stock where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        If Not rec1.EOF Then
            db.Execute ("update stock set MRP=" & Val(Me.DBGrid1.Columns(5)) & " where  ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        End If
    End If
    If ColIndex = 4 Then
        Set rec1 = db.OpenRecordset("select * from Stock where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        If Not rec1.EOF Then
            db.Execute ("update stock set lose=" & Val(Me.DBGrid1.Columns(4)) & " where  ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        End If
    End If
    If ColIndex = 10 Then
        db.Execute ("update stock set PRate=" & Val(Me.DBGrid1.Columns(10)) & " where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
    End If
    If ColIndex = 7 Then
        db.Execute ("update stock set vat=" & Val(Me.DBGrid1.Columns(7)) & " where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
        db.Execute ("update stockdetails set vat=" & Val(Me.DBGrid1.Columns(7)) & " where ProductCode=" & Val(Me.DBGrid1.Columns(1)))
    End If
End Sub

Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    BEF_QTY = Val(Me.DBGrid1.Columns(9))
    BEF_ITEM = Trim(Me.DBGrid1.Columns(2))
    BEF_MRP = Val(Me.DBGrid1.Columns(5))
    BEF_PRATE = Val(Me.DBGrid1.Columns(7))
    BEF_PRDCODE = Val(Me.DBGrid1.Columns(1))
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
ans = MsgBox("Do You Want to Delete the Product?", vbYesNo)
If ans = 6 Then
    db.Execute "Delete * from stock where productcode=" & Me.DBGrid1.Columns(1)
    db.Execute ("Delete * from stockdetails where productcode=" & Me.DBGrid1.Columns(1))
Else
    Cancel = True
End If

End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    If usertype = "Admin" Then
        Me.DBGrid1.AllowDelete = True
        Me.DBGrid1.AllowUpdate = True
    Else
        Me.DBGrid1.AllowDelete = False
        Me.DBGrid1.AllowUpdate = False
    End If
    Data1.databasename = dbname
    Set rec1 = db.OpenRecordset("select distinct itemtype from itemmaster")
    If Not rec1.EOF Then
        While Not rec1.EOF
            'If Not IsNull(rec1!itemtype) Then
            '    Me.cboitemtype.AddItem (rec1("ItemType"))
            'End If
            rec1.MoveNext
        Wend
        'If Me.cboitemtype.ListCount > 0 Then
        '    Me.cboitemtype.ListIndex = 0
        'End If
    End If
    Data1.RecordSource = "SELECT * FROM itemmaster order by item"
    Data1.Refresh
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtBar_Change()
    Data1.RecordSource = "select * from itemmaster where barcode like '" & Replace(Me.txtBar.Text, "'", "''") & "*'"
    Data1.Refresh
End Sub

Private Sub txtsearch_Change()
    Data1.RecordSource = "select * from itemmaster where item like '*" & Replace(Me.txtsearch.Text, "'", "''") & "*'"
    Data1.Refresh
End Sub
