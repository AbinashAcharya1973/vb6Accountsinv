VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmdeliverychallan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Challan/Mandays"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8760
   Begin VB.TextBox txttotalqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6900
      TabIndex        =   10
      Text            =   "0"
      Top             =   7410
      Width           =   1515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\videotutorial\EnLite\EnliteFinal\DATA\2022-2023\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tempstocktran"
      Top             =   7500
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      TabIndex        =   9
      Top             =   7860
      Width           =   1065
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4410
      TabIndex        =   8
      Top             =   7860
      Width           =   1065
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3300
      TabIndex        =   7
      Top             =   7860
      Width           =   1065
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2190
      TabIndex        =   6
      Top             =   7860
      Width           =   1065
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6930
      TabIndex        =   5
      Text            =   "0"
      Top             =   2220
      Width           =   1575
   End
   Begin VB.ComboBox cbobatchno 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2190
      Width           =   3075
   End
   Begin VB.ComboBox cboproduct 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1740
      Width           =   7335
   End
   Begin VB.ComboBox cboparty 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   7335
   End
   Begin VB.TextBox txtslno 
      Appearance      =   0  'Flat
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
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   0
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmdeliverychallan.frx":0000
      Height          =   4455
      Left            =   240
      OleObjectBlob   =   "frmdeliverychallan.frx":0014
      TabIndex        =   11
      Top             =   2850
      Width           =   8325
   End
   Begin MSMask.MaskEdBox txtdate 
      Height          =   315
      Left            =   6810
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   19
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblqty 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   4260
      TabIndex        =   18
      Top             =   2250
      Width           =   1485
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   16
      Top             =   2220
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   15
      Top             =   1770
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   14
      Top             =   570
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6210
      TabIndex        =   13
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   12
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmdeliverychallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec1 As Recordset, rec2 As Recordset, deletef As Boolean

Private Sub cbobatchno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtqty.SetFocus
End If
End Sub

Private Sub cboparty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cboproduct.SetFocus
End If
End Sub

Private Sub cboproduct_Change()
Set rec2 = db.OpenRecordset("select * from stockdetails where productcode=" & Me.cboproduct.ItemData(Me.cboproduct.ListIndex))
While Not rec2.EOF
    Me.cbobatchno.AddItem rec2("batchno")
    Me.lblqty.Caption = rec2("qty")
    rec2.MoveNext
Wend
If Me.cbobatchno.ListCount > 0 Then
    Me.cbobatchno.ListIndex = 0
End If
End Sub

Private Sub cboproduct_Click()
cboproduct_Change
End Sub

Private Sub cboproduct_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cbobatchno.SetFocus
End If
If KeyCode = 27 Then
    Me.cmdsave.SetFocus
End If
End Sub

Private Sub cbopurpose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cboproduct.SetFocus
End If
End Sub

Private Sub cmddelete_Click()
deletef = True
Me.txtslno.SetFocus
Me.txtslno.Locked = False
End Sub

Private Sub cmdedit_Click()
Me.txtslno.Locked = False
Me.txtslno.SetFocus
deletef = False
End Sub

Private Sub cmdprint_Click()
'frmprintoutchallan.Show 0
frmdchallanreport.Show 0
End Sub

Private Sub cmdsave_Click()
ans = MsgBox("Save the Delivery Challan?", vbYesNo)
If ans = 6 Then
    Set rec = db.OpenRecordset("select * from deliverychallan where challanno=" & Me.txtslno.Text)
    If Not rec.EOF Then
'        Set rec1 = db.OpenRecordset("select * from deliverychallandetails where challanno=" & Me.txtslno.Text)
'        While Not rec1.EOF
'            db.Execute "update stock set qty=qty+" & rec1("qty") & " where productcode=" & rec1("productcode")
'            db.Execute "update stockdetails set qty=qty+" & rec1("qty") & " where productcode=" & rec1("productcode") & " and batchno='" & rec1("batchno") & "'"
'            rec1.MoveNext
'        Wend
        db.Execute ("delete * from deliverychallan where challanno=" & Me.txtslno.Text)
        db.Execute ("delete * from deliverychallandetails where challanno=" & Me.txtslno.Text)
    End If
    
    
    db.Execute "insert into deliverychallan (ChallanNO,ChallanDaate,Party,accid,TotalQty) values(" & Me.txtslno.Text & ",'" & Me.txtdate.Text & "','" & Me.cboparty.Text & "'," & Me.cboparty.ItemData(Me.cboparty.ListIndex) & "," & Me.txttotalqty.Text & ")"
    
    Set rec = db.OpenRecordset("select * from tempstocktran")
    While Not rec.EOF
        db.Execute "insert into deliverychallandetails (ChallanNo,ProductCode,ProductName,Qty,Batchno) values(" & Me.txtslno.Text & "," & rec("productcode") & ",'" & rec("ProductName") & "'," & rec("Qty") & ",'" & rec("Batchno") & "')"
        'db.Execute "update " & Me.cbofromstock.Text & " set qty=qty-" & rec("qty") & " where productcode=" & rec("productcode")
        'db.Execute "update " & Me.cbofromstock.Text & "details set qty=qty-" & rec("qty") & " where productcode=" & rec("productcode") & " and batchno='" & rec("batchno") & "'"
        
        rec.MoveNext
    Wend
    db.Execute "delete * from tempstocktran"
    Me.txttotalqty.Text = 0
    Me.Data1.Refresh
    Me.DBGrid1.Refresh
    Me.txtslno.Text = Val(Me.txtslno.Text) + 1
    deletef = False
    Me.txtdate.SetFocus
End If
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
Me.txttotalqty.Text = Val(Me.txttotalqty.Text) - Val(Me.DBGrid1.Columns(3))
End Sub

Private Sub Form_Load()
Me.Data1.databasename = db.Name
Me.txtdate.Text = Format(Date, "dd/mm/yyyy")
Set rec = db.OpenRecordset("select max(challanno) as maxslno from deliverychallan")
If Not IsNull(rec!maxslno) Then
    Me.txtslno.Text = rec!maxslno + 1
Else
    Me.txtslno.Text = 1
End If

Set rec1 = db.OpenRecordset("select * from partydr")
While Not rec1.EOF
    Me.cboparty.AddItem rec1("party")
    Me.cboparty.ItemData(Me.cboparty.NewIndex) = rec1("accid")
    rec1.MoveNext
Wend
If Me.cboparty.ListCount > 0 Then
    Me.cboparty.ListIndex = 0
End If
Set rec1 = db.OpenRecordset("select * from itemmaster")
While Not rec1.EOF
    Me.cboproduct.AddItem rec1("item")
    Me.cboproduct.ItemData(Me.cboproduct.NewIndex) = rec1("productcode")
    rec1.MoveNext
Wend
If Me.cboproduct.ListCount > 0 Then
    Me.cboproduct.ListIndex = 0
End If
  deletef = False
End Sub

Private Sub frmproduct_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
db.Execute "delete * from tempstocktran"
End Sub

Private Sub txtdate_GotFocus()
Me.txtdate.SelStart = 0
Me.txtdate.SelLength = Len(Me.txtdate.Text)
End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cboparty.SetFocus
End If
End Sub

Private Sub txtqty_GotFocus()
Me.txtqty.SelStart = 0
Me.txtqty.SelLength = Len(Me.txtqty.Text)
End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    db.Execute "insert into tempstocktran (ProductCode,ProductName,Batchno,qty) values(" & Me.cboproduct.ItemData(Me.cboproduct.ListIndex) & ",'" & Me.cboproduct.Text & "','" & Me.cbobatchno.Text & "'," & Me.txtqty.Text & ")"
    Me.Data1.Refresh
    Me.DBGrid1.Refresh
    Me.txttotalqty.Text = Val(Me.txttotalqty.Text) + Val(Me.txtqty.Text)
    Me.cboproduct.SetFocus
End If
End Sub

Private Sub txtslno_GotFocus()
Me.txtslno.SelStart = 0
Me.txtslno.SelLength = Len(Me.txtslno.Text)
End Sub

Private Sub txtslno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec = db.OpenRecordset("select * from deliverychallan where challanno=" & Me.txtslno.Text)
        If rec.EOF Then
            MsgBox "Challan No not Found", vbCritical
        Else
            i = 0
            Me.txtdate.Text = Format(rec("ChallanDaate"), "dd/mm/yyyy")
            Me.txttotalqty.Text = rec("totalqty")
            Me.txtdate.SetFocus
            Do While i < Me.cboparty.ListCount
                If Me.cboparty.ItemData(i) = rec("accid") Then

                    Me.cboparty.ListIndex = i
                    Exit Do
                Else
                    i = i + 1
                End If
            Loop
            
            
            Set rec1 = db.OpenRecordset("select * from deliverychallandetails where challanno=" & Me.txtslno.Text)
            While Not rec1.EOF
                db.Execute "insert into tempstocktran (ProductCode,ProductName,Qty,Batchno) values(" & rec1("ProductCode") & ",'" & rec1("ProductName") & "'," & rec1("Qty") & ",'" & rec1("Batchno") & "')"
                rec1.MoveNext
            Wend
            Me.Data1.Refresh
            Me.DBGrid1.Refresh
            If deletef = True Then
                ans = MsgBox("Delete the Challan?", vbYesNo)
                If ans = 6 Then
                    
                    db.Execute ("delete * from deliverychallan where challanno=" & Me.txtslno.Text)
                    db.Execute ("delete * from deliverychallandetails where challanno=" & Me.txtslno.Text)
                    db.Execute ("delete * from tempstocktran")
                    Me.Data1.Refresh
                    Me.DBGrid1.Refresh
                    Me.txtdate.SetFocus
                    Set rec2 = db.OpenRecordset("select max(challanno) as maxslno from deliverychallan")
                    If Not IsNull(rec2!maxslno) Then
                        Me.txtslno.Text = rec2!maxslno + 1
                    Else
                        Me.txtslno.Text = 1
                    End If
                    deletef = False
                End If
            End If
        End If
    End If
End Sub


