VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "Crystl32.OCX"
Begin VB.Form frmItemSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Wise Sales Report"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6000
   Begin MSMask.MaskEdBox txtto 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtfrom 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cboproducttype 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox cboitemtype 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox cbobrandname 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cboitemname 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   4095
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4440
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "To"
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
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "From "
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
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Product Type"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Item Type"
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
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Brand"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Item Name"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frmItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432

Private Sub cbobrandname_Change()
    Set rec1 = db.OpenRecordset("select distinct Item from ItemMaster where ProductType='" & Me.cboproducttype.Text & "' and ItemType='" & Me.cboitemtype.Text & "' and Brand='" & Me.cbobrandname.Text & "'")
    Me.cboitemname.Clear
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboitemname.AddItem (rec1("Item"))
            rec1.MoveNext
        Wend
        If Me.cboitemname.ListCount > 0 Then
            Me.cboitemname.ListIndex = 0
        End If
    End If

End Sub
Private Sub cbobrandname_Click()
    cbobrandname_Change
End Sub

Private Sub cbobrandname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboitemname.SetFocus
    End If
End Sub

Private Sub cboitemtype_Change()
    Set rec1 = db.OpenRecordset("select Distinct Brand from ItemMaster where ProductType='" & Me.cboproducttype.Text & "' and ItemType='" & Me.cboitemtype.Text & "'")
    Me.cbobrandname.Clear
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cbobrandname.AddItem (rec1("Brand"))
            rec1.MoveNext
        Wend
        If Me.cbobrandname.ListCount > 0 Then
            Me.cbobrandname.ListIndex = 0
        End If
    End If
End Sub
Private Sub cboitemtype_Click()
    cboitemtype_Change
End Sub
Private Sub cboproducttype_Change()
    Set rec1 = db.OpenRecordset("select distinct ItemType from ItemMaster where ProductType='" & Me.cboproducttype.Text & "'")
    Me.cboitemtype.Clear
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboitemtype.AddItem (rec1("ItemType"))
            rec1.MoveNext
        Wend
        If Me.cboitemtype.ListCount > 0 Then
            Me.cboitemtype.ListIndex = 0
        End If
    End If
End Sub
Private Sub cboproducttype_Click()
    cboproducttype_Change
End Sub
Private Sub CmdSearch_Click()
'Set rec1 = Db.OpenRecordset("select distinct size form itemmaster where ItemType='" & Me.cboitemtype.Text & "' and Item='" & Me.cboitemname.Text & "'")
'While Not rec1.EOF
'Db.Execute ("insert into itemsearch (itemtype,item,size) values('" & Me.cboitemtype.Text & "','" & Me.cboitemname.Text & "','" & rec1("size") & "')")
'rec1.MoveNext
'Wend
'Set rec1 = Db.OpenRecordset("select InvoiceHead.InvDate,InvoiceDetails.ItemType,InvoiceDetails.Itemname,InvoiceDetails.size,InvoiceDetails.PrRate,InvoiceDetails.Qty from invoicehead inner join InvoiceDetails on InvoiceHead.InvNo=InvoiceDetails.InvNo where InvoiceHead.InvDate  between #" & Me.txtfrom_dt_m.Text & "/" & Me.txtfrom_dt_d.Text & "/" & Me.txtfrom_dt_y.Text & "# and #" & Me.txtto_dt_m.Text & "/" & Me.txtto_dt_d.Text & "/" & Me.txtto_dt_y.Text & "#")
'While Not rec1.EOF
'    'db.Execute("update itemsearch set qty=qty+ " & rec1("Qty") & " where itemtype='" & me.cboitemtype.Text & "'
'    Debug.Print rec1("invdate")
'    rec1.MoveNext
'    Wend
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.CrystalReport1.ReportFileName = App.Path & "\itemsearch.rpt"
    Set rec1 = db.OpenRecordset("select * from product")
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboproducttype.AddItem (rec1("Productname"))
            rec1.MoveNext
        Wend
        If Me.cboproducttype.ListCount > 0 Then
            Me.cboproducttype.ListIndex = 0
        End If
    End If


End Sub
Private Sub txtSortNo_Change()

End Sub

Private Sub txtFrom_GotFocus()
    Me.txtFrom.SelStart = 0
    Me.txtFrom.SelLength = Len(Me.txtFrom.Text)
End Sub

Private Sub txtTo_GotFocus()
    Me.txtTo.SelStart = 0
    Me.txtTo.SelLength = Len(Me.txtTo.Text)
End Sub
