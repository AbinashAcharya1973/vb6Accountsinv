VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmitemslabmaster 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Slab"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7935
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\FMCG\DATA\2010-2011\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ItemSlab"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   7695
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmitemslabmaster.frx":0000
         Height          =   3495
         Left            =   120
         OleObjectBlob   =   "frmitemslabmaster.frx":0014
         TabIndex        =   18
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H0000FFFF&
         Caption         =   "Save"
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtproductcode 
         BackColor       =   &H0000C000&
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
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtdiscount 
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5880
         TabIndex        =   6
         Text            =   "0"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtqty 
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   5
         Text            =   "0"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cbomrp 
         BackColor       =   &H0000C000&
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
         Left            =   1440
         TabIndex        =   1
         Text            =   "cbomrp"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cbounit 
         BackColor       =   &H0000C000&
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
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cboitemname 
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "cboitemname"
         Top             =   240
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtdateto 
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   49152
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
      Begin MSMask.MaskEdBox txtdatefrom 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   49152
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "M.R.P"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmitemslabmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset
Private Sub cboitemname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SEARCHWORD = Trim(Me.cboitemname.Text)
        frmproductlist.Show vbModal
    End If
End Sub

Private Sub cbounit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtdatefrom.SetFocus
    End If
End Sub

Private Sub CmdSave_Click()
    ans = MsgBox("Save it!", vbYesNo)
    If ans = 6 Then
        Set rec1 = db.OpenRecordset("select * from ItemSlab where Productcode=" & Val(Me.txtproductcode.Text) & " and fromdate between #" & Mid(Me.txtdatefrom.Text, 4, 2) & "/" & Mid(Me.txtdatefrom.Text, 1, 2) & "/" & Mid(Me.txtdatefrom.Text, 7, 4) & "# and #" & Mid(Me.txtdateto.Text, 4, 2) & "/" & Mid(Me.txtdateto.Text, 1, 2) & "/" & Mid(Me.txtdateto.Text, 7, 4) & "# and qty=" & Val(Me.txtQty.Text))
        If rec1.EOF Then
            db.Execute ("insert into ItemSlab (ProductCode,ItemName,Units,Qty,FromDate,ToDate,Discount,MRP) values(" & Val(Me.txtproductcode.Text) & ",'" & Me.cboitemname.Text & "','" & Me.cbounit.Text & "'," & Val(Me.txtQty.Text) & ",'" & Me.txtdatefrom.Text & "','" & Me.txtdateto.Text & "','" & Trim(Me.txtdiscount.Text) & "'," & Me.cbomrp.Text & ")")
            Me.txtQty.Text = 0
            Me.txtdiscount.Text = 0
            Me.cboitemname.SetFocus
        Else
            MsgBox "Allready exist!", vbCritical
            Me.cboitemname.SetFocus
        End If
        Me.Data1.RecordSource = "select * from ItemSlab order by fromdate"
        Me.Data1.Refresh
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    FORMNAME = "itemslab"
    Me.cbounit.AddItem ("CBB")
    Me.cbounit.AddItem ("PKT")
    Me.cbounit.ListIndex = 0
    Me.txtdatefrom.Text = Format(Date, "dd/mm/yyyy")
    Me.txtdateto.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub txtdatefrom_GotFocus()
    Me.txtdatefrom.SelStart = 0
    Me.txtdatefrom.SelLength = Len(Me.txtdatefrom.Text)
End Sub

Private Sub txtdatefrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtdateto.SetFocus
    End If
End Sub
Private Sub txtdateto_GotFocus()
    Me.txtdateto.SelStart = 0
    Me.txtdateto.SelLength = Len(Me.txtdateto.Text)
End Sub

Private Sub txtdateto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtQty.SetFocus
    End If
End Sub

Private Sub txtdiscount_GotFocus()
    Me.txtdiscount.SelStart = 0
    Me.txtdiscount.SelLength = Len(Me.txtdiscount.Text)
End Sub

Private Sub txtdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub

Private Sub txtqty_GotFocus()
    Me.txtQty.SelStart = 0
    Me.txtQty.SelLength = Len(Me.txtQty.Text)
End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtdiscount.SetFocus
    End If
End Sub
