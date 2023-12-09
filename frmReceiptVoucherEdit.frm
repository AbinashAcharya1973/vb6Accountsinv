VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReceiptVoucherEdit 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Voucher Edit"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13305
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmReceiptVoucherEdit.frx":0000
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "frmReceiptVoucherEdit.frx":0014
         TabIndex        =   9
         Top             =   2280
         Width           =   12855
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmReceiptVoucherEdit.frx":109F
         Height          =   2055
         Left            =   120
         OleObjectBlob   =   "frmReceiptVoucherEdit.frx":10B3
         TabIndex        =   8
         Top             =   120
         Width           =   12855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   13095
      Begin MSMask.MaskEdBox txtToDt 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         HideSelection   =   0   'False
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
      Begin MSMask.MaskEdBox txtfromdt 
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         HideSelection   =   0   'False
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
      Begin VB.Label Label1 
         Caption         =   "Financial Year         :"
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
         Left            =   8640
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Assessment Year    :"
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
         Left            =   8640
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "From"
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   120
         Width           =   975
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ReceiptHead"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      RecordSource    =   "ReceiptDetails"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmReceiptVoucherEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BEFORE_AMOUNT, BEFORE_DISCOUNT, rec1 As DAO.Recordset, rec2 As Recordset
Attribute BEFORE_DISCOUNT.VB_VarUserMemId = 1073938432
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432

Private Sub Command1_Click()

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    Data2.RecordSource = "select * from receiptdetails where Receiptno=" & Me.DBGrid1.Columns(0) & ""
    Data2.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Data2.databasename = dbname
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtfromdt_GotFocus()
    Me.txtfromdt.SelStart = 0
    Me.txtfromdt.SelLength = Len(Me.txtfromdt.Text)
End Sub
Private Sub txtfromdt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtToDt.SetFocus
    End If
End Sub
Private Sub txtToDt_GotFocus()
    Me.txtToDt.SelStart = 0
    Me.txtToDt.SelLength = Len(Me.txtToDt.Text)
End Sub
Private Sub txtToDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        temp_day = Left((Me.txtToDt.Text), 2)
        temp_month = Mid((Me.txtToDt.Text), 4, 2)
        temp_year = Right((Me.txtToDt.Text), 4)

        Accperiod_day = Left(Me.txtfromdt.Text, 2)
        Accperiod_month = Mid(Me.txtfromdt.Text, 4, 2)
        Accperiod_year = Right(Me.txtfromdt.Text, 4)
        Data1.RecordSource = "select * from Receipthead  where ReceiptDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "# order by ReceiptDate,receiptno"
        Data1.Refresh
        Me.txtfromdt.SetFocus
    End If
End Sub
