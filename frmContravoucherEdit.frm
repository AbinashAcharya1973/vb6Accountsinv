VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmContravoucherEdit 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contra Voucher Edit"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   13320
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      RecordSource    =   "ContraDetails"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   13095
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtToDt 
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox txtfromdt 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   0
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Left            =   8520
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
         Left            =   8520
         TabIndex        =   6
         Top             =   480
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
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmContravoucherEdit.frx":0000
         Height          =   2295
         Left            =   120
         OleObjectBlob   =   "frmContravoucherEdit.frx":0014
         TabIndex        =   10
         Top             =   2160
         Width           =   12855
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmContravoucherEdit.frx":108F
         Height          =   1935
         Left            =   120
         OleObjectBlob   =   "frmContravoucherEdit.frx":10A3
         TabIndex        =   9
         Top             =   120
         Width           =   12855
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
         RecordSource    =   "ContraHead"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmContravoucherEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset, BEFORE_AMOUNT
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute BEFORE_AMOUNT.VB_VarUserMemId = 1073938432
Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    ans = MsgBox("Update This?", vbYesNo)
    If ans = 6 Then
        temp_day = Left(Trim(Me.DBGrid1.Columns(1)), 2)
        temp_month = Mid(Trim(Me.DBGrid1.Columns(1)), 4, 2)
        temp_year = Right(Trim(Me.DBGrid1.Columns(1)), 4)
        If ColIndex = 6 Then
            '----------Dr Account Update--------------
            Diff_Amount = BEFORE_AMOUNT - Val(Me.DBGrid1.Columns(6))
            Set rec1 = db.OpenRecordset("select * from LedgerMaster where Accid=" & Me.DBGrid1.Columns(2))
            If Not rec1.EOF Then
                temp_sign = rec1("Cr")
                Set rec2 = db.OpenRecordset("select * from LedgerTran where Accid=" & Me.DBGrid1.Columns(2) & " and VoucherType='Contra' and VoucherSlno=" & Trim(Me.DBGrid1.Columns(0)) & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                If Not rec2.EOF Then
                    db.Execute ("update ledgerTran set Dr=Dr-" & Diff_Amount & ",balance=balance " & temp_sign & Diff_Amount & " where AccId=" & Me.DBGrid1.Columns(2) & "  and VoucherType='Contra' and VoucherSlno=" & Trim(Me.DBGrid1.Columns(0)) & " and SlNo=" & rec2("SlNo") & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                    db.Execute ("update ledgerTran set Balance=Balance " & temp_sign & Diff_Amount & " where AccId=" & Me.DBGrid1.Columns(2) & " and Slno>" & rec2("SlNo"))
                End If
            End If
            '----------Cr Account Update---------------
            Set rec1 = db.OpenRecordset("select * from LedgerMaster where Accid=" & Me.DBGrid1.Columns(4))
            If Not rec1.EOF Then
                temp_crsign = rec1("Dr")
                Set rec2 = db.OpenRecordset("select * from LedgerTran where Accid=" & Me.DBGrid1.Columns(4) & " and VoucherType='Contra' and VoucherSlno=" & Trim(Me.DBGrid1.Columns(0)) & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                If Not rec2.EOF Then
                    db.Execute ("update ledgerTran set Cr=Cr-" & Diff_Amount & ",Balance=Balance " & temp_crsign & Diff_Amount & " where  Accid=" & Me.DBGrid1.Columns(4) & " and VoucherType='Contra' and VoucherSlno=" & Trim(Me.DBGrid1.Columns(0)) & " and SlNo=" & rec2("SlNo") & " and TDate=#" & temp_month & "/" & temp_day & "/" & temp_year & "#")
                    db.Execute ("update LedgerTran set Balance=Balance " & temp_crsign & Diff_Amount & " where AccId=" & Me.DBGrid1.Columns(4) & " and SlNo>" & rec2("SlNo"))
                End If
            End If
        End If
    End If
End Sub

Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    BEFORE_AMOUNT = Val(Me.DBGrid1.Columns(6))
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Me.Data1.databasename = dbname
    Me.Data2.databasename = dbname
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
        Data1.RecordSource = "select * from ContraVoucher  where ContraDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "# order by ContraDate,SLno"
        Data1.Refresh
        Me.txtfromdt.SetFocus
    End If

End Sub
