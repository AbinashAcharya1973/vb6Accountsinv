VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmJournalVoucherView 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Voucher Edit"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   13305
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   13095
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
         TabIndex        =   8
         Top             =   120
         Width           =   975
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
         TabIndex        =   7
         Top             =   120
         Width           =   1095
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
         TabIndex        =   5
         Top             =   120
         Width           =   1935
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "JournalTr"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmJournalVoucherView.frx":0000
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "frmJournalVoucherView.frx":0014
         TabIndex        =   3
         Top             =   120
         Width           =   12855
      End
   End
End
Attribute VB_Name = "frmJournalVoucherView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.txtfromdt.Text = Format(Date, "DD/MM/YYYY")
    Me.txtToDt.Text = Format(Date, "DD/MM/YYYY")
    Data1.databasename = dbname
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
        Data1.RecordSource = "select * from JournalTr  where TDate between #" & Accperiod_month & "/" & Accperiod_day & "/" & Accperiod_year & "# and #" & temp_month & "/" & temp_day & "/" & temp_year & "# order by Tdate,SLno"
        Data1.Refresh
        Me.txtfromdt.SetFocus
    End If
End Sub
