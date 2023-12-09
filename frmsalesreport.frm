VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmsalesreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmsalesreport.frx":0000
   ScaleHeight     =   1845
   ScaleWidth      =   4680
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3840
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "View/Print"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin MSMask.MaskEdBox txtdateto 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   360
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
   Begin MSMask.MaskEdBox txtdatefrom 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   360
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From "
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
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmsalesreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
    temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
    temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)

    temp_from = temp_from_year & "/" & temp_from_month & "/" & temp_from_date

    temp_to_date = Mid(Me.txtdateto.Text, 1, 2)
    temp_to_month = Mid(Me.txtdateto.Text, 4, 2)
    temp_to_year = Mid(Me.txtdateto.Text, 7, 4)

    temp_to = temp_to_year & "/" & temp_to_month & "/" & temp_to_date

    Me.CrystalReport1.SelectionFormula = "{InvoiceHead.InvDate} in Date (" & temp_from_year & "," & temp_from_month & "," & temp_from_date & ") to Date (" & temp_to_year & "," & temp_to_month & "," & temp_to_date & ")"
    Me.CrystalReport1.PrinterName = Printer.DeviceName
    Me.CrystalReport1.PrinterDriver = Printer.DriverName
    Me.CrystalReport1.PrinterPort = Printer.Port
    Me.CrystalReport1.PrintReport
End Sub

Private Sub cmdprint_GotFocus()
    Me.txtdateto.SelStart = 0
    Me.txtdateto.SelLength = Len(Me.txtdateto.Text)
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.CrystalReport1.ReportFileName = App.Path & "\salesreport.rpt"
    Me.Top = 0
    Me.Left = 0
    Me.txtdatefrom.Text = Format(Date, "dd/mm/yyyy")
    Me.txtdateto.Text = Format(Date, "dd/mm/yyyy")
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

