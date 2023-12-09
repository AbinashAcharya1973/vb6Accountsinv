VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmTaxutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Tax Out Put"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4860
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   360
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   4635
      Begin VB.Frame Frame2 
         Caption         =   "For"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   3060
         TabIndex        =   2
         Top             =   180
         Width           =   1365
         Begin VB.CheckBox chkretailinv 
            Caption         =   "RETAIL INVOICE"
            Height          =   330
            Left            =   90
            TabIndex        =   4
            Top             =   855
            Width           =   1185
         End
         Begin VB.CheckBox chktaxinv 
            Caption         =   "TAX INVOICE"
            Height          =   330
            Left            =   90
            TabIndex        =   3
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdprint 
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
         Height          =   240
         Left            =   1620
         TabIndex        =   1
         Top             =   1170
         Width           =   1230
      End
      Begin MSMask.MaskEdBox txtdateto 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   720
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
         TabIndex        =   6
         Top             =   240
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
         TabIndex        =   8
         Top             =   240
         Width           =   855
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
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmTaxutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
    temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
    temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)

    temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year

    temp_to_date = Mid(Me.txtdateto.Text, 1, 2)
    temp_to_month = Mid(Me.txtdateto.Text, 4, 2)
    temp_to_year = Mid(Me.txtdateto.Text, 7, 4)

    temp_to = temp_to_month & "/" & temp_to_date & "/" & temp_to_year

    db.Execute ("Update InvoiceHead Set fromdt='" & temp_from & "',todt='" & temp_to & "'")

    If chktaxinv.Value = 1 And chkretailinv.Value = 0 Then
        Me.CrystalReport2.SelectionFormula = "{InvoiceHead.InvDate} in Date (" & temp_from_year & ", " & temp_from_month & ", " & temp_from_date & ") to Date (" & temp_to_year & ", " & temp_to_month & ", " & temp_to_date & ") and {InvoiceHead.InvType} = 'TAX'"
        Me.CrystalReport2.PrintReport
    End If
    Me.CrystalReport2.PrinterName = Printer.DeviceName
    Me.CrystalReport2.PrinterDriver = Printer.DriverName
    Me.CrystalReport2.PrinterPort = Printer.Port
    If chkretailinv.Value = 1 And chktaxinv.Value = 0 Then
        Me.CrystalReport2.SelectionFormula = "{InvoiceHead.InvDate} in Date (" & temp_from_year & ", " & temp_from_month & ", " & temp_from_date & ") to Date (" & temp_to_year & ", " & temp_to_month & ", " & temp_to_date & ") and {InvoiceHead.InvType} = 'RETAIL'"
        Me.CrystalReport2.PrintReport

    End If
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
Me.txtdatefrom.Text = Format(Date, "dd/mm/yyyy")
Me.txtdateto.Text = Format(Date, "dd/mm/yyyy")
Me.CrystalReport2.ReportFileName = App.Path & "\taxoutput.rpt"
Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
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
