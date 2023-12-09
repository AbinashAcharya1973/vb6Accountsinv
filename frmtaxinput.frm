VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmtaxinput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Input Report"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4965
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4755
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
         Height          =   360
         Left            =   1620
         TabIndex        =   3
         Top             =   1050
         Width           =   1470
      End
      Begin MSMask.MaskEdBox txtdateto 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
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
      Begin MSMask.MaskEdBox txtdatefrom 
         Height          =   315
         Left            =   720
         TabIndex        =   0
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
         Left            =   2400
         TabIndex        =   5
         Top             =   240
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
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmtaxinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    temp_from_date = Mid(Me.txtdatefrom.Text, 1, 2)
    temp_from_month = Mid(Me.txtdatefrom.Text, 4, 2)
    temp_from_year = Mid(Me.txtdatefrom.Text, 7, 4)

    temp_from = temp_from_date & "/" & temp_from_month & "/" & temp_from_year

    temp_to_date = Mid(Me.txtdateto.Text, 1, 2)
    temp_to_month = Mid(Me.txtdateto.Text, 4, 2)
    temp_to_year = Mid(Me.txtdateto.Text, 7, 4)

    temp_to = temp_to_date & "/" & temp_to_month & "/" & temp_to_year

    db.Execute ("update purchasehead set fromdt='" & temp_from & "',todt='" & temp_to & "'")

    Me.CrystalReport1.SelectionFormula = "{PurchaseHead.InvDate} in Date (" & temp_from_year & ", " & temp_from_month & ", " & temp_from_date & ") to Date (" & temp_to_year & ", " & temp_to_month & ", " & temp_to_date & ")"
    Me.CrystalReport1.PrinterName = Printer.DeviceName
    Me.CrystalReport1.PrinterDriver = Printer.DriverName
    Me.CrystalReport1.PrinterPort = Printer.Port
    Me.CrystalReport1.PrintReport
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
