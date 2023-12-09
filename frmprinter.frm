VERSION 5.00
Begin VB.Form frmprinter 
   BorderStyle     =   0  'None
   Caption         =   "Select Printer"
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Print"
      Height          =   915
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.ComboBox cboprinter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3675
      End
      Begin VB.Label Label1 
         Caption         =   "Select Printer"
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
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1155
      Left            =   30
      Top             =   30
      Width           =   5535
   End
End
Attribute VB_Name = "frmprinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function SelectPrinter(ByVal printer_name As String) As Boolean
    Dim i As Integer
 
    SelectPrinter = True
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = printer_name Then
            Set Printer = Printers(i)
            SelectPrinter = False
            Exit For
        End If
    Next i
End Function

Private Sub cboprinter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If cboprinter.Text = "Default Printer" Then
            Unload Me
        Else
            SelectPrinter (Me.cboprinter.Text)
            Unload Me
        End If
    End If
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    cboprinter.AddItem "Default Printer"
    Dim i As Integer
    For i = 0 To Printers.Count - 1
        cboprinter.AddItem Printers(i).DeviceName
    Next i
    Me.cboprinter.ListIndex = 0
End Sub
