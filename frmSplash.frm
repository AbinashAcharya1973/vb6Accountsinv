VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4995
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   1620
      Top             =   3420
   End
   Begin VB.Frame Frame1 
      Height          =   4995
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   495
         Left            =   6300
         TabIndex        =   8
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "EnLite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1275
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   3795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SAS with websync"
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
         TabIndex        =   6
         Top             =   2100
         Width           =   4935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4020
         TabIndex        =   4
         Top             =   4620
         Width           =   1815
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   8280
         TabIndex        =   1
         Top             =   4620
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   4560
         Width           =   9375
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8160
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform : Windows x86/64"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   5985
         Left            =   -240
         Picture         =   "frmSplash.frx":000C
         Top             =   -240
         Width           =   10980
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmstartingscreen.Show 0
    Unload Me
End If
If KeyAscii = 27 Then
    Unload Me
End If

    
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()

End Sub

Private Sub Timer1_Timer()
    Me.Timer1.Enabled = False
    frmstartingscreen.Show 0
    Unload Me
End Sub
