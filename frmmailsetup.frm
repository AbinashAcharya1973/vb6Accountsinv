VERSION 5.00
Begin VB.Form frmmailsetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail Setup"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbossl 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2820
      Width           =   675
   End
   Begin VB.TextBox txtemailid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   4
      Top             =   180
      Width           =   3555
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "009"
      Top             =   840
      Width           =   3555
   End
   Begin VB.TextBox txtserver 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   2
      Top             =   1500
      Width           =   3555
   End
   Begin VB.TextBox txtsmtpport 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   3555
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3480
      Width           =   850
   End
   Begin VB.Label Label1 
      Caption         =   "Email ID"
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
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
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
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "SMTP Server"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Label6 
      Caption         =   "SMTP Port"
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
      TabIndex        =   6
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "SSL"
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
      TabIndex        =   5
      Top             =   2820
      Width           =   975
   End
End
Attribute VB_Name = "frmmailsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset
Private Sub cmdupdate_Click()
On Error GoTo errtrap
db.Execute "update mailsetup set sendermailid='" & Me.txtemailid.Text & "',smtpserver='" & Me.txtserver.Text & "',smtpport='" & Me.txtsmtpport.Text & "',password='" & Me.txtpassword.Text & "',smtpusessl='" & Me.cbossl.Text & "'"
MsgBox "Mail Setup Updated", vbOKOnly
Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.cbossl.AddItem "0"
    Me.cbossl.AddItem "1"
    Me.cbossl.ListIndex = 0
    Set rec1 = db.OpenRecordset("select * from mailsetup")
    If Not rec1.EOF Then
        Me.txtemailid.Text = IIf(IsNull(rec1("sendermailid")), "", rec1("sendermailid"))
        Me.txtserver.Text = IIf(IsNull(rec1("smtpserver")), "", rec1("smtpserver"))
        Me.txtsmtpport.Text = IIf(IsNull(rec1("smtpport")), "", rec1("smtpport"))
        Me.txtpassword.Text = IIf(IsNull(rec1("password")), "", rec1("password"))
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub gstno_Change()

End Sub

Private Sub txtAddress_Change()

End Sub

Private Sub txtEmail_Change()

End Sub
