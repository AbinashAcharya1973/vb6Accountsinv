VERSION 5.00
Begin VB.Form frmCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Master"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5025
   Begin VB.TextBox txtaddress2 
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
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1320
      Width           =   3735
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
      TabIndex        =   14
      Top             =   5040
      Width           =   850
   End
   Begin VB.TextBox txtdealin 
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
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4500
      Width           =   3735
   End
   Begin VB.TextBox gstno 
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
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3960
      Width           =   3735
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3360
      Width           =   3735
   End
   Begin VB.TextBox txtpin 
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtPhone 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox txtCompany 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label7 
      Caption         =   "Deal In"
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
      TabIndex        =   13
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "GST No"
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
      TabIndex        =   11
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Email @"
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
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "PIN"
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
      TabIndex        =   8
      Top             =   1980
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Phone"
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
      Top             =   2820
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
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
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Company"
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
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset

Private Sub cmdupdate_Click()
    ans = MsgBox("Update the Company Info?", vbYesNo)
    If ans = 6 Then
        db.Execute "update companymaster set company='" & Me.txtCompany.Text & "',address='" & Me.txtAddress.Text & "',address1='" & Me.txtaddress2.Text & "',phone='" & Me.txtPhone.Text & _
                   "',email='" & Me.txtEmail.Text & "',taxno='" & Me.gstno.Text & "',dealin='" & Me.txtdealin.Text & "',pin='" & Me.txtpin.Text & "'"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Set rec1 = db.OpenRecordset("select * from companymaster")
    If Not rec1.EOF Then
        Me.txtCompany.Text = rec1("company")
        Me.txtAddress.Text = IIf(IsNull(rec1("address")), "", rec1("address"))
        Me.txtaddress2.Text = IIf(IsNull(rec1("address1")), "", rec1("address1"))
        Me.txtPhone.Text = IIf(IsNull(rec1("phone")), "", rec1("phone"))
        Me.txtEmail.Text = IIf(IsNull(rec1("email")), "", rec1("email"))
        Me.gstno.Text = IIf(IsNull(rec1("taxno")), "", rec1("taxno"))
        Me.txtdealin.Text = IIf(IsNull(rec1("dealin")), "", rec1("dealin"))
        Me.txtpin.Text = IIf(IsNull(rec1("pin")), "", rec1("pin"))
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Label8_Click()

End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

        Me.cmdupdate.SetFocus
    End If

End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtPhone.SetFocus
    End If
End Sub

Private Sub txtCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtAddress.SetFocus
    End If
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtEmail.SetFocus
    End If
End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Me.txtFax.SetFocus
    End If
End Sub
