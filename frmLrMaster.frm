VERSION 5.00
Begin VB.Form frmLrMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LR Master"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5580
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtFax 
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
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtPhone 
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
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtLrName 
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
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Fax"
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
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label4 
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
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Email"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   615
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
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
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmLrMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Private Sub CmdSave_Click()
    ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then
        Set rec1 = db.OpenRecordset("select max(slno) as max_no from LrLedger")
        If Not IsNull(rec1!max_no) Then
            temp_slno = rec1!max_no + 1
        Else
            temp_slno = 1
        End If
        db.Execute ("insert into LrMaster(Slno,LrName,Address,Phone,Fax,Email) values(" & temp_slno & ",'" & Me.txtLrName.Text & "','" & Me.txtaddress.Text & "','" & Me.txtPhone.Text & "','" & Me.txtFax.Text & "','" & Me.txtemail.Text & "')")
        db.Execute ("insert into LrLedger (Tdate,Particulars,VoucherType,VoucherNo,Debit,Credit,Balance,SlNo,LrCode,Remarks) values('" & Format(Date, "dd/mm/mm") & "','As Opening Balance','VoucherType','VoucherNo',0,0,0,1,'" & temp_slno & "','Remarks'")
        Me.txtLrName.Text = ""
        Me.txtaddress.Text = ""
        Me.txtPhone.Text = ""
        Me.txtFax.Text = ""
        Me.txtemail.Text = ""
        Me.txtLrName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0

End Sub
