VERSION 5.00
Begin VB.Form frmBankMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Master"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5145
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtphone 
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
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtaddress 
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
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox txtBank 
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
      Top             =   240
      Width           =   4095
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
      TabIndex        =   4
      Top             =   1200
      Width           =   855
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
      TabIndex        =   2
      Top             =   720
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmBankMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As Recordset
Private Sub cmdSave_Click()
ans = MsgBox("Save This?", vbYesNo)
If ans = 6 Then
    Set rec1 = Db.OpenRecordset("select max(slno) as max_no from Bankmaster")
    If Not IsNull(rec1!Max_no) Then
    temp_slno = rec1!Max_no + 1
    Else
    temp_slno = 1
    End If
    Db.Execute ("insert into BankMaster (Bank,Address,Phone,Slno) values('" & Me.txtBank.Text & "','" & Me.txtAddress.Text & "','" & Me.txtphone.Text & "'," & temp_slno & ")")
    Db.Execute ("insert into BankAccount (Tdate,Particulars,VoucherType,VoucherNo,Debit,Credit,Balance,SlNo,BankCode,Remarks) values('" & Format(Date, "dd/mm/yyyy") & "','As Opening Balance','VoucherType','VoucherNo',0,0,0,1,'" & temp_slno & "','Remarks')")
    
End If
Me.txtBank.Text = ""
Me.txtAddress.Text = ""
Me.txtphone.Text = ""
Me.txtBank.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtphone.SetFocus
End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBank_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtAddress.SetFocus
End If
End Sub

Private Sub txtBank_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.cmdsave.SetFocus
End If

End Sub
