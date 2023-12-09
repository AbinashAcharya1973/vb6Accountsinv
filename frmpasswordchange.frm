VERSION 5.00
Begin VB.Form frmpasswordchange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PassWord Change"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4110
   Begin VB.TextBox txtnewpassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtnewuser 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtpassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtuser 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblnewpassword 
      Caption         =   "NewPassword"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblnewuser 
      Caption         =   "New User"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Exiting Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Existing User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmpasswordchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset

Private Sub txtnewpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from usertable where username='" & Me.txtuser.Text & "' and password='" & Me.txtpassword.Text & "'")
        If Not rec1.EOF Then
            db.Execute ("update usertable set username='" & Me.txtnewuser.Text & "',password='" & Me.txtnewpassword.Text & "' where username='" & rec1("username") & "' and password='" & rec1("Password") & "'")
        End If
        MsgBox "Change Sucessfully", vbOKOnly
        Me.lblnewpassword.Visible = False
        Me.lblnewuser.Visible = False
        Me.txtnewpassword.Visible = False
        Me.txtnewuser.Visible = False
        Me.txtuser.Text = ""
        Me.txtpassword.Text = ""
    End If
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from usertable where username='" & Me.txtuser.Text & "' and password='" & Me.txtpassword.Text & "'")
        If Not rec1.EOF Then
            Me.lblnewuser.Visible = True
            Me.lblnewpassword.Visible = True
            Me.txtnewuser.Visible = True
            Me.txtnewpassword.Visible = True
            Me.txtuser.Locked = True
            Me.txtpassword.Locked = True
        End If
    End If
End Sub

Private Sub txtuser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtpassword.SetFocus
    End If
End Sub
