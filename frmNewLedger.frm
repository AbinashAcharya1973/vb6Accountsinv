VERSION 5.00
Begin VB.Form frmNewLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Ledger"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmNewLedger.frx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   7200
   Begin VB.TextBox TxtAddress2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox TxtAddress1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   5055
   End
   Begin VB.ComboBox cbodr_cr 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtobalance 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4200
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtacc_name 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.ComboBox cbounder 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Under"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmNewLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp_maxid
Private Sub cbodr_cr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ans = MsgBox("Create the Ledger?", vbYesNo)
        If ans = 6 Then
            Set rs = db.OpenRecordset("select * from groups where groupid=" & Me.cbounder.ItemData(Me.cbounder.ListIndex))
            If Not rs.EOF Then
                temp_tran_type = rs("groupnature")
                Set rs = db.OpenRecordset("select * from account_nature where accounttype='" & temp_tran_type & "'")
                If Not rs.EOF Then
                    temp_dr_sign = rs("Dr")
                    temp_Cr_sign = rs("cr")
                End If
            End If
            db.Execute ("insert into LedgerMaster (AccID,AccName,GroupID,Dr,Cr,TransactionType,OBalance,Balancetype,GroupName,Address1,Address2) values(" & temp_maxid & ",'" & Me.txtacc_name.Text & "'," & Me.cbounder.ItemData(Me.cbounder.ListIndex) & ",'" & temp_dr_sign & "','" & temp_Cr_sign & "','" & temp_tran_type & "'," & Me.txtobalance.Text & ",'" & Me.cbodr_cr.Text & "','" & Me.cbounder.Text & "','" & Me.TxtAddress1.Text & "','" & Me.TxtAddress2.Text & "')")

            'Db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId) values(1,'" & AccountingPeriod & "','Opening Transaction',0,0,0," & temp_maxid & ",'Remarks','VoucherType',0," & Me.cbounder.ItemData(Me.cbounder.ListIndex) & ")")

            temp_maxid = temp_maxid + 1
            Me.txtacc_name.Text = ""
            Me.TxtAddress1.Text = ""
            Me.TxtAddress2.Text = ""
            Me.txtacc_name.SetFocus
        End If
    End If
End Sub

Private Sub cbounder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtAddress1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Set rs = db.OpenRecordset("select * from groups")
    While Not rs.EOF
        Me.cbounder.AddItem (rs("groupname"))
        Me.cbounder.ItemData(Me.cbounder.NewIndex) = rs("groupid")
        rs.MoveNext
    Wend
    If Me.cbounder.ListCount > 0 Then
        Me.cbounder.ListIndex = 0
    End If
    Me.cbodr_cr.AddItem "Dr"
    Me.cbodr_cr.AddItem "Cr"
    Me.cbodr_cr.ListIndex = 0
    Set rs = db.OpenRecordset("select max(accid) as maxid from ledgermaster")
    If Not IsNull(rs!maxid) Then
        temp_maxid = rs!maxid + 1
    Else
        temp_maxid = 1
    End If
End Sub



Private Sub txtacc_name_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbounder.SetFocus
    End If
End Sub

Private Sub txtacc_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtAddress1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtAddress2.SetFocus
    End If
End Sub

Private Sub TxtAddress2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtobalance.SetFocus
    End If
End Sub

Private Sub txtobalance_GotFocus()
    Me.txtobalance.SelStart = 0
    Me.txtobalance.SelLength = Len(Me.txtobalance.Text)
End Sub

Private Sub txtobalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbodr_cr.SetFocus
    End If
End Sub
