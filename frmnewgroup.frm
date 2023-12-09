VERSION 5.00
Begin VB.Form frmnewgroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Group"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7320
   Begin VB.ComboBox cboaffectgp 
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
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ComboBox cbonature 
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
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
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
      TabIndex        =   2
      Top             =   720
      Width           =   5055
   End
   Begin VB.TextBox txtgroup_name 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Affect Gross Profit"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "A/c Nature"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Group Name"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmnewgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp_groupcode

Private Sub cboaffectgp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ans = MsgBox("Save the Group?", vbYesNo)
        If ans = 6 Then
            db.Execute ("insert into groups (groupid,groupname,parentid,parentname,groupnature,affect_gp) values(" & temp_groupcode & ",'" & Me.txtgroup_name.Text & "'," & Me.cbounder.ItemData(Me.cbounder.ListIndex) & ",'" & Me.cbonature.Text & "','" & Me.cbonature.Text & "','" & Me.cboaffectgp.Text & "')")
            temp_groupcode = temp_groupcode + 1
            Me.txtgroup_name.Text = ""
            MsgBox "A new Group has been Created", vbOKOnly
        End If
    End If
End Sub

Private Sub cbonature_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Me.cbonature.Text = "Income" Or Me.cbonature.Text = "Expence" Then
            Me.cboaffectgp.SetFocus
        Else
            ans = MsgBox("Save the Group?", vbYesNo)
            If ans = 6 Then
                db.Execute ("insert into groups (groupid,groupname,parentid,parentname,groupnature) values(" & temp_groupcode & ",'" & Me.txtgroup_name.Text & "'," & Me.cbounder.ItemData(Me.cbounder.ListIndex) & ",'" & Me.cbonature.Text & "','" & Me.cbonature.Text & "')")
                temp_groupcode = temp_groupcode + 1
                Me.txtgroup_name.Text = ""
                Me.txtgroup_name.SetFocus
                MsgBox "A new Group has been Created", vbOKOnly
            End If
        End If
    End If
End Sub

Private Sub cbounder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Me.cbounder.Text = "PRIMARY" Then
            Me.cbonature.SetFocus
        Else
        ans = MsgBox("Save the Group?", vbYesNo)
        If ans = 6 Then
            Set rs = db.OpenRecordset("select * from groups where groupid=" & Me.cbounder.ItemData(Me.cbounder.ListIndex))
            db.Execute ("insert into groups (groupid,groupname,parentid,parentname,groupnature) values(" & temp_groupcode & ",'" & Me.txtgroup_name.Text & "'," & Me.cbounder.ItemData(Me.cbounder.ListIndex) & ",'" & Me.cbounder.Text & "','" & rs("groupnature") & "')")
            temp_groupcode = temp_groupcode + 1
            Me.txtgroup_name.Text = ""
            Me.txtgroup_name.SetFocus
            MsgBox "A new Group has been Created", vbOKOnly
        End If
        End If
    End If
End Sub

Private Sub Form_Load()
    temp_groupcode = 0
    Me.cboaffectgp.AddItem "N"
    Me.cboaffectgp.AddItem "Y"
    Me.cboaffectgp.ListIndex = 0
    Set rs = db.OpenRecordset("select * from account_nature")
    While Not rs.EOF
        Me.cbonature.AddItem rs("accounttype")
        rs.MoveNext
    Wend
    If Me.cbonature.ListCount > 0 Then
        Me.cbonature.ListIndex = 0
    End If
    Set rs = db.OpenRecordset("select max(groupid) as max_group from groups")
    If Not IsNull(rs!max_group) Then
        temp_groupcode = rs!max_group + 1
    Else
        temp_groupcode = 1
    End If
    Me.cbounder.AddItem ("PRIMARY")
    Me.cbounder.ItemData(Me.cbounder.NewIndex) = 0
    Set rs = db.OpenRecordset("SELECT * FROM GROUPS")
    While Not rs.EOF
        Me.cbounder.AddItem (rs("GROUPNAME"))
        Me.cbounder.ItemData(Me.cbounder.NewIndex) = rs("GROUPID")
        rs.MoveNext
    Wend
    If Me.cbounder.ListCount > 0 Then
        Me.cbounder.ListIndex = 0
    End If
End Sub

Private Sub txtgroup_name_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbounder.SetFocus
    End If
End Sub
