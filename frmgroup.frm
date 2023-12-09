VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmgroup 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Groups"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   12345
   Begin VB.Frame Frame3 
      Caption         =   "View Groups"
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4635
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4935
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   8705
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Description"
      Height          =   4335
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   7335
      Begin VB.Frame Frame4 
         BackColor       =   &H80000009&
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   7095
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Under :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label lblacc_under 
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2280
            TabIndex        =   12
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Balance:"
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
            Left            =   240
            TabIndex        =   11
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label lblopening 
            BackStyle       =   0  'Transparent
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   10
            Top             =   1200
            Width           =   2055
         End
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   7440
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label lblunder 
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
         Left            =   2160
         TabIndex        =   6
         Top             =   540
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Group Under"
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton cmdnew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6000
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cbogroups 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Acc. Groups"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Data1 As DAO.Database, rec1 As DAO.Recordset, rec2 As Recordset, rec As Recordset
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec.VB_VarUserMemId = 1073938432

Private Sub cbogroups_Change()
    Set rs = db.OpenRecordset("select * from groups where groupname='" & Me.cbogroups.Text & "'")
    If Not rs.EOF Then
        Me.lblunder.Caption = (rs("parentname"))
    End If

End Sub

Private Sub cbogroups_Click()
    cbogroups_Change
End Sub

Private Sub cmdnew_Click()
    frmnewgroup.Show 0
End Sub

Private Sub CmdSave_Click()

End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Set rec1 = db.OpenRecordset("select groupname from groups")
    While Not rec1.EOF
        cbogroups.AddItem (rec1("groupname"))
        rec1.MoveNext
    Wend
    If cbogroups.ListCount > 0 Then
        cbogroups.ListIndex = 0
    End If


    Set rec1 = db.OpenRecordset("SELECT * FROM GROUPS WHERE PARENTID=0")
    While Not rec1.EOF
        Set xnode = TreeView1.Nodes.Add()
        xnode.Text = rec1("GROUPNAME")
        xnode.ForeColor = vbBlack
        xnode.Bold = True
        tempslno = xnode.index
        xnode.Tag = rec1("GROUPID")
        rec1.MoveNext
    Wend
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtgroupname_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.ForeColor = vbBlack Or Node.ForeColor = vbRed Then
        Set rec1 = db.OpenRecordset("SELECT * FROM GROUPS WHERE PARENTID=" & Node.Tag)
        If Not rec1.EOF Then
            While Not rec1.EOF
                Set xnode = TreeView1.Nodes.Add(Node.index, tvwChild)
                xnode.Text = rec1("GROUPNAME") & "  (Group)"
                xnode.ForeColor = vbRed
                xnode.Tag = rec1("GROUPID")
                tempslno = xnode.index
                rec1.MoveNext
            Wend
        Else
            Set rec1 = db.OpenRecordset("SELECT * FROM LEDGERMASTER WHERE GROUPID=" & Node.Tag)
            While Not rec1.EOF
                Set xnode = TreeView1.Nodes.Add(Node.index, tvwChild)
                xnode.Text = rec1("ACCNAME") & "  (Ledger)"
                xnode.ForeColor = vbBlue
                xnode.Tag = rec1("ACCID")
                tempslno = xnode.index
                rec1.MoveNext
            Wend
        End If
    End If
    If Node.ForeColor = vbGreen Then
        Set rs = db.OpenRecordset("select * from ledgermaster where accid=" & Node.Tag)
        If Not rs.EOF Then
            Me.lblacc_under.Caption = rs("groupname")
            Me.lblopening.Caption = Format(rs("obalance"), "##########0.00")
        End If
    End If
End Sub
