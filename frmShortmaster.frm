VERSION 5.00
Begin VB.Form frmShortmaster 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort Master - Opening Stock Entry"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.ComboBox cboItem 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtSortno 
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Item Name"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Sort No."
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
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmShortmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset

Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
frmItemsmaster.Show vbModal
End If
If KeyCode = 13 Then
Me.txtSortno.SetFocus
End If
End Sub



Private Sub CmdSave_Click()

End Sub

Private Sub Form_Load()
formid = 1000
Me.Top = 0
Me.Left = 0
Set rec1 = Db.OpenRecordset("select * from ItemMaster")
While Not rec1.EOF
Me.cboItem.AddItem (rec1("Items"))
rec1.MoveNext
Wend
If Me.cboItem.ListCount > 0 Then
Me.cboItem.ListIndex = 0
End If
End Sub
Private Sub txtOpeningStock_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'Me.CmdSave.SetFocus
'End If
End Sub
Private Sub txtSortNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Set rec1 = Db.OpenRecordset("select * from SortMaster where SortNo='" & Me.txtSortno.Text & "'")
If Not rec1.EOF Then
    MsgBox "allready exist", vbCritical
Else
ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then
    Db.Execute ("insert into SortMaster (SortNo,ItemName) values('" & Me.txtSortno.Text & "','" & Me.cboItem.Text & "')")
    Me.txtSortno.Text = ""
    Me.cboItem.SetFocus
    End If
End If
End If
End Sub
