VERSION 5.00
Begin VB.Form frmItemsmaster 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Master"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtItems 
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
         Left            =   840
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Items."
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
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmItemsmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub
Private Sub txtItems_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
ans = MsgBox("Save This?", vbYesNo)
If ans = 6 Then
Set rec1 = Db.OpenRecordset("select * from ItemMaster where  Items='" & Me.txtItems.Text & "'")
If Not rec1.EOF Then
MsgBox "AllReady Exist?", vbCritical
Else
Set rec2 = Db.OpenRecordset("select max(Slno) as max_no from ItemMaster")
If Not rec2.EOF Then
temp_slno = rec2!Max_no + 1
Else
temp_slno = 1
End If
Db.Execute ("insert into ItemMaster (Items,SlNo) values('" & Me.txtItems.Text & "'," & temp_slno & ")")
'    If formid = 1000 Then
'    frmShortmaster.cboItem.AddItem (Me.txtItems.Text)
'    frmShortmaster.cboItem.ListIndex = frmShortmaster.cboItem.ListCount - 1
'    Unload Me
'    End If
    
Me.txtItems.Text = ""
End If
End If
End If
End Sub
Private Sub txtItems_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
