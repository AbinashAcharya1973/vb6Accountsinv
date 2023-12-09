VERSION 5.00
Begin VB.Form frmShadeMaster 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shade Master"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtShades 
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
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Shade No"
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
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmShadeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset

Private Sub Form_Load()
Me.Top = 3375
Me.Left = 4000
End Sub
Private Sub txtShades_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
If KeyCode = 13 Then
If Me.txtShades.Text <> "" Then
Set rec1 = Db.OpenRecordset("select * from ShadeMaster where Shades='" & Me.txtShades.Text & "'")
    If Not rec1.EOF Then
    MsgBox "Allready Exist?", vbCritical
    Else
    Db.Execute ("insert into ShadeMaster (Shades) values('" & Me.txtShades.Text & "')")
    Me.txtShades.Text = ""
    End If
        If formname = "stockin" Then
        frmStockin.txtShade.Text = Me.txtShades.Text
        Unload Me
        End If
'If formid = 1000 Then
'frmShortmaster.txtShadesNo.Text = Me.txtShades.Text
'Unload Me
'End If
'Me.txtShades.Text = ""
End If
End If
End Sub

Private Sub Text1_Change()

End Sub
