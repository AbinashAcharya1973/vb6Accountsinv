VERSION 5.00
Begin VB.Form frmcreate_dsn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create DSN"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdcreate 
      Caption         =   "Create Data Link"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmcreate_dsn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim count1
Private Sub cmdcreate_Click()
CreateAccessODBC App.Path & "\mf.mdb", "NewMf", "RLFinance"
End Sub

Private Sub txtdb_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    Me.txtdescription.SetFocus
End If
End Sub

Private Sub txtdsn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    Me.txtdb.SetFocus
End If
End Sub

