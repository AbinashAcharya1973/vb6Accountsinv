VERSION 5.00
Begin VB.Form frminstmsg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   615
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpwd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Email ID"
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
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frminstmsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_sRetValue As String, trail_c
Public Function GetPassword(Optional Prompt As String, Optional Title As String, Optional XPos As Variant, Optional YPos As Variant) As String
    Me.Caption = Title
    Label1.Caption = Prompt
    Me.txtpwd.PasswordChar = "*"
    
    If Not IsMissing(XPos) Then Me.Left = XPos
    If Not IsMissing(YPos) Then Me.Top = YPos
    
    Me.Show vbModal
    GetPassword = m_sRetValue
    Unload Me 'DO NOT REMOVE THIS 2nd UNLOAD
End Function
    

Private Sub txtpwd_GotFocus()
Me.txtpwd.SelStart = 0
Me.txtpwd.SelLength = Len(Me.txtpwd.Text)
End Sub


Private Sub txtpwd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    m_sRetValue = Me.txtpwd.Text
    Unload Me
End If
If KeyCode = 27 Then
    m_sRetValue = vbNullString
    Unload Me
End If
End Sub

Public Function GetEmailID(Optional Prompt As String, Optional Title As String, Optional dValue As Variant, Optional XPos As Variant, Optional YPos As Variant) As String
    Me.Caption = Title
    Label1.Caption = Prompt
    Me.txtpwd.Text = dValue
    
    If Not IsMissing(XPos) Then Me.Left = XPos
    If Not IsMissing(YPos) Then Me.Top = YPos
    
    Me.Show vbModal
    GetEmailID = m_sRetValue
    Unload Me 'DO NOT REMOVE THIS 2nd UNLOAD
End Function
