VERSION 5.00
Begin VB.Form frmacclist 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ledger List"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstledger 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmacclist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Set rs = Db.OpenRecordset("select * from ledgerMaster")
While Not rs.EOF
    Me.lstledger.AddItem rs("accname")
    Me.lstledger.ItemData(Me.lstledger.NewIndex) = rs("accid")
    rs.MoveNext
Wend

End Sub

Private Sub lstledger_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    frmJournal.DBGrid1.Columns(1) = lstledger.Text
    Unload Me
End If
End Sub
