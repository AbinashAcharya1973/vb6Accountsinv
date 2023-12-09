VERSION 5.00
Begin VB.Form frmDebtorList 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sundry Debtor"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstdebtor 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5940
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmDebtorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec As Recordset
Private Sub Form_Load()
Set rec1 = Db.OpenRecordset("select * from LedgerMaster where Groupname like 'Sundry Debtor'")
While Not rec1.EOF
Me.lstdebtor.AddItem (rec1("AccName"))
Me.lstdebtor.ItemData(Me.lstdebtor.NewIndex) = rec1("AccID")
rec1.MoveNext
Wend
End Sub

