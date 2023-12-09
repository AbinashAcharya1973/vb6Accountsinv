VERSION 5.00
Begin VB.Form frmledger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ledger"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   10695
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdnewledger 
         Caption         =   "Create New Ledger"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ListBox lstledger 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3030
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   4185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   3495
      Left            =   4590
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Label lblamount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   4305
      End
      Begin VB.Label lblunder 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   4545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Under :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3705
      End
   End
End
Attribute VB_Name = "frmledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnewledger_Click()
    frmNewLedger.Show 0
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Set rs = db.OpenRecordset("select * from ledgerMaster")
    While Not rs.EOF
        Me.lstledger.AddItem rs("accname")
        Me.lstledger.ItemData(Me.lstledger.NewIndex) = rs("accid")
        rs.MoveNext
    Wend
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub lstledger_Click()
    Set rs = db.OpenRecordset("select * from ledgermaster where accid=" & Me.lstledger.ItemData(Me.lstledger.ListIndex))
    If Not rs.EOF Then
        Me.lblunder.Caption = rs("groupname")
        Me.lblamount.Caption = Format(rs("obalance"), "##########0.00")
    End If
End Sub

Private Sub lstledger_KeyDown(KeyCode As Integer, Shift As Integer)
    Set rs = db.OpenRecordset("select * from ledgermaster where accid=" & Me.lstledger.ItemData(Me.lstledger.ListIndex))
    If Not rs.EOF Then
        Me.lblunder.Caption = rs("groupname")
        Me.lblamount.Caption = Format(rs("obalance"), "##########0.00")
    End If
End Sub

