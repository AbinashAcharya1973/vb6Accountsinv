VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmbrandname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brand Name Entry"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmbrandname.frx":0000
   ScaleHeight     =   2700
   ScaleWidth      =   6405
   Begin VB.TextBox txtsale 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtpurchase 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmbrandname.frx":D4E3
      Height          =   2265
      Left            =   120
      OleObjectBlob   =   "frmbrandname.frx":D4F7
      TabIndex        =   1
      Top             =   270
      Width           =   6135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\FMCG\DATA\2011-2012\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Brandmaster"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtbrandname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmbrandname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
End Sub
Private Sub txtbrandname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtpurchase.SetFocus
    End If
End Sub
Private Sub txtbrandname_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtpurchase_GotFocus()
    Me.txtpurchase.SelStart = 0
    Me.txtpurchase.SelLength = Len(Me.txtpurchase.Text)
End Sub

Private Sub txtpurchase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtsale.SetFocus
    End If
End Sub

Private Sub txtsale_GotFocus()
    Me.txtsale.SelStart = 0
    Me.txtsale.SelLength = Len(Me.txtsale.Text)
End Sub

Private Sub txtsale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If KeyCode = 13 Then
            Set rec1 = db.OpenRecordset("select * from Brandmaster where brand='" & Me.txtbrandname.Text & "'")
            If Not rec1.EOF Then
                MsgBox "Allready exists", vbCritical
            Else
                ans = MsgBox("Save This?", vbYesNo)
                If ans = 6 Then
                    db.Execute ("insert into Brandmaster (brand,Purchase,Sale) values('" & Me.txtbrandname.Text & "'," & Val(Me.txtpurchase.Text) & "," & Val(Me.txtsale.Text) & ")")
                    Me.txtbrandname.Text = ""
                End If
                Data1.Refresh
            End If
        End If
    End If
End Sub
