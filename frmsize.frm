VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmsize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Size Entry"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4920
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmsize.frx":0000
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmsize.frx":0014
      TabIndex        =   2
      Top             =   840
      Width           =   4575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SizeMaster"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtsize 
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
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Width           =   1575
   End
End
Attribute VB_Name = "frmsize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset
Private Sub txtbrandname_Change()

End Sub

Private Sub txtbrandname_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Me.Data1.databasename = dbname
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtsize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from SizeMaster where size='" & Me.txtsize.Text & "'")
        If Not rec1.EOF Then
            MsgBox "Allready Exists", vbCritical
        Else
            ans = MsgBox("Save This?", vbYesNo)
            If ans = 6 Then
                db.Execute ("insert into SizeMaster (size) values('" & Me.txtsize.Text & "')")
                Me.txtsize.Text = ""
            End If
        End If
    End If
End Sub
