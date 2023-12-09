VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmproducttype 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Type Entry"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5145
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
      RecordSource    =   "Product"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmproducttype.frx":0000
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmproducttype.frx":0014
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txtproducttype 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Product Type"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmproducttype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC1 As Recordset

Private Sub Form_Load()
    Me.Data1.databasename = dbname
End Sub

Private Sub txtproducttype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set REC1 = db.OpenRecordset("select * from Product where Productname='" & Me.txtproducttype.Text & "'")
        If Not REC1.EOF Then
            MsgBox "Allready Exists", vbCritical
        Else
            ans = MsgBox("Save This?", vbYesNo)
            If ans = 6 Then
                db.Execute ("insert into Product (Productname) values('" & Me.txtproducttype.Text & "')")
                Me.txtproducttype.Text = ""
            End If
            Me.Data1.Refresh
        End If
    End If

End Sub

Private Sub txtproducttype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
