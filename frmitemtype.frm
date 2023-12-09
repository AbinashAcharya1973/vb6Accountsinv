VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmitemtype 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Type Entry"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5070
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmitemtype.frx":0000
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "frmitemtype.frx":0014
      TabIndex        =   4
      Top             =   1560
      Width           =   4815
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
      RecordSource    =   "ItemType"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboproducttype 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox txtitemtype 
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
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label2 
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
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Item Type"
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
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmitemtype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Data1.databasename = dbname
    Set rec1 = db.OpenRecordset("select * from Product")
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboproducttype.AddItem (rec1("Productname"))
            rec1.MoveNext
        Wend
        If Me.cboproducttype.ListCount > 0 Then
            Me.cboproducttype.ListIndex = 0
        End If
    End If
End Sub
Private Sub txtitemtype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from ItemType where producttype='" & Me.cboproducttype.Text & "' and Item_Type='" & Me.txtitemtype.Text & "'")
        If Not rec1.EOF Then
            MsgBox "Allready Exists?", vbCritical
        Else
            ans = MsgBox("Save This?", vbYesNo)
            If ans = 6 Then
                db.Execute ("insert into ItemType (Item_Type,ProductType) values('" & Me.txtitemtype.Text & "','" & Me.cboproducttype.Text & "')")
                Me.txtitemtype.Text = ""
            End If
        End If
    End If
End Sub
Private Sub txtitemtype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

