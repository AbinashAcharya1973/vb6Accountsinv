VERSION 5.00
Begin VB.Form frmopeningstock 
   BackColor       =   &H0095B4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Entry or Opening Stock Entry"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   6480
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6255
      Begin VB.ComboBox cbounit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtbarcode 
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
         Left            =   1920
         TabIndex        =   25
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txtopeningstock 
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
         TabIndex        =   11
         Text            =   "0"
         Top             =   5640
         Width           =   1215
      End
      Begin VB.TextBox txtsalerate 
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
         Left            =   4560
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtpurchaserate 
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
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txttax 
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
         Left            =   4560
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtmrp 
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
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtlose 
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
         Left            =   4560
         TabIndex        =   5
         Text            =   "0"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.ComboBox cbosize 
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
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3240
         Width           =   3255
      End
      Begin VB.ComboBox cbobrandname 
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
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtItemName 
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
         TabIndex        =   2
         Top             =   2040
         Width           =   4095
      End
      Begin VB.ComboBox cboitemtype 
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
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   3255
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
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label13 
         Caption         =   "Opening Stock"
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
         Left            =   120
         TabIndex        =   24
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Sale Rate"
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
         Left            =   3480
         TabIndex        =   23
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Purchase Rate"
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
         Left            =   120
         TabIndex        =   22
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Tax %"
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
         Left            =   3480
         TabIndex        =   21
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "M.R.P"
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
         Left            =   120
         TabIndex        =   20
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Lose"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Unit Type"
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
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Size"
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
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "BarCode / Colour"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Brand Name"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Item Name"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label itemtype 
         Caption         =   "Item Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Product Type"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmopeningstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset, rec2 As DAO.Recordset
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.cbounit.AddItem ("BOX")
Me.cbounit.AddItem ("PAIR")
Me.cbounit.AddItem ("PCS")
Me.cbounit.ListIndex = 0

Set rec1 = Db.OpenRecordset("select * from Product")
If Not rec1.EOF Then
    While Not rec1.EOF
    Me.cboproducttype.AddItem (rec1("Productname"))
    rec1.MoveNext
    Wend
    If Me.cboproducttype.ListCount > 0 Then
    Me.cboproducttype.ListIndex = 0
    End If
End If

Set rec1 = Db.OpenRecordset("select * from ItemType")
If Not rec1.EOF Then
    While Not rec1.EOF
    Me.cboitemtype.AddItem (rec1("Item_Type"))
    rec1.MoveNext
    Wend
    If Me.cboitemtype.ListCount > 0 Then
    Me.cboitemtype.ListIndex = 0
    End If
End If

Set rec1 = Db.OpenRecordset("select * from Brandmaster")
If Not rec1.EOF Then
    While Not rec1.EOF
    Me.cbobrandname.AddItem (rec1("brand"))
    rec1.MoveNext
    Wend
    If Me.cbobrandname.ListCount > 0 Then
    Me.cbobrandname.ListIndex = 0
    End If
End If

Set rec1 = Db.OpenRecordset("select * from SizeMaster")
If Not rec1.EOF Then
    While Not rec1.EOF
    Me.cbosize.AddItem (rec1("size"))
    rec1.MoveNext
    Wend
    If Me.cbosize.ListCount > 0 Then
    Me.cbosize.ListIndex = 0
    End If
End If

End Sub

Private Sub txtItemName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtlose_Change()

End Sub

Private Sub txtopeningstock_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If KeyCode = 13 Then
    Set rec1 = Db.OpenRecordset("select * from ItemMaster where ProductType='" & Me.cboproducttype.Text & "' and ItemType='" & Me.cboitemtype.Text & "' and Brand='" & Me.cbobrandname.Text & "' and Item='" & Me.txtItemName.Text & "' and Barcode='" & Me.txtbarcode.Text & "' and Size='" & Me.cbosize.Text & "' and UnitType='" & Me.cbounit.Text & "' and Lose=" & Me.txtlose.Text & " and MRP=" & Me.txtmrp.Text & " and Tax=" & Me.txttax.Text & " and Purchaserate=" & Me.txtpurchaserate.Text & " and Salerate=" & Me.txtsalerate.Text)
    If Not rec1.EOF Then
    MsgBox "Allready Exists", vbCritical
    Else
        ans = MsgBox("Save This?", vbYesNo)
        If ans = 6 Then
        Db.Execute ("insert into ItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,Purchaserate,Salerate,Openingstock) values('" & Me.cboproducttype.Text & "','" & Me.cboitemtype.Text & "','" & Me.cbobrandname.Text & "','" & Me.txtItemName.Text & "','" & Me.txtbarcode.Text & "','" & Me.cbosize.Text & "','" & Me.cbounit.Text & "'," & Me.txtlose.Text & "," & Me.txtmrp.Text & "," & Me.txttax.Text & "," & Me.txtpurchaserate.Text & "," & Me.txtsalerate.Text & "," & Me.txtopeningstock.Text & ")")
        Me.txtItemName.Text = ""
        Me.txtbarcode.Text = ""
        Me.txtmrp.Text = "0.00"
        Me.txtlose.Text = 0
        Me.txtpurchaserate.Text = "0.00"
        Me.txtsalerate.Text = "0.00"
        Me.txtopeningstock.Text = 0
        Me.cboproducttype.SetFocus
        End If
    End If
End If

End If
End Sub
