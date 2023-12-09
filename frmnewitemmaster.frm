VERSION 5.00
Begin VB.Form frmnewitemmaster 
   BackColor       =   &H00008080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Item Master & Opening Stock Entry"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmnewitemmaster.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   7695
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txthsn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         TabIndex        =   37
         Tag             =   "0"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdimport 
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2610
         TabIndex        =   36
         Top             =   7200
         Width           =   1335
      End
      Begin VB.TextBox txtcolour 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Left            =   1890
         TabIndex        =   34
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox txtnewbrand 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtnewitemtype 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtnewcategory 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Import from Old Database"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   7320
         Width           =   2055
      End
      Begin VB.ComboBox cbotax 
         BackColor       =   &H8000000E&
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
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   4740
         Width           =   1215
      End
      Begin VB.ComboBox cbounit 
         BackColor       =   &H8000000E&
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4140
         Width           =   1215
      End
      Begin VB.ComboBox cboproducttype 
         BackColor       =   &H8000000E&
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
      Begin VB.ComboBox cboitemtype 
         BackColor       =   &H8000000E&
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
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Top             =   2040
         Width           =   4095
      End
      Begin VB.ComboBox cbobrandname 
         BackColor       =   &H8000000E&
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
      Begin VB.ComboBox cbosize 
         BackColor       =   &H8000000E&
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
         TabIndex        =   5
         Text            =   "cbosize"
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox txtlose 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Tag             =   "0"
         Text            =   "1"
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox txtmrp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   5340
         Width           =   1215
      End
      Begin VB.TextBox txttax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   5340
         Width           =   1215
      End
      Begin VB.TextBox txtpurchaserate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   5940
         Width           =   1215
      End
      Begin VB.TextBox txtsalerate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   5940
         Width           =   1215
      End
      Begin VB.TextBox txtopeningstock 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         TabIndex        =   12
         Text            =   "0"
         Top             =   6540
         Width           =   1215
      End
      Begin VB.TextBox txtbarcode 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         TabIndex        =   4
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "HSN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Colour"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   90
         TabIndex        =   35
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   4740
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label itemtype 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   4140
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Pack"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   4140
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   5340
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   5340
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   5940
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   5940
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   6600
         Width           =   1695
      End
   End
   Begin VB.Label lblmessage 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   7800
      Width           =   6495
   End
End
Attribute VB_Name = "frmnewitemmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, temp_barcode

Private Sub cbobrandname_GotFocus()
Me.lblmessage.Caption = "Press Esc to Add New Brand"
End Sub

Private Sub cbobrandname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtItemName.SetFocus
    End If
    If KeyCode = 27 Then
        Me.txtnewbrand.Left = Me.cbobrandname.Left
        Me.txtnewbrand.Width = Me.cbobrandname.Width
        Me.cbobrandname.Visible = False
        Me.txtnewbrand.Visible = True
        Me.txtnewbrand.SetFocus
    End If
    
End Sub


Private Sub cbobrandname_LostFocus()
Me.lblmessage.Caption = ""
End Sub

Private Sub cboitemtype_GotFocus()
Me.lblmessage.Caption = "Press Esc to Add New Item Type"
End Sub

Private Sub cboitemtype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbobrandname.SetFocus
    End If
    If KeyCode = 27 Then
        Me.txtnewitemtype.Left = Me.cboitemtype.Left
        Me.txtnewitemtype.Width = Me.cboitemtype.Width
        Me.cboitemtype.Visible = False
        Me.txtnewitemtype.Visible = True
        Me.txtnewitemtype.SetFocus
    End If
End Sub

Private Sub cboitemtype_LostFocus()
Me.lblmessage.Caption = ""
End Sub

Private Sub cboproducttype_Change()
    Set rec1 = db.OpenRecordset("select Item_Type from ItemType where ProductType='" & Me.cboproducttype.Text & "'")
    Me.cboitemtype.Clear
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboitemtype.AddItem (rec1("Item_Type"))
            rec1.MoveNext
        Wend
        If Me.cboitemtype.ListCount > 0 Then
            Me.cboitemtype.ListIndex = 0
        End If
    End If
End Sub
Private Sub cboproducttype_Click()
    cboproducttype_Change
End Sub

Private Sub cboproducttype_GotFocus()
Me.lblmessage.Caption = "Press Esc to Add New Cateogry"
End Sub

Private Sub cboproducttype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboitemtype.SetFocus
    End If
    If KeyCode = 27 Then
        Me.cboproducttype.Visible = False
        Me.txtnewcategory.Left = Me.cboproducttype.Left
        Me.txtnewcategory.Width = Me.cboproducttype.Width
        Me.txtnewcategory.Visible = True
        Me.txtnewcategory.SetFocus
        
    End If
End Sub

Private Sub cboproducttype_LostFocus()
Me.lblmessage.Caption = ""
End Sub

Private Sub cbosize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbounit.SetFocus
    End If
End Sub

Private Sub cbotax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Me.txttax.SetFocus
        Me.txthsn.SetFocus
    End If
End Sub

Private Sub cbounit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtlose.SetFocus
    End If
End Sub

Private Sub cmdimport_Click()
'frmimportitems.Show vbModal
End Sub

Private Sub Command1_Click()
    Dim tempfilename
    
    tempfilename = "j:\enlite\itemmaster1.xlsx"
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Double
    Total_qty = 0
    ans = MsgBox("Do You Want to Import Items", vbYesNo)
    
    If ans = vbYes Then
        '    Set rec1 = db.OpenRecordset("select * from headertemplate")
        '    If Not rec1.EOF Then
        '        tempfilename = rec1("PDTFileName")
        '    End If
            '----------------------------ItemMasterTemplate
            'ProductTypeCol = Val(rec1("ProductType_Col"))
            'ProductTypeCol2 = Val(rec1("ProductType_Col2"))
            'ItemTypeCol = Val(rec1("ItemType_Col"))
            'ItemTypeCol2 = Val(rec1("ItemType_Col2"))
            'BrandCol = Val(rec1("BrandName_Col"))
            ItemCol = 5
            'BarcodeCol = Val(rec1("BarCode_Col"))
            'SizeCol = Val(rec1("Size_Col"))
            'UnitTypeCol = Val(rec1("UnitType"))
            'MRPCol = Val(rec1("MRP_Col"))
            TaxCol = 8
            'RateCol = Val(rec1("Rate_Col"))
            'ColorCol = Val(rec1("TColor"))
            HSNCol = 7
            'QtyCol = Val(rec1("Qty_Col"))
        
        Set ExcelObj = CreateObject("Excel.Application")
        Set ExcelSheet = CreateObject("Excel.Sheet")
        Me.lblmessage.Caption = "Importing PDT file, Wait"
        ExcelObj.Workbooks.Open tempfilename

        Set ExcelBook = ExcelObj.Workbooks(1)
        Set ExcelSheet = ExcelBook.Worksheets(1)

        
        With ExcelSheet
            i = 2
            Do Until .Cells(i, 1) & "" = ""
                tempptype = .Cells(i, 1)
                Set rec1 = db.OpenRecordset("SELECT MAX(PRODUCTCODE) AS MAXCODE FROM ITEMMASTER")
                If Not IsNull(rec1!MAXCODE) Then
                    NEXTCODE = rec1!MAXCODE + 1
                Else
                    NEXTCODE = 1001
                End If
                db.Execute ("insert into ItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,hsn,PRODUCTCODE) values('" & .Cells(i, 1) & "','MEDICINE','" & .Cells(i, 4) & "','" & .Cells(i, 5) & "','" & .Cells(i, 5) & "','','1S',1,0," & .Cells(i, 8) & ",'" & .Cells(i, 7) & "'," & NEXTCODE & ")")
                db.Execute ("insert into STOCK (ProductType,ItemType,Brand,ItemNAME,Barcode,Size,UniYType,Lose,MRP,VAT,hsn,PRODUCTCODE) values('" & .Cells(i, 1) & "','MEDICINE','" & .Cells(i, 4) & "','" & .Cells(i, 5) & "','" & .Cells(i, 5) & "','','1S',1,0," & .Cells(i, 8) & ",'" & .Cells(i, 7) & "'," & NEXTCODE & ")")
                i = i + 1
                tempqty = 0
            Loop

        End With
        'Me.txtdisc.SetFocus
    Else
        'MsgBox "Missing", vbOKOnly
    End If
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Set rec1 = db.OpenRecordset("select * from UnitMaster")
    While Not rec1.EOF
        Me.cbounit.AddItem (rec1("UnitName"))
        rec1.MoveNext
    Wend
    If Me.cbounit.ListCount > 0 Then
        Me.cbounit.ListIndex = 0
    End If

    Me.cbotax.AddItem ("SALES")
    Me.cbotax.AddItem ("MRP")
    Me.cbotax.AddItem ("INCLUSIVE MRP")
    Me.cbotax.AddItem ("FREE")
    Me.cbotax.ListIndex = 0

    Set rec1 = db.OpenRecordset("select * from Product")
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboproducttype.AddItem (rec1("Productname"))
            Me.cboproducttype.ItemData(Me.cboproducttype.NewIndex) = rec1("Pid")
            rec1.MoveNext
        Wend
        If Me.cboproducttype.ListCount > 0 Then
            Me.cboproducttype.ListIndex = 0
        End If
    End If

    Set rec1 = db.OpenRecordset("select * from SizeMaster")
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cbosize.AddItem (rec1("size"))
            rec1.MoveNext
        Wend
        If Me.cbosize.ListCount > 0 Then
            Me.cbosize.ListIndex = 0
        End If
    End If

    Set rec1 = db.OpenRecordset("select * from Brandmaster")
    If Not rec1.EOF Then
        Me.cbobrandname.Clear
        While Not rec1.EOF
            Me.cbobrandname.AddItem (rec1("brand"))
            Me.cbobrandname.ItemData(Me.cbobrandname.NewIndex) = rec1("BrandId")
            rec1.MoveNext
        Wend
        If Me.cbobrandname.ListCount > 0 Then
            Me.cbobrandname.ListIndex = 0
        End If
    End If
If Me.NewBarcode <> "" Then
    Me.txtbarcode.Text = Me.NewBarcode
End If
Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
End Sub

Private Sub txtbarcode_GotFocus()
Me.txtbarcode.SelStart = 0
Me.txtbarcode.SelLength = Len(Me.txtbarcode.Text)
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Me.cbosize.SetFocus
        Me.txtcolour.SetFocus
    End If
End Sub

Private Sub txtcolour_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cbosize.SetFocus
End If
End Sub

Private Sub txthsn_GotFocus()
Me.txthsn.SelStart = 0
Me.txthsn.SelLength = Len(Me.txthsn.Text)
End Sub

Private Sub txthsn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txttax.SetFocus
End If
End Sub

Private Sub txtItemName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Me.txtbarcode.Text = "" Then
            Me.txtbarcode.Text = Me.txtItemName.Text
        End If
        Me.txtbarcode.SetFocus
    End If
End Sub
Private Sub txtItemName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtlose_GotFocus()
    Me.txtlose.SelStart = 0
    Me.txtlose.SelLength = Len(Me.txtlose.Text)
End Sub
Private Sub txtlose_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbotax.SetFocus
    End If
End Sub
Private Sub txtmrp_GotFocus()
    Me.txtmrp.SelStart = 0
    Me.txtmrp.SelLength = Len(Me.txtmrp.Text)
End Sub
Private Sub txtmrp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtpurchaserate.Text = Me.txtmrp.Text
        Me.txtsalerate.Text = Me.txtmrp.Text
        Me.txtpurchaserate.SetFocus
    End If
End Sub

Private Sub txtmrp_LostFocus()
    'Me.txtpurchaserate.Text = "0.00"
    Set rec1 = db.OpenRecordset("Select * from Brandmaster where brand='" & Me.cbobrandname.Text & "'")
    If Not rec1.EOF Then
        Me.txtpurchaserate.Text = Format(Val(Me.txtmrp.Text) * (Val(rec1("purchase") / 100)), "####0.00")
        Me.txtsalerate.Text = Format(Val(Me.txtmrp.Text) * (Val(rec1("Sale") / 100)), "######0.00")
    Else
        'Me.txtpurchaserate.Text = "0.00"
        'Me.txtsalerate.Text = "0.00"
    End If
End Sub

Private Sub txtnewbrand_GotFocus()
Me.lblmessage.Caption = "Enter to Accept / Esc to Cancel"
End Sub

Private Sub txtnewbrand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtnewbrand.Text = ""
    Me.cbobrandname.Visible = True
    Me.txtnewbrand.Visible = False
    Me.cbobrandname.SetFocus
End If
If KeyCode = 13 Then
            Set rec1 = db.OpenRecordset("select * from Brandmaster where brand='" & Me.txtnewbrand.Text & "'")
            If Not rec1.EOF Then
                MsgBox "Allready exists", vbCritical
            Else
                ans = MsgBox("Save This?", vbYesNo)
                If ans = 6 Then
                    db.Execute ("insert into Brandmaster (brand,Purchase,Sale) values('" & Me.txtnewbrand.Text & "',0, 0)")
                    Me.cbobrandname.AddItem Me.txtnewbrand.Text
                    Me.cbobrandname.ListIndex = Me.cbobrandname.NewIndex
                    Me.txtnewbrand.Text = ""
                    Me.cbobrandname.Visible = True
                    Me.txtnewbrand.Visible = False
                    Me.cbobrandname.SetFocus
                End If
                
            End If
        End If
End Sub

Private Sub txtnewbrand_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnewbrand_LostFocus()
Me.lblmessage.Caption = ""
End Sub

Private Sub txtnewcategory_GotFocus()
Me.lblmessage.Caption = "Enter to Accept / Esc to Cancel"
End Sub

Private Sub txtnewcategory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtnewcategory.Text = ""
    Me.txtnewcategory.Visible = False
    Me.cboproducttype.Visible = True
    Me.cboproducttype.SetFocus
End If
If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from Product where Productname='" & Me.txtnewcategory.Text & "'")
        If Not rec1.EOF Then
            MsgBox "Allready Exists", vbCritical
        Else
            ans = MsgBox("Save This?", vbYesNo)
            If ans = 6 Then
                db.Execute ("insert into Product (Productname) values('" & Me.txtnewcategory.Text & "')")
                Me.cboproducttype.AddItem Me.txtnewcategory.Text
                Me.cboproducttype.ListIndex = Me.cboproducttype.NewIndex
                Me.txtnewcategory.Text = ""
                Me.txtnewcategory.Visible = False
                Me.cboproducttype.Visible = True
                Me.cboproducttype.SetFocus
                
            End If
            
        End If
    End If
End Sub

Private Sub txtnewcategory_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnewcategory_LostFocus()
Me.lblmessage.Caption = ""
End Sub

Private Sub txtnewitemtype_GotFocus()
Me.lblmessage.Caption = "Enter to Accept / Esc to Cancel"
End Sub

Private Sub txtnewitemtype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtnewitemtype.Text = ""
    Me.cboitemtype.Visible = True
    Me.txtnewitemtype.Visible = False
    Me.cboitemtype.SetFocus
End If
If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from ItemType where producttype='" & Me.cboproducttype.Text & "' and Item_Type='" & Me.txtnewitemtype.Text & "'")
        If Not rec1.EOF Then
            MsgBox "Allready Exists?", vbCritical
        Else
            ans = MsgBox("Save This?", vbYesNo)
            If ans = 6 Then
                db.Execute ("insert into ItemType (Item_Type,ProductType) values('" & Me.txtnewitemtype.Text & "','" & Me.cboproducttype.Text & "')")
                Me.cboitemtype.AddItem Me.txtnewitemtype.Text
                Me.cboitemtype.ListIndex = Me.cboitemtype.NewIndex
                Me.txtnewitemtype.Text = ""
                Me.cboitemtype.Visible = True
                Me.txtnewitemtype.Visible = False
                Me.cboitemtype.SetFocus
            End If
        End If
        
    End If
End Sub

Private Sub txtnewitemtype_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnewitemtype_LostFocus()
Me.lblmessage.Caption = ""
End Sub

Private Sub txtopeningstock_GotFocus()
    Me.txtopeningstock.SelStart = 0
    Me.txtopeningstock.SelLength = Len(Me.txtopeningstock.Text)
End Sub
Private Sub txtOpeningStock_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errtrap
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from ItemMaster where Item='" & Replace(Me.txtItemName.Text, "'", "''") & "' and MRP=" & Val(Me.txtmrp.Text))
        If Not rec1.EOF Then
            MsgBox "Allready Exists", vbCritical
        Else
            ans = MsgBox("Save This?", vbYesNo)
            If ans = 6 Then
                Set rec1 = db.OpenRecordset("select max(productcode) as productid from ItemMaster")
                If Not IsNull(rec1!productid) Then
                    Productcode = rec1!productid + 1
                Else
                    Productcode = 1000
                End If
                db.Execute ("insert into ItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,Purchaserate,Salerate,Openingstock,ProductCode,tax_type,BrandId,CategoryId,hsn) values('" & Me.cboproducttype.Text & "','" & Me.cboitemtype.Text & "','" & Me.cbobrandname.Text & "','" & Replace(Me.txtItemName.Text, "'", "''") & "','" & Replace(Me.txtbarcode.Text, "'", "''") & "','" & Me.cbosize.Text & "','" & Me.cbounit.Text & "'," & Me.txtlose.Text & "," & Me.txtmrp.Text & "," & Me.txttax.Text & "," & Me.txtpurchaserate.Text & "," & Me.txtsalerate.Text & "," & Val(Me.txtopeningstock.Text) * Val(Me.txtlose.Text) & "," & Productcode & ",'" & Me.cbotax.Text & "'," & Me.cbobrandname.ItemData(Me.cbobrandname.ListIndex) & "," & Me.cboproducttype.ItemData(Me.cboproducttype.ListIndex) & ",'" & Me.txthsn.Text & "')")
                db.Execute ("insert into Stock (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,hsn) values('" & Me.cboproducttype.Text & "','" & Me.cboitemtype.Text & "','" & Replace(Me.txtItemName.Text, "'", "''") & "','" & Me.cbobrandname.Text & "','" & Replace(Me.txtbarcode.Text, "'", "''") & "','" & Me.cbosize.Text & "'," & Me.txtmrp.Text & "," & Me.txtpurchaserate.Text & "," & Val(Me.txtopeningstock.Text) * Val(Me.txtlose.Text) & "," & Productcode & "," & Val(Me.txttax.Text) & "," & Val(Me.txtlose.Text) & ",'" & Me.cbounit.Text & "'," & Val(Me.txtsalerate.Text) & ",'" & Me.txthsn.Text & "')")
                If Val(Me.txtopeningstock.Text) > 0 Then
                    db.Execute ("insert into Stockdetails (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,mfgdate,batchno,expdate,hsn) values('" & Me.cboproducttype.Text & "','" & Me.cboitemtype.Text & "','" & Replace(Me.txtItemName.Text, "'", "''") & "','" & Me.cbobrandname.Text & "','" & Replace(Me.txtbarcode.Text, "'", "''") & "','" & Me.cbosize.Text & "'," & Me.txtmrp.Text & "," & Me.txtpurchaserate.Text & "," & Me.txtopeningstock.Text & "," & Productcode & "," & Val(Me.txttax.Text) & "," & Val(Me.txtlose.Text) & ",'" & Me.cbounit.Text & "'," & Val(Me.txtsalerate.Text) & ",'" & Date & "','Opening Stock',' ','" & Me.txthsn.Text & "')")
                End If
                Me.txtbarcode.Text = ""
                Me.txtmrp.Text = "0.00"
                Me.txtlose.Text = 1
                Me.txtpurchaserate.Text = "0.00"
                Me.txtsalerate.Text = "0.00"
                Me.txtopeningstock.Text = 0
                If Me.NewBarcode = "" Then
                    Me.txtItemName.Text = ""
                    Productcode = ""
                    Me.cboproducttype.SetFocus
                Else
                    frmStockin.cboitemname.AddItem Me.txtItemName.Text
                    frmStockin.txtproductcode.Text = Productcode
                    Productcode = ""
                    Me.NewBarcode = ""
                    Unload Me
                End If
            End If
        End If
    End If
    Exit Sub
errtrap:
MsgBox Err.Description, vbCritical
End Sub

Private Sub txtpurchaserate_GotFocus()
    Me.txtpurchaserate.SelStart = 0
    Me.txtpurchaserate.SelLength = Len(Me.txtpurchaserate.Text)
End Sub

Private Sub txtpurchaserate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtsalerate.SetFocus
    End If
End Sub

Private Sub txtsalerate_GotFocus()
    Me.txtsalerate.SelStart = 0
    Me.txtsalerate.SelLength = Len(Me.txtsalerate.Text)
End Sub

Private Sub txtsalerate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtopeningstock.SetFocus
    End If
End Sub
Private Sub txttax_GotFocus()
    Me.txttax.SelStart = 0
    Me.txttax.SelLength = Len(Me.txttax.Text)
End Sub

Private Sub txttax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtmrp.SetFocus
    End If
End Sub

Public Property Get NewBarcode() As Variant
NewBarcode = temp_barcode
End Property

Public Property Let NewBarcode(ByVal vNewValue As Variant)
temp_barcode = vNewValue
End Property
