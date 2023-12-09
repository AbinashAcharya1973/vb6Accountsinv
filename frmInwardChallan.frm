VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInwardChallan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inward Challan"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstinv 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   1200
      TabIndex        =   103
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Frame frnewitem 
      BackColor       =   &H8000000D&
      Caption         =   "Item Name"
      Height          =   1335
      Left            =   0
      TabIndex        =   87
      Top             =   4320
      Visible         =   0   'False
      Width           =   14295
      Begin VB.TextBox txttaxp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11400
         TabIndex        =   93
         Text            =   "0"
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtnewbrand 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8700
         TabIndex        =   100
         Top             =   540
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtnewmtype 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5460
         TabIndex        =   99
         Top             =   540
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txthsn 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10440
         TabIndex        =   92
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtiname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   91
         Top             =   540
         Width           =   3375
      End
      Begin VB.ComboBox cbobrand 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   540
         Width           =   2220
      End
      Begin VB.ComboBox cbomtype 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   540
         Width           =   2220
      End
      Begin VB.ComboBox cbocategory 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   540
         Width           =   2220
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   11400
         TabIndex        =   101
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "HSN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   10440
         TabIndex        =   98
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   6960
         TabIndex        =   97
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   4680
         TabIndex        =   96
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   95
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   60
         TabIndex        =   94
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   45
      Top             =   1680
      Width           =   14190
      Begin VB.ComboBox cbobatch 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   86
         Text            =   "cbobatch"
         Top             =   840
         Width           =   3360
      End
      Begin VB.TextBox txtcd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         TabIndex        =   79
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txttaxamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12285
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtnet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13125
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtspecialdiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10215
         TabIndex        =   13
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txttaxtype 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtdiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9615
         TabIndex        =   12
         Text            =   "0"
         Top             =   840
         Width           =   480
      End
      Begin VB.TextBox txtfree 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6375
         TabIndex        =   8
         Text            =   "0"
         Top             =   840
         Width           =   480
      End
      Begin VB.TextBox TxtVat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11640
         TabIndex        =   14
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtproductcode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   15
         TabIndex        =   54
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtpack 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4455
         TabIndex        =   5
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8655
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPrate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6975
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   840
         Width           =   720
      End
      Begin VB.TextBox txtmrp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7815
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cbounit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5775
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Top             =   840
         Width           =   480
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5055
         TabIndex        =   6
         Text            =   "0"
         Top             =   840
         Width           =   600
      End
      Begin VB.ComboBox cboitemname 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   615
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   3840
      End
      Begin MSMask.MaskEdBox txtexpdate 
         Height          =   315
         Left            =   7140
         TabIndex        =   84
         Top             =   300
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtmfgdate 
         Height          =   315
         Left            =   3420
         TabIndex        =   85
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp. Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6900
         TabIndex        =   83
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Mfg. Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3420
         TabIndex        =   82
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "SERIAL NO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   81
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label34 
         Caption         =   "QD"
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
         Left            =   10920
         TabIndex        =   80
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label31 
         Caption         =   "Tax Amt"
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
         Left            =   12285
         TabIndex        =   70
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Code"
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
         Left            =   15
         TabIndex        =   69
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "Net Amt"
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
         Left            =   13125
         TabIndex        =   68
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "S.Dis %"
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
         Left            =   10215
         TabIndex        =   67
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Free"
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
         Left            =   6375
         TabIndex        =   66
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Dis%"
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
         Left            =   9615
         TabIndex        =   60
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "GST %"
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
         Left            =   11685
         TabIndex        =   55
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Pack"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   52
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
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
         Left            =   8655
         TabIndex        =   51
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6975
         TabIndex        =   50
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "M.r.p"
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
         Left            =   7815
         TabIndex        =   49
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5775
         TabIndex        =   48
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   5055
         TabIndex        =   47
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   615
         TabIndex        =   46
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   14115
      Begin VB.TextBox txtstate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11280
         TabIndex        =   78
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtstatecode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13200
         TabIndex        =   77
         Top             =   600
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtinvdate 
         Height          =   315
         Left            =   6960
         TabIndex        =   2
         Top             =   1020
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtwaybill 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11760
         TabIndex        =   65
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtPartyCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10320
         TabIndex        =   37
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtLrNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11760
         TabIndex        =   17
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txtInvno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Text            =   "0"
         Top             =   1020
         Width           =   4335
      End
      Begin VB.TextBox txtaddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   36
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox cboSupplier 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   4335
      End
      Begin VB.TextBox txtslno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   240
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtStockindate 
         Height          =   315
         Left            =   6960
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 to Pull Online Invoices"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1080
         TabIndex        =   104
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label16 
         Caption         =   "Way Bill"
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
         Left            =   10800
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "L R No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         TabIndex        =   44
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   6000
         TabIndex        =   41
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sl No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\FMCG\DATA\2010-2011\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Temp_Stockin"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   14115
      Begin VB.TextBox txtetamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtcst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtcstamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "0.00"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtdiscountamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0.00"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtmrpamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   59
         Text            =   "0.00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtfreight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12360
         TabIndex        =   56
         Text            =   "0.00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080FF80&
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox txtGrandtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H0080FFFF&
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1320
         TabIndex        =   25
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox TxtNetAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12360
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtRoundup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12360
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txttotalgross 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtVatAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtTotalqty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H000000FF&
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2520
         TabIndex        =   19
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Label Label33 
         Caption         =   "Entry Tax"
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
         Left            =   7800
         TabIndex        =   73
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "C.s.t"
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
         Left            =   7800
         TabIndex        =   71
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   61
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Mrp Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   58
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Add Freight"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   57
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Total Qty"
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
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Grand Total"
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
         Left            =   11040
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "Total Gross"
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
         Left            =   7800
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "GST Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   30
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Net Amount"
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
         Left            =   11040
         TabIndex        =   29
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "Round Up"
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
         Left            =   11040
         TabIndex        =   28
         Top             =   960
         Width           =   1695
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmInwardChallan.frx":0000
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "frmInwardChallan.frx":0014
      TabIndex        =   53
      Top             =   3000
      Width           =   14115
   End
   Begin VB.Label lblmessage 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
      TabIndex        =   102
      Top             =   8760
      Width           =   14115
   End
End
Attribute VB_Name = "frmInwardChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinHttpReq As WinHttp.WinHttpRequest
      Private Const LOCALE_SSHORTDATE = &H1F
      Private Const WM_SETTINGCHANGE = &H1A
      'same as the old WM_WININICHANGE
      Private Const HWND_BROADCAST = &HFFFF&

      Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
          "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
          Long, ByVal lpLCData As String) As Boolean
      Private Declare Function PostMessage Lib "user32" Alias _
          "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function GetSystemDefaultLCID Lib "kernel32" _
          () As Long
Dim p As Object, JSONRec As Object, OnlineInv As Boolean, OnlineInvSlno, JSONInvDetails As Object
Dim rec As Recordset, rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset, rec5 As Recordset, TAXAMOUNT, temp_lr_balance, temp_lr_slno, TEMP_GROUPID, temp_discount_amount, PURCHASEDELETE
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432
Attribute rec5.VB_VarUserMemId = 1073938432
Attribute TAXAMOUNT.VB_VarUserMemId = 1073938432
Attribute temp_lr_balance.VB_VarUserMemId = 1073938432
Attribute temp_lr_slno.VB_VarUserMemId = 1073938432
Attribute TEMP_GROUPID.VB_VarUserMemId = 1073938432
Attribute temp_discount_amount.VB_VarUserMemId = 1073938432
Attribute PURCHASEDELETE.VB_VarUserMemId = 1073938432
Private Sub cboRate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtQty.SetFocus
    End If
End Sub

Private Sub cbobatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtmfgdate.SetFocus
    End If
End Sub

Private Sub cbobrand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtnewbrand.Top = Me.cbobrand.Top
    Me.txtnewbrand.Left = Me.cbobrand.Left
    Me.txtnewbrand.Width = Me.cbobrand.Width
    Me.txtnewbrand.Visible = True
    Me.txtnewbrand.SetFocus
End If
If KeyCode = 13 Then
    Me.txtiname.SetFocus
End If
End Sub

Private Sub cbocategory_Change()
    Set rec1 = db.OpenRecordset("select Item_Type from ItemType where ProductType='" & Me.cbocategory.Text & "'")
    Me.cbomtype.Clear
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cbomtype.AddItem (rec1("Item_Type"))
            rec1.MoveNext
        Wend
        If Me.cbomtype.ListCount > 0 Then
            Me.cbomtype.ListIndex = 0
        End If
    End If

End Sub

Private Sub cbocategory_Click()
cbocategory_Change
End Sub

Private Sub cboitemname_Change()
    txtproductcode.Text = Me.cboitemname.ItemData(Me.cboitemname.ListIndex)
    Set rec1 = db.OpenRecordset("SELECT * FROM ITEMMASTER WHERE PRODUCTCODE=" & Me.txtproductcode.Text)
    If Not rec1.EOF Then
        txtmrp.Text = rec1("MRP")
        'txtsalerate.Text = REC1("SALERATE")
        Me.txtPrate.Text = rec1("PURCHASERATE")
        TxtVat.Text = rec1("TAX")
        cbounit.Text = rec1("UNITTYPE")
        txttaxtype.Text = rec1("TAX_TYPE")
        '        Set rec2 = db.OpenRecordset("SELECT QTY FROM STOCK WHERE PRODUCTCODE=" & Me.txtproductcode.Text)
        '        If Not rec2.EOF Then
        '            Me.TXTSTOCK.Text = "Stock:: " & rec2("qty")
        '        End If
'        Set rec1 = db.OpenRecordset("select * from stockdetails where productcode=" & Me.txtproductcode.Text)
'        If Me.cbobatch.ListCount > 0 Then
'            Me.cbobatch.Clear
'        End If
'        While Not rec1.EOF
'            Me.cbobatch.AddItem rec1("batchno")
'            rec1.MoveNext
'        Wend
'        If Me.cbobatch.ListCount > 0 Then
'            Me.cbobatch.ListIndex = 0
'        End If
    End If
End Sub

Private Sub cboitemname_Click()
cboitemname_Change
End Sub

Private Sub cboitemname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'frmproductlist.Show vbModal
        Me.cbobatch.SetFocus
    End If
    If KeyCode = 27 Then
        Me.txtcst.SetFocus
    End If
    If KeyCode = vbKeyF1 Then
        Me.cboitemname.Clear
        Set rec = db.OpenRecordset("select * from stock")
        While Not rec.EOF
            Me.cboitemname.AddItem rec("itemname")
            Me.cboitemname.ItemData(Me.cboitemname.NewIndex) = rec("productcode")
            rec.MoveNext
        Wend
        If Me.cboitemname.ListCount > 0 Then
            Me.cboitemname.ListIndex = 0
        End If
    End If
    If KeyCode = vbKeyF2 Then
        Me.frnewitem.Top = Me.Frame2.Top + 50
        Me.frnewitem.Left = Frame2.Left
        'Me.frnewitem.ZOrder (0)
        'Frame2.ZOrder (1)
        Me.frnewitem.Visible = True
        Me.cbomtype.SetFocus
    End If
End Sub

Private Sub cboitemname_LostFocus()
    Set rec2 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
    If Not rec2.EOF Then
        TEMP_GROUPID = rec2("GroupId")
    End If
End Sub

Private Sub cbomtype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtnewmtype.Top = Me.cbomtype.Top
    Me.txtnewmtype.Left = Me.cbomtype.Left
    Me.txtnewmtype.Visible = True
    Me.txtnewmtype.Width = Me.cbomtype.Width
    Me.txtnewmtype.SetFocus
End If
If KeyCode = 13 Then
    Me.cbobrand.SetFocus
End If
If KeyCode = vbKeyF2 Then
    Me.frnewitem.Visible = False
    Me.cboitemname.SetFocus
End If
End Sub

Private Sub cboSupplier_Change()
    Set rec1 = db.OpenRecordset("select * from PartyCr where AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
    If Not rec1.EOF Then
        Me.txtaddress.Text = rec1("Address")
        Me.txtPartyCode.Text = Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex)
        If Not IsNull(rec1("statecode")) Then
            Set rec2 = db.OpenRecordset("select * from statecode where stcode=" & rec1("statecode"))
            If Not rec2.EOF Then
                Me.txtstate.Text = rec2("statename")
                Me.txtstatecode.Text = rec2("stcode")
            End If
        Else
            Me.txtstate.Text = ""
            Me.txtstatecode.Text = ""
        End If
    End If
    Set rec2 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
    If Not rec2.EOF Then
        TEMP_GROUPID = rec2("GroupId")
    End If
    
End Sub
Private Sub cboSupplier_Click()
    cboSupplier_Change
End Sub

Private Sub cboSupplier_GotFocus()
Me.lblmessage.Caption = "Press Esc to Add New Supplier"
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec2 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
        If Not rec2.EOF Then
            TEMP_GROUPID = rec2("GroupId")
        End If
        Me.txtInvno.SetFocus
        
    End If
    If KeyCode = 27 Then
        frmPartyCr.Show 0
    End If
End Sub
Private Sub cbounit_type_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtQty.SetFocus
    End If
End Sub

Private Sub cboSupplier_LostFocus()
On Error Resume Next
    Set rec2 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
    If Not rec2.EOF Then
        TEMP_GROUPID = rec2("GroupId")
    End If
    Me.lblmessage.Caption = ""
End Sub




Private Sub CmdDelete_Click()
    Me.txtslno.Locked = False
    PURCHASEDELETE = "Y"
    db.Execute ("delete * from Temp_Stockin")
    Data1.Refresh
    Me.txtInvno.Text = ""

    Me.txtTotalqty.Text = "0.00"
    Me.txttotalgross.Text = "0.00"
        Me.TxtVat.Text = "0.00"
    Me.TxtVatAmount.Text = "0.00"

    Me.TxtNetAmount.Text = "0.00"
    Me.TxtRoundup.Text = "0.00"

    Me.txtGrandtotal.Text = "0.00"
    Me.txtslno.SetFocus

End Sub
Private Sub CmdEdit_Click()
    Me.txtslno.Locked = False
    db.Execute ("delete * from Temp_Stockin")
    Data1.Refresh
    Me.txtInvno.Text = ""

    Me.txtTotalqty.Text = "0.00"
    Me.txttotalgross.Text = "0.00"
    
    Me.TxtVat.Text = "0.00"
    Me.TxtVatAmount.Text = "0.00"

    Me.TxtNetAmount.Text = "0.00"
    Me.TxtRoundup.Text = "0.00"

    Me.txtGrandtotal.Text = "0.00"
    Me.txtslno.SetFocus
End Sub
Private Sub CmdSave_Click()
    ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then

        If Val(Me.txtslno.Text) > 0 Then
            Slno = Val(Me.txtslno.Text)
        End If
        Set rec = db.OpenRecordset("select * from Temp_Stockin")
        If Not rec.EOF Then
            temp_grandtotal = Me.txtGrandtotal.Text
            X = NumberToWord(Trim(str(Round(Val(temp_grandtotal)))))
            Set rec1 = db.OpenRecordset("select * from Purchasehead where Slno=" & Slno)
            If rec1.EOF Then
                AccId = 0
                db.Execute ("insert into PurchaseHead (Slno,Purchasedate,InvNo,InvDate,Supplier,AccId,LrNo,TotalQty,TotalGross,VatAmount,NetAmount,RValue,GrandTotal,AmountInText,Freight,TotalMrp,lessDiscount,Waybill,CST,CSTAmount,ETax,ETaxAmount) values(" & Slno & ",'" & Me.txtStockindate.Text & "','" & Me.txtInvno.Text & "','" & Me.txtinvdate.Text & "','" & Me.cboSupplier.Text & "'," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ",'" & Me.txtLrNo.Text & "'," & Me.txtTotalqty.Text & "," & Me.txttotalgross.Text & "," & Me.TxtVatAmount.Text & "," & Me.TxtNetAmount.Text & "," & Me.TxtRoundup.Text & "," & Me.txtGrandtotal.Text & ",'" & X & "'," & Val(Me.txtfreight.Text) & "," & Val(Me.txtmrpamount.Text) & "," & Val(Me.txtdiscountamount.Text) & ",'" & Me.txtwaybill.Text & "'," & Me.txtcst.Text & "," & Me.txtcstamount.Text & "," & Me.txtet.Text & "," & Me.txtetamount.Text & ")")
            Else
                AccId = rec1("AccId")
                db.Execute ("update PurchaseHead set Purchasedate='" & Me.txtStockindate.Text & "',InvNo='" & Me.txtInvno.Text & "',InvDate='" & Me.txtinvdate.Text & "',Supplier='" & Me.cboSupplier.Text & "',AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ",LrNo='" & Me.txtLrNo.Text & "',TotalQty=" & Me.txtTotalqty.Text & ",TotalGross=" & Me.txttotalgross.Text & ",VatAmount=" & Me.TxtVatAmount.Text & ",NetAmount=" & Me.TxtNetAmount.Text & ",RValue=" & Me.TxtRoundup.Text & ",GrandTotal=" & Me.txtGrandtotal.Text & ",AmountInText='" & X & "',Freight=" & Val(Me.txtfreight.Text) & ",TotalMrp=" & Val(Me.txtmrpamount.Text) & ",lessDiscount=" & Val(Me.txtdiscountamount.Text) & ",Waybill='" & Me.txtwaybill.Text & "',CST=" & Me.txtcst.Text & ",CSTAmount=" & Me.txtcstamount.Text & ",ETax=" & Me.txtet.Text & ",ETaxAmount=" & Me.txtetamount.Text & " where Slno=" & Slno)
            End If

            Set rec2 = db.OpenRecordset("select * from PurchaseDetails where Slno=" & Slno)
            If Not rec2.EOF Then
                While Not rec2.EOF
                    Set rec3 = db.OpenRecordset("select * from Stock where productcode=" & rec2("Productcode"))
                    If Not rec3.EOF Then
                        db.Execute ("update stock set Qty=Qty - " & (rec2("Qty") + rec2("Free_Qty")) & " where productcode=" & rec2("Productcode"))
                        db.Execute ("update stockdetails set Qty=Qty - " & (rec2("Qty") + rec2("Free_Qty")) & " where productcode=" & rec2("Productcode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                    End If
                    rec2.MoveNext
                Wend
                db.Execute ("delete * from PurchaseDetails where SlNo=" & Slno)
            End If

            'Set rec2 = db.OpenRecordset("select * from Temp_Stockin where Usname='" & usname & "'")
            Set rec2 = db.OpenRecordset("select * from Temp_Stockin")
            If Not rec2.EOF Then
                While Not rec2.EOF
                    db.Execute ("insert into PurchaseDetails (Slno,ItemName,Units,Pack,Qty,MRP,PrRate,Amount,ProductCode,Vat,VatAmount,Net,Free_Qty,Discount,SpDiscount,Discount_amount,cd,mfgdate,expdate,batchno) values(" & Slno & ",'" & rec2("ItemName") & "','" & rec2("Units") & "'," & rec2("Pack") & "," & rec2("Qty") & "," & rec2("MRP") & "," & rec2("PrRate") & "," & rec2("Amount") & "," & rec2("ProductCode") & "," & rec2("Vat") & "," & rec2("VatAmount") & "," & rec2("Net") & "," & rec2("Free_Qty") & "," & rec2("Discount") & "," & rec2("SpDiscount") & "," & rec2("Discount_amount") & "," & rec2("cd") & ",'" & rec2("mfgdate") & "','" & rec2("expdate") & "','" & rec2("batchno") & "')")
                    Set rec3 = db.OpenRecordset("select * from Stock where ProductCode=" & rec2("ProductCode"))
                    If Not rec3.EOF Then
                        db.Execute ("update stock set Qty=Qty + " & (rec2("Qty") + rec2("Free_Qty")) & " where ProductCode=" & rec2("ProductCode"))
                    Else
                        db.Execute ("insert into Stock (itemname,MRP,PRate,Qty,ProductCode) values('" & rec2("itemname") & "'," & rec2("MRP") & "," & rec2("PrRate") & "," & (rec2("Qty") + rec2("Free_Qty")) & "," & rec2("ProductCode") & ")")
                    End If
                    Set rec3 = db.OpenRecordset("select * from Stockdetails where productcode=" & rec2("Productcode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                    If Not rec3.EOF Then
                        db.Execute ("update stockdetails set Qty=Qty + " & (rec2("Qty") + rec2("Free_Qty")) & " where productcode=" & rec2("Productcode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                    Else
                        Set rec1 = db.OpenRecordset("select * from itemmaster where productcode=" & rec2("productcode"))
                        If Not rec1.EOF Then
                            db.Execute ("insert into Stockdetails (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,mfgdate,batchno,expdate,hsn) values('" & rec1("ProductType") & "','" & rec1("itemtype") & "','" & rec1("item") & "','" & rec1("brand") & "','" & rec1("barcode") & "','" & rec1("size") & "'," & rec1("mrp") & "," & rec1("purchaserate") & "," & rec2("qty") + rec2("free_qty") & "," & rec2("Productcode") & "," & rec2("vat") & "," & rec1("Lose") & ",'" & rec1("unittype") & "'," & rec1("SaleRate") & ",'" & rec2("mfgdate") & "','" & rec2("batchno") & "','" & rec2("expdate") & "','" & rec1("HSN") & "')")
                        End If
                    End If
                    rec2.MoveNext
                Wend

                '               db.Execute ("delete * from LedgerTran Where VoucherType='Purchase' and VoucherSlno=" & Me.txtslno.Text)
                '                '--------Party Ledger New Entry--------------------------
                '                Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccName like 'Purchase*'")
                '                If Not rec3.EOF Then
                '                    prAccId = rec3("AccId")
                '                End If
                '                Set rec3 = db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & " and Slno=(select max(SlNo) from LedgerTran where AccId=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                '                If Not rec3.EOF Then
                '                    temp_ledger_balance = rec3("Balance")
                '                    temp_ledger_slno = rec3("Slno") + 1
                '                Else
                '                    temp_ledger_slno = 1
                '                    temp_ledger_balance = 0
                '                End If
                '                db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values (" & temp_ledger_slno & ",'" & Me.txtStockindate.Text & "','By Purchase',0," & temp_grandtotal & "," & temp_ledger_balance + temp_grandtotal & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ",'Inv No:" & Me.txtInvno.Text & "-" & Me.txtinvdate.Text & "','Purchase'," & Slno & "," & TEMP_GROUPID & "," & prAccId & ")")
                '                '-------Purchase Ledger New Entry---------------------------
                '                Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccName like 'Purchase*'")
                '                If Not rec3.EOF Then
                '                    temp_purchaseamount = Val(Me.txtGrandtotal.Text) - Val(Me.TxtVatAmount.Text)
                '                    Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec3("AccId") & " and slno=(select max(slno) from LedgerTran where AccId=" & rec3("AccId") & ")")
                '                    If Not rec4.EOF Then
                '                        TEMP_PURCHASE_LEDGER_SLNO = rec4("Slno") + 1
                '                        TEMP_PURCHASE_LEDGER_BALANCE = rec4("Balance")
                '                    Else
                '                        TEMP_PURCHASE_LEDGER_SLNO = 1
                '                        TEMP_PURCHASE_LEDGER_BALANCE = 0
                '                    End If
                '                    db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & TEMP_PURCHASE_LEDGER_SLNO & ",'" & Me.txtStockindate.Text & "','" & Me.cboSupplier.Text & "'," & temp_purchaseamount & ",0," & TEMP_PURCHASE_LEDGER_BALANCE + temp_purchaseamount & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvno.Text & "-" & Me.txtinvdate.Text & "','Purchase'," & Slno & "," & rec3("GroupId") & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                '                End If
                '                '--------New Transaction Entry Of Vat Ledger-----------------
                '                Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccName like 'Vat*'")
                '                If Not rec3.EOF Then
                '                    If Val(Me.TxtVatAmount.Text) > 0 Then
                '                        Set rec4 = db.OpenRecordset("select * from  LedgerTran where AccId=" & rec3("AccId") & " and SlNo=(select max(SlNo) from LedgerTran where AccId=" & rec3("AccId") & ")")
                '                        If Not rec4.EOF Then
                '                            TAX_LEDGER_SLNO = rec4("Slno") + 1
                '                            TAX_LEDGER_BALANCE = rec4("Balance")
                '                        Else
                '                            TAX_LEDGER_SLNO = 1
                '                            TAX_LEDGER_BALANCE = 0
                '                        End If
                '                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values (" & TAX_LEDGER_SLNO & ",'" & Me.txtStockindate.Text & "','" & Me.cboSupplier.Text & "'," & Me.TxtVatAmount.Text & ",0," & TAX_LEDGER_BALANCE - Val(Me.TxtVatAmount.Text) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','Purchase'," & Slno & "," & rec3("GroupId") & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                '                    End If
                '                End If

                '--------New Transaction Entry Of SGST Ledger-----------------
                '                If Val(Me.txtstatecode.Text) = 21 Then
                '                Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccName like 'SGST'")
                '                If Not rec3.EOF Then
                '                    If Val(Me.TxtVatAmount.Text) > 0 Then
                '                        Set rec4 = db.OpenRecordset("select * from  LedgerTran where AccId=" & rec3("AccId") & " and SlNo=(select max(SlNo) from LedgerTran where AccId=" & rec3("AccId") & ")")
                '                        If Not rec4.EOF Then
                '                            TAX_LEDGER_SLNO = rec4("Slno") + 1
                '                            TAX_LEDGER_BALANCE = rec4("Balance")
                '                        Else
                '                            TAX_LEDGER_SLNO = 1
                '                            TAX_LEDGER_BALANCE = 0
                '                        End If
                '                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values (" & TAX_LEDGER_SLNO & ",'" & Me.txtStockindate.Text & "','" & Me.cboSupplier.Text & "'," & Val(Me.TxtVatAmount.Text) / 2 & ",0," & TAX_LEDGER_BALANCE - (Val(Me.TxtVatAmount.Text) / 2) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvno.Text & "-" & Me.txtinvdate.Text & "','Purchase'," & Slno & "," & rec3("GroupId") & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                '                    End If
                '                End If
                '                End If
                '
                '                '--------New Transaction Entry Of CGST Ledger-----------------
                '                If Val(Me.txtstatecode.Text) = 21 Then
                '                Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccName like 'CGST'")
                '                If Not rec3.EOF Then
                '                    If Val(Me.TxtVatAmount.Text) > 0 Then
                '                        Set rec4 = db.OpenRecordset("select * from  LedgerTran where AccId=" & rec3("AccId") & " and SlNo=(select max(SlNo) from LedgerTran where AccId=" & rec3("AccId") & ")")
                '                        If Not rec4.EOF Then
                '                            TAX_LEDGER_SLNO = rec4("Slno") + 1
                '                            TAX_LEDGER_BALANCE = rec4("Balance")
                '                        Else
                '                            TAX_LEDGER_SLNO = 1
                '                            TAX_LEDGER_BALANCE = 0
                '                        End If
                '                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values (" & TAX_LEDGER_SLNO & ",'" & Me.txtStockindate.Text & "','" & Me.cboSupplier.Text & "'," & Val(Me.TxtVatAmount.Text) / 2 & ",0," & TAX_LEDGER_BALANCE - (Val(Me.TxtVatAmount.Text) / 2) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvno.Text & "-" & Me.txtinvdate.Text & "','Purchase'," & Slno & "," & rec3("GroupId") & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                '                    End If
                '                End If
                '                End If
                '                '--------New Transaction Entry Of IGST Ledger-----------------
                '                If Val(Me.txtstatecode.Text) <> 21 Then
                '                Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccName like 'IGST'")
                '                If Not rec3.EOF Then
                '                    If Val(Me.TxtVatAmount.Text) > 0 Then
                '                        Set rec4 = db.OpenRecordset("select * from  LedgerTran where AccId=" & rec3("AccId") & " and SlNo=(select max(SlNo) from LedgerTran where AccId=" & rec3("AccId") & ")")
                '                        If Not rec4.EOF Then
                '                            TAX_LEDGER_SLNO = rec4("Slno") + 1
                '                            TAX_LEDGER_BALANCE = rec4("Balance")
                '                        Else
                '                            TAX_LEDGER_SLNO = 1
                '                            TAX_LEDGER_BALANCE = 0
                '                        End If
                '                        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values (" & TAX_LEDGER_SLNO & ",'" & Me.txtStockindate.Text & "','" & Me.cboSupplier.Text & "'," & Val(Me.TxtVatAmount.Text) / 2 & ",0," & TAX_LEDGER_BALANCE - (Val(Me.TxtVatAmount.Text) / 2) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvno.Text & "-" & Me.txtinvdate.Text & "','Purchase'," & Slno & "," & rec3("GroupId") & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                '                    End If
                '                End If
                '                End If

            End If
            db.Execute ("delete * from Temp_Stockin")
            Data1.Refresh
            AccId = 0
            Slno = 0
            prAccId = 0
            'Me.txtinvdate.Text = ""
            Me.txtInvno.Text = ""
            Me.txtslno.Locked = True
            TAX_LEDGER_SLNO = 0
            TAX_LEDGER_BALANCE = 0
            TEMP_PURCHASE_LEDGER_SLNO = 0
            TEMP_PURCHASE_LEDGER_BALANCE = 0
            temp_ledger_balance = 0
            temp_ledger_slno = 0
            Me.txtTotalqty.Text = "0.00"
            Me.txttotalgross.Text = "0.00"

            Me.txtfreight.Text = "0.00"
            Me.txtdiscountamount.Text = "0.00"
            Me.TxtVatAmount.Text = "0.00"

            Me.TxtNetAmount.Text = "0.00"
            Me.TxtRoundup.Text = "0.00"

            Me.txtGrandtotal.Text = "0.00"
            PURCHASEDELETE = "N"
            Set rec3 = db.OpenRecordset("select max(Slno)  as max_no from PurchaseHead")
            If Not IsNull(rec3!max_no) Then
                Me.txtslno.Text = rec3!max_no + 1
            Else
                Me.txtslno.Text = 1
            End If
            Me.txtStockindate.SetFocus
        End If
    End If
End Sub

Private Sub Combo2_Change()

End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Me.txtTotalqty.Text = Format(Val(Me.txtTotalqty.Text) - Val(Me.DBGrid1.Columns(3)) - Val(Me.DBGrid1.Columns(5)), "############0.00")
    Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.DBGrid1.Columns(8)), "#########0.00")
    Me.txtmrpamount.Text = Format(Val(Me.txtmrpamount.Text) - (Val(Me.DBGrid1.Columns(3)) * Val(Me.DBGrid1.Columns(7))), "########0.00")
    Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) - Val(Me.DBGrid1.Columns(13)), "#######0.00")
    Me.txtdiscountamount.Text = Format(Val(Me.txtdiscountamount.Text) - Val(Me.DBGrid1.Columns(11)), "####0.00")
End Sub
Private Sub Form_Load()
    FORMNAME = "stockin"
    Set WinHttpReq = New WinHttpRequest
    Data1.databasename = dbname
    formid = 3
    FORMNAME = "Purchase"
    Me.Top = 0
    Me.Left = 0
    OnlineInvSlno = 0
    OnlineInv = False
    Me.txtStockindate.Text = Format(Date, "dd/mm/yyyy")
    Me.txtinvdate.Text = Format(Date, "dd/mm/yyyy")
    Set rec1 = db.OpenRecordset("select max(slno) as max_no from PurchaseHead")
    If Not IsNull(rec1!max_no) Then
        Me.txtslno.Text = rec1!max_no + 1
        If SoftwareVersion = "Demo" And Val(Me.txtslno.Text) > 3 Then
                MsgBox "Demo Expired", vbCritical
                Unload frmMain
                End
            End If
    Else
        Me.txtslno.Text = 1
    End If

    Set rec1 = db.OpenRecordset("select * from Partycr")
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboSupplier.AddItem (rec1("Party"))
            Me.cboSupplier.ItemData(Me.cboSupplier.NewIndex) = rec1("AccId")
            rec1.MoveNext
        Wend
    End If
    If Me.cboSupplier.ListCount > 0 Then
        Me.cboSupplier.ListIndex = 0
    End If
    Set rec = db.OpenRecordset("select * from stock")
    While Not rec.EOF
        Me.cboitemname.AddItem rec("itemname")
        Me.cboitemname.ItemData(Me.cboitemname.NewIndex) = rec("productcode")
        rec.MoveNext
    Wend
    If Me.cboitemname.ListCount > 0 Then
        Me.cboitemname.ListIndex = 0
    End If
    
    Set rec = db.OpenRecordset("select * from Product")
    If Not rec.EOF Then
        While Not rec.EOF
            Me.cbocategory.AddItem (rec("Productname"))
            Me.cbocategory.ItemData(Me.cbocategory.NewIndex) = rec("Pid")
            rec.MoveNext
        Wend
        If Me.cbocategory.ListCount > 0 Then
            Me.cbocategory.ListIndex = 0
        End If
    End If
    Set rec = db.OpenRecordset("select * from Brandmaster")
    If Not rec.EOF Then
        Me.cbobrand.Clear
        While Not rec.EOF
            Me.cbobrand.AddItem (rec("brand"))
            Me.cbobrand.ItemData(Me.cbobrand.NewIndex) = rec("BrandId")
            rec.MoveNext
        Wend
        If Me.cbobrand.ListCount > 0 Then
            Me.cbobrand.ListIndex = 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("delete * from Temp_Stockin where Usname='" & usname & "'")
    db.Execute ("delete * from Temp_Stockin")
    TAXAMOUNT = 0
    formid = 0
    temp_ledger_balance = 0
    temp_ledger_slno = 0
    temp_lr_balance = 0
    temp_lr_slno = 0
    TEMP_GROUPID = 0
    TEMP_PURCHASE_LEDGER_SLNO = 0
    TEMP_PURCHASE_LEDGER_BALANCE = 0
    PURCHASE_ACCID = 0
    PURCHASE_GROUPID = 0
    TAX_LEDGER_ACCID = 0
    TAX_LEDGER_GROUPID = 0
    TAX_LEDGER_SLNO = 0
    TAX_LEDGER_BALANCE = 0
End Sub

Private Sub Text1_Change()

End Sub

Private Sub lstinv_Click()
    invid = Me.lstinv.ListIndex + 1
    Me.txtinvdate.Text = JSONRec(invid).Item("InvDate")
    Me.txtInvno.Text = JSONRec(invid).Item("InvNo")
    Me.txtTotalqty.Text = JSONRec(invid).Item("TotalQty")
    Me.txttotalgross.Text = JSONRec(invid).Item("TotalGross")
    Me.TxtVatAmount.Text = JSONRec(invid).Item("VatAmount")
    Me.TxtNetAmount.Text = JSONRec(invid).Item("NetAmount")
    Me.txtGrandtotal.Text = JSONRec(invid).Item("GrandTotal")
    Me.txtmrpamount.Text = JSONRec(invid).Item("MrpAmount")
    Me.txtdiscountamount.Text = Val(JSONRec(invid).Item("TradeDiscount")) + Val(JSONRec(invid).Item("SpecialDiscount"))
    Me.txtfreight.Text = JSONRec(invid).Item("Frieght")
    Me.TxtNetAmount.Text = JSONRec(invid).Item("Net")
    Me.TxtRoundup.Text = JSONRec(invid).Item("RndUp")
    Me.txtGrandtotal.Text = JSONRec(invid).Item("GrandTotal")
        
End Sub

Private Sub lstinv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    invid = Me.lstinv.ListIndex + 1
    Me.txtinvdate.Text = JSONRec(invid).Item("InvDate")
    Me.txtInvno.Text = JSONRec(invid).Item("InvNo")
    OnlineInv = True
    OnlineInvSlno = JSONRec(invid).Item("invslno")
    Me.txtInvno.SetFocus
    Me.lstinv.Visible = False
End If
If KeyCode = vbKeyEscape Then
    OnlineInv = False
    OnlineInvSlno = 0
    
    Me.txtinvdate.Text = "__/__/____"
    Me.txtInvno.Text = ""
    Me.txtTotalqty.Text = 0
    Me.txttotalgross.Text = "0.00"
    Me.TxtVatAmount.Text = "0.00"
    Me.TxtNetAmount.Text = "0.00"
    Me.txtGrandtotal.Text = "0.00"
End If
End Sub

Private Sub txtAmount_GotFocus()
    Me.txtamount.SelStart = 0
    Me.txtamount.SelLength = Len(Me.txtamount.Text)
End Sub
Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

    End If
End Sub

Private Sub txtcd_GotFocus()
Me.txtcd.SelStart = 0
Me.txtcd.SelLength = Len(Me.txtcd.Text)
End Sub

Private Sub txtcd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.TxtVat.SetFocus
End If
End Sub

Private Sub TXTCST_GotFocus()
Me.txtcst.SelStart = 0
Me.txtcst.SelLength = Len(Me.txtcst.Text)
End Sub

Private Sub TXTCST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtcstamount.Text = Round((Val(Me.txttotalgross.Text) - Val(Me.txtdiscountamount.Text)) * (Val(Me.txtcst.Text) / 100), 2)
    Me.txtet.SetFocus
End If
End Sub

Private Sub txtdiscount_GotFocus()
    Me.txtdiscount.SelStart = 0
    Me.txtdiscount.SelLength = Len(Me.txtdiscount.Text)
End Sub

Private Sub txtdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtspecialdiscount.SetFocus
End If
End Sub

Private Sub txtet_GotFocus()
Me.txtet.SelStart = 0
Me.txtet.SelLength = Len(Me.txtet.Text)
End Sub

Private Sub txtet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtetamount.Text = Round(Val(Me.txttotalgross.Text) * (Val(Me.txtet.Text) / 100), 2)
    Me.txtetamount.SetFocus
End If
End Sub

Private Sub txtetamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtfreight.SetFocus
End If
End Sub

Private Sub txtexpdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtpack.SetFocus
End If
End Sub

Private Sub txtfree_Change()
If Not ValidateNumeric(Me.txtfree.Text) Then
    Me.txtQty.Text = 0
    txtfree_GotFocus
End If
End Sub

Private Sub txtfree_GotFocus()
    Me.txtfree.SelStart = 0
    Me.txtfree.SelLength = Len(Me.txtfree.Text)
End Sub

Private Sub txtfree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtPrate.SetFocus
    End If
End Sub

Private Sub txtfreight_GotFocus()
    Me.txtfreight.SelStart = 0
    Me.txtfreight.SelLength = Len(Me.txtfreight.Text)
End Sub

Private Sub txtfreight_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtNetAmount.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.txtdiscountamount.Text) + Val(Me.TxtVatAmount.Text) + Val(Me.txtcstamount.Text) + Val(Me.txtetamount.Text) + Val(Me.txtfreight.Text), "#########0.00")
        Me.TxtNetAmount.SetFocus
    End If
End Sub

Private Sub txtGrandtotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub

Private Sub txthsn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        Me.txttaxp.SetFocus
    End If
End Sub

Private Sub txtiname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txthsn.SetFocus
End If
End Sub

Private Sub txtiname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtInvDate_GotFocus()
Me.txtinvdate.SelStart = 0
Me.txtinvdate.SelLength = Len(Me.txtinvdate.Text)
End Sub

Private Sub txtInvdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtLrNo.SetFocus
    End If
End Sub

Private Sub txtInvNo_GotFocus()
Me.lblmessage.Caption = "Press F2 to Pull Online Invoices"
End Sub

Private Sub txtInvno_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim fromPartyGID, ToPartyGID, FromPartyExist As Boolean, ToPartyExist As Boolean
    FromPartyExist = False
    ToPartyExist = False

    If KeyCode = 13 Then
        If OnlineInv = True Then
            GetOnlineInvDetails
        Else
            Me.txtinvdate.SetFocus
            Me.lstinv.Visible = False
        End If

    End If
    If KeyCode = vbKeyDown Then
        Me.lstinv.SetFocus
    End If
    If KeyCode = vbKeyF2 Then
        Dim strRes As String, reccount, countRecords
        Me.txtInvno.SelStart = 0
        Me.lstinv.Visible = True
        Me.lstinv.Clear
        Set rec1 = db.OpenRecordset("select * from companymaster")
        If rec1("gid") <> 0 Then
            FromPartyExist = True
            fromPartyGID = rec1("gid")
        Else
            WinHttpReq.Open "GET", _
                            "http://techspark.xp3.biz/enlite/getgid.php?Mobile=" & rec1("phone") & "&gstno=" & rec1("taxno"), False
            WinHttpReq.Send
            If WinHttpReq.ResponseText Like "*Not Found*" Then
                MsgBox "You are not a Registered User of EnLite", vbCritical
            Else
                strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
                Set p = JSON.parse(strRes)
                db.Execute "update companymaster set gid=" & p.Item("CompanyID")
                fromPartyGID = p.Item("CompanyID")
                FromPartyExist = True
            End If
        End If
        Set rec1 = db.OpenRecordset("select * from partycr where accid=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
        If rec1("gid") <> 0 Then
            ToPartyExist = True
            ToPartyGID = rec1("gid")
        Else

        End If
        If FromPartyExist And ToPartyExist Then
            
            WinHttpReq.Open "GET", _
                            "http://techspark.xp3.biz/enlite/getinv.php?frompartyid=" & ToPartyGID & "&topartyid=" & fromPartyGID, False
            WinHttpReq.Send
            If WinHttpReq.ResponseText Like "*Not Found*" Then
                MsgBox "Not Found", vbCritical
            Else
                strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
                countRecords = UBound(Split(strRes, "InvNo"))
                Set JSONRec = JSON.parse(strRes)
                If countRecords > 1 Then
                    reccount = 1
                    While reccount <= countRecords
                        Me.lstinv.AddItem JSONRec(reccount).Item("InvNo") & "-" & JSONRec(reccount).Item("InvDate")
                        reccount = reccount + 1
                    Wend
                Else
                    Me.lstinv.AddItem JSONRec(1).Item("InvNo") & "-" & JSONRec(1).Item("InvDate")
                End If
            End If
        Else
            MsgBox "Unregistered Users", vbCritical
        End If
    End If
End Sub

Private Sub txtlose_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub txtInvno_LostFocus()
'Me.lstinv.Visible = False
End Sub

Private Sub txtLrno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Me.cboSupplier.SetFocus
        Me.cboitemname.SetFocus
    End If
End Sub

Private Sub txtmfgdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'Me.txtexpdate.SetFocus
    Me.txtpack.SetFocus
End If
End Sub

Private Sub txtmrp_GotFocus()
    Me.txtmrp.SelStart = 0
    Me.txtmrp.SelLength = Len(Me.txtmrp.Text)
End Sub

Private Sub txtmrp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select Mrp from ItemMaster where ProductCode=" & Me.txtproductcode.Text)
        If Not rec1.EOF Then
            If rec1("Mrp") <> Val(Me.txtmrp.Text) Then
                ans = MsgBox("Update The MRP", vbYesNo)
                If ans = 6 Then
                    db.Execute ("Update ItemMaster set Mrp=" & Me.txtmrp.Text & " where ProductCode=" & Val(Me.txtproductcode.Text))
                Else
                    Me.txtdiscount.SetFocus
                End If
            End If
        End If
        Me.txtdiscount.SetFocus
    End If
End Sub

Private Sub txtnet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        discount_amount = Round(Val(Me.txtamount.Text) * (Val(Me.txtdiscount.Text) / 100), 2)
        Special_Discount = Round((Val(Me.txtamount.Text) - discount_amount) * (Val(Me.txtspecialdiscount.Text) / 100), 2)
        db.Execute ("insert into Temp_Stockin (ItemName,Units,Pack,Qty,Mrp,PrRate,Amount,ProductCode,Vat,VatAmount,Net,Usname,Free_Qty,Discount,SpDiscount,Discount_amount,cd,batchno,mfgdate,expdate) values('" & Me.cboitemname.Text & "','" & Me.cbounit.Text & "'," & Me.txtpack.Text & "," & (Val(Me.txtQty.Text)) & "," & Me.txtmrp.Text & "," & Val(Me.txtPrate.Text) & "," & Me.txtamount.Text & "," & Me.txtproductcode.Text & "," & Val(Me.TxtVat.Text) & "," & Me.txttaxamount.Text & "," & Me.txtnet.Text & ",'" & usname & "'," & Me.txtfree.Text & "," & Me.txtdiscount.Text & "," & Me.txtspecialdiscount.Text & "," & discount_amount + Special_Discount + Val(Me.txtcd.Text) & "," & Me.txtcd.Text & ",'" & Me.cbobatch.Text & "','" & Me.txtmfgdate.Text & "','" & Me.txtexpdate.Text & "')")
        Data1.Refresh
        Me.txtTotalqty.Text = Val(Me.txtTotalqty.Text) + Val(Me.txtQty.Text) + Val(Me.txtfree.Text)
        Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) + Val(Me.txtamount.Text), "################0.00")
        Me.txtdiscountamount.Text = Format(Val(Me.txtdiscountamount.Text) + discount_amount + Special_Discount + Val(Me.txtcd.Text), "######0.00")
        Me.txtmrpamount.Text = Format(Val(Me.txtmrpamount.Text) + (Val(Me.txtmrp.Text) * Val(Me.txtQty.Text)), "########0.00")
        Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) + Val(Me.txttaxamount.Text), "########0.00")
        Me.txtmrp.Text = "0.00"
        'Me.TxtVat.Text = 0
        Me.txttaxamount.Text = "0.00"
        Me.txtnet.Text = "0.0"
        'Me.txtPrate.Text = "0.00"
        Me.txtamount.Text = "0.00"
        Me.txtdiscount.Text = 0
        discount_amount = 0
        Me.txtcd.Text = "0"
        Me.txtspecialdiscount.Text = 0
        discount_amount = 0
        Special_Discount = 0
        VatAmount = 0
        Me.cboitemname.SetFocus
    End If
End Sub

Private Sub txtNetamount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtRoundup.Text = Format(Round(Val(Me.TxtNetAmount.Text)) - Val(Me.TxtNetAmount.Text), "########0.00")
        Me.TxtRoundup.SetFocus
    End If
End Sub

Private Sub txtnewbrand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtnewbrand.Visible = False
    Me.cbobrand.SetFocus
End If
If KeyCode = 13 Then
    Set rec1 = db.OpenRecordset("select * from Brandmaster where brand='" & Me.txtnewbrand.Text & "'")
            If Not rec1.EOF Then
                MsgBox "Allready exists", vbCritical
            Else
                ans = MsgBox("Save This?", vbYesNo)
                If ans = 6 Then
                    db.Execute ("insert into Brandmaster (brand,Purchase,Sale) values('" & Me.txtnewbrand.Text & "',0, 0)")
                    Me.cbobrand.AddItem Me.txtnewbrand.Text
                    Me.cbobrand.ListIndex = Me.cbobrand.NewIndex
                    Me.txtnewbrand.Text = ""
                    Me.cbobrand.Visible = True
                    Me.txtnewbrand.Visible = False
                    Me.cbobrand.SetFocus
                End If
                
            End If
End If
End Sub

Private Sub txtnewbrand_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnewmtype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.txtnewmtype.Visible = False
    Me.cbomtype.SetFocus
End If
If KeyCode = 13 Then
     Set rec1 = db.OpenRecordset("select * from ItemType where producttype='" & Me.cbocategory.Text & "' and Item_Type='" & Me.txtnewmtype.Text & "'")
        If Not rec1.EOF Then
            MsgBox "Allready Exists?", vbCritical
        Else
            ans = MsgBox("Save This?", vbYesNo)
            If ans = 6 Then
                db.Execute ("insert into ItemType (Item_Type,ProductType) values('" & Me.txtnewmtype.Text & "','" & Me.cbocategory.Text & "')")
                Me.cbomtype.AddItem Me.txtnewmtype.Text
                Me.txtnewmtype.Text = ""
                Me.txtnewmtype.Visible = False
                Me.cbomtype.SetFocus
            End If
        End If
End If
End Sub

Private Sub txtnewmtype_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtpack_Change()
If Not ValidateNumeric(Me.txtpack.Text) Then
    Me.txtpack.Text = 0
    txtpack_GotFocus
End If
End Sub

Private Sub txtpack_GotFocus()
Me.txtpack.SelStart = 0
Me.txtpack.SelLength = Len(Me.txtpack.Text)
End Sub

Private Sub txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Val(Me.txtpack.Text) > 0 Then
        Set rec1 = db.OpenRecordset("select Lose from ItemMaster where ProductCode=" & Val(Me.txtproductcode.Text))
        If Not rec1.EOF Then
            Me.txtQty.Text = Val(Me.txtpack.Text) * Val(rec1("Lose"))
        Else
            Me.txtQty.Text = 0
        End If
    End If
    Me.txtQty.SetFocus
End If

End Sub

Private Sub txtPrate_GotFocus()
    Me.txtPrate.SelStart = 0
    Me.txtPrate.SelLength = Len(Me.txtPrate.Text)
End Sub
Private Sub txtPrate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select purchaserate from ItemMaster where ProductCode=" & Me.txtproductcode.Text)
        If Not rec1.EOF Then
            If rec1("purchaserate") <> Val(Me.txtPrate.Text) Then
                ans = MsgBox("Update The Rate", vbYesNo)
                If ans = 6 Then
                    db.Execute ("Update ItemMaster set Purchaserate=" & Me.txtPrate.Text & " where ProductCode=" & Val(Me.txtproductcode.Text))
                Else
                    Me.txtdiscount.SetFocus
                End If
            End If
        End If
        Me.txtamount.Text = Format(Round((Val(Me.txtQty.Text)) * Val(Me.txtPrate.Text), 2), "###########0.00")
        Me.txtmrp.SetFocus
    End If
End Sub

Private Sub txtQty_Change()
If Not ValidateNumeric(Me.txtQty.Text) Then
    Me.txtQty.Text = 0
    txtqty_GotFocus
End If
End Sub

Private Sub txtqty_GotFocus()
    Me.txtQty.SelStart = 0
    Me.txtQty.SelLength = Len(Me.txtQty.Text)
End Sub
Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtfree.SetFocus
    End If
End Sub
Private Sub TxtRoundup_GotFocus()
    Me.TxtRoundup.SelStart = 0
    Me.TxtRoundup.SelLength = Len(Me.TxtRoundup.Text)
End Sub
Private Sub TxtRoundup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtGrandtotal.Text = Format(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text) + Val(Me.TxtRoundup.Text), "#####0.00")
        Me.txtGrandtotal.SetFocus
    End If
End Sub
Private Sub txtslno_GotFocus()
    Me.txtslno.SelStart = 0
    Me.txtslno.SelLength = Len(Me.txtslno.Text)
End Sub
Private Sub txtSlno_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then
        db.Execute ("delete * from Temp_Stockin")
        Set rec2 = db.OpenRecordset("select * from PurchaseHead where Slno=" & Me.txtslno.Text)
        If Not rec2.EOF Then
            Me.txtStockindate.Text = rec2("Purchasedate")
            Me.txtInvno.Text = rec2("InvNo")
            Me.txtinvdate.Text = rec2("InvDate")
            Me.txtwaybill.Text = rec2("Waybill")
            Me.txttotalgross.Text = Format(rec2("TotalGross"), "######0.00")
            Me.txtmrpamount.Text = Format(rec2("TotalMrp"), "########0.00")
        
            Me.txtTotalqty.Text = Format(rec2("TotalQty"), "#####0.00")
            
            Me.TxtVatAmount.Text = Format(rec2("VatAmount"), "#####0.00")
            Me.txtcst.Text = rec2("CST")
            Me.txtcstamount.Text = rec2("CSTAmount")
            Me.txtet.Text = rec2("ETax")
            Me.txtetamount.Text = rec2("ETaxAmount")
            
            Me.TxtNetAmount.Text = Format(rec2("NetAmount"), "######0.00")
            Me.txtGrandtotal.Text = Format(rec2("GrandTotal"), "######0.00")
            
            Me.TxtRoundup.Text = rec2("RValue")
            Me.txtfreight.Text = Format(rec2("Freight"), "#######0.00")
            Me.txtdiscountamount.Text = rec2("lessDiscount")

            Set rec3 = db.OpenRecordset("select * from ledgermaster where AccId=" & rec2("AccId"))
            If Not rec3.EOF Then
                Me.cboSupplier.Text = rec3("AccName")
            End If
            Set rec1 = db.OpenRecordset("select * from PurchaseDetails where SlNo=" & Me.txtslno.Text)
            If Not rec1.EOF Then
                While Not rec1.EOF
                    db.Execute ("insert into Temp_Stockin (ItemName,Units,Pack,Qty,Mrp,PrRate,Amount,ProductCode,Vat,VatAmount,Net,Free_Qty,Discount,SpDiscount,Discount_amount,cd,mfgdate,expdate,batchno) values('" & rec1("ItemName") & "','" & rec1("Units") & "'," & rec1("Pack") & "," & rec1("Qty") & "," & rec1("Mrp") & "," & rec1("PrRate") & "," & rec1("Amount") & "," & rec1("ProductCode") & "," & rec1("Vat") & "," & rec1("VatAmount") & "," & rec1("Net") & "," & rec1("Free_Qty") & "," & rec1("Discount") & "," & rec1("SpDiscount") & "," & rec1("Discount_amount") & "," & rec1("cd") & ",'" & rec1("mfgdate") & "','" & rec1("expdate") & "','" & rec1("batchno") & "')")
                    Data1.Refresh
                    rec1.MoveNext
                Wend
                Me.txtslno.Locked = True
            End If
            Me.txtInvno.SetFocus
        Else
            MsgBox "Not Found", vbCritical
        End If
        '//--------------deleteing purchase------------
        If PURCHASEDELETE = "Y" Then
            ans = MsgBox("Delete this purchase?", vbYesNo)
            If ans = 6 Then
                '//Deleting Purchasereturn details-----------------
                Set rec2 = db.OpenRecordset("select * from PurchaseDetails where Slno=" & Me.txtslno.Text)
                If Not rec2.EOF Then
                    While Not rec2.EOF
                        db.Execute ("update stock set Qty=Qty - " & rec2("Qty") + rec2("Free_Qty") & " where ProductCode=" & rec2("ProductCode"))
                        db.Execute ("update stockdetails set Qty=Qty - " & rec2("Qty") + rec2("Free_Qty") & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                        rec2.MoveNext
                    Wend
                    db.Execute ("delete * from PurchaseDetails where SlNo=" & Me.txtslno.Text)
                End If
                '// deleting party ledger
                db.Execute ("delete * from LedgerTran where VoucherType='Purchase' and VoucherSlno=" & Me.txtslno.Text)
                '//deleting purchase head-----------------
                db.Execute ("delete * from Purchasehead where Slno=" & Me.txtslno.Text)
            End If
            db.Execute ("delete * from Temp_Stockin")
            Data1.Refresh
            AccId = 0
            prAccId = 0
            Me.txtinvdate.Text = ""
            Me.txtInvno.Text = ""
            Me.txtslno.Locked = True
            Me.txtTotalqty.Text = "0.00"
            Me.txttotalgross.Text = "0.00"
            
            Me.TxtVat.Text = "0.00"
            Me.TxtVatAmount.Text = "0.00"

            Me.TxtNetAmount.Text = "0.00"
            Me.TxtRoundup.Text = "0.00"

            Me.txtGrandtotal.Text = "0.00"

            Set rec3 = db.OpenRecordset("select max(Slno)  as max_no from PurchaseHead")
            If Not IsNull(rec3!max_no) Then
                Me.txtslno.Text = rec3!max_no + 1
            Else
                Me.txtslno.Text = 1
            End If
            Me.txtStockindate.SetFocus

        End If


        PURCHASEDELETE = "N"
    End If
End Sub

Private Sub txtspecialdiscount_GotFocus()
Me.txtspecialdiscount.SelStart = 0
Me.txtspecialdiscount.SelLength = Len(Me.txtspecialdiscount.Text)
End Sub

Private Sub txtspecialdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'Me.TxtVat.SetFocus
    Me.txtcd.SetFocus
End If
End Sub

Private Sub txtStockindate_GotFocus()
    Me.txtStockindate.SelStart = 0
    Me.txtStockindate.SelLength = Len(Me.txtStockindate.Text)
    Me.lblmessage.Caption = "Change Date and Press Enter to Go to Next Field"
End Sub
Private Sub txtStockindate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtwaybill.SetFocus
    End If
End Sub
Private Sub TxtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtVat.SetFocus
    End If
End Sub

Private Sub txttaxp_GotFocus()
Me.txttaxp.SelStart = 0
Me.txttaxp.SelLength = Len(Me.txttaxp.Text)
End Sub

Private Sub txttaxp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select * from ItemMaster where Item='" & Me.txtiname.Text & "' and brand='" & cbobrand.Text & "'")
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
                db.Execute ("insert into ItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,Purchaserate,Salerate,Openingstock,ProductCode,tax_type,BrandId,CategoryId,hsn) values('" & Me.cbocategory.Text & "','" & Me.cbomtype.Text & "','" & Me.cbobrand.Text & "','" & Me.txtiname.Text & "','" & Me.txtiname.Text & "',' ','PCS',0,0," & Me.txttaxp.Text & ",0,0,0," & Productcode & ",'SALES'," & Me.cbobrand.ItemData(Me.cbobrand.ListIndex) & "," & Me.cbocategory.ItemData(Me.cbocategory.ListIndex) & ",'" & Me.txthsn.Text & "')")
                db.Execute ("insert into Stock (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,hsn) values('" & Me.cbocategory.Text & "','" & Me.cbomtype.Text & "','" & Me.txtiname.Text & "','" & Me.cbobrand.Text & "','" & Me.txtiname.Text & "','',0,0,0," & Productcode & "," & Me.txttaxp.Text & ",1,'PCS',0,'" & Me.txthsn.Text & "')")
                Me.cboitemname.AddItem Me.txtiname.Text
                Me.cboitemname.ItemData(Me.cboitemname.NewIndex) = Productcode
                Me.cboitemname.ListIndex = Me.cboitemname.NewIndex
                Me.txttaxp.Text = "0"
                Me.txtiname.Text = ""
                Me.cboitemname.SetFocus
                Me.frnewitem.Visible = False

            End If
        End If
    End If
End Sub

Private Sub txtvat_GotFocus()
    Me.TxtVat.SelStart = 0
    Me.TxtVat.SelLength = Len(Me.TxtVat.Text)
End Sub

Private Sub TxtVat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select tax from ItemMaster where ProductCode=" & Me.txtproductcode.Text)
        If Not rec1.EOF Then
            If rec1("tax") <> Val(Me.TxtVat.Text) Then
                ans = MsgBox("Update The GST", vbYesNo)
                If ans = 6 Then
                    db.Execute ("Update ItemMaster set tax=" & Me.TxtVat.Text & " where ProductCode=" & Val(Me.txtproductcode.Text))
                Else
                   Me.txtdiscount.SetFocus
                End If
            End If
        End If
        discount_amount = Round(Val(Me.txtamount.Text) * (Val(Me.txtdiscount.Text) / 100), 2)
        Special_Discount = (Val(Me.txtamount.Text) - discount_amount) * (Val(Me.txtspecialdiscount.Text) / 100)
        temp_dis1 = Val(Me.txtPrate.Text) * (Val(Me.txtdiscount.Text) / 100)
        temp_dis2 = (Val(Me.txtPrate.Text) - (temp_dis1)) * (Val(Me.txtspecialdiscount.Text) / 100)
        temp_rate = Val(Me.txtPrate.Text) - (temp_dis1 + temp_dis2)
        
        If Me.txttaxtype.Text = "MRP" Then
            Me.txttaxamount.Text = Round((Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text))) * (Val(Me.TxtVat.Text) / 100), 2)
        End If
        If Me.txttaxtype.Text = "SALES" Then
            Me.txttaxamount.Text = Round((Val(Me.txtamount.Text) - discount_amount - Special_Discount - Val(Me.txtcd.Text)) * (Val(Me.TxtVat.Text) / 100), 2)
        End If
        If Me.txttaxtype.Text = "INCLUSIVE MRP" Then
            Me.txttaxamount.Text = Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) - Format((Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) / ((Me.TxtVat.Text / 100) + 1)), "########0.00")
        End If
        Me.txtnet.Text = Format(Val(Me.txtamount.Text) - discount_amount - Special_Discount - Val(Me.txtcd.Text) + Val(Me.txttaxamount.Text), "########0.00")
        Me.txtnet.SetFocus
End If
End Sub

Private Sub txtwaybill_GotFocus()
Me.lblmessage.Caption = "Press Enter to Go to Next Field"
End Sub

Private Sub txtwaybill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'Me.txtInvno.SetFocus
    Me.cboSupplier.SetFocus
End If
End Sub

Public Sub GetOnlineInvDetails()
    Dim strRes As String
    If OnlineInv = True And OnlineInvSlno <> 0 Then
        WinHttpReq.Open "GET", _
                        "http://techspark.xp3.biz/enlite/getinvdetails.php?invslno=" & OnlineInvSlno, False
        WinHttpReq.Send
        If WinHttpReq.ResponseText Like "*Not Found*" Then
            MsgBox "Not Found", vbCritical
        Else
            strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
            countRecords = UBound(Split(strRes, "InvNo"))
            Set JSONRec = JSON.parse(strRes)
            reccount = 1
            While reccount <= countRecords
                PType = JSONRec(reccount).Item("ProductType")
                IType = JSONRec(reccount).Item("ItemType")
                BName = JSONRec(reccount).Item("Brandname")
                IName = JSONRec(reccount).Item("Itemname")
                TSize = JSONRec(reccount).Item("Size")
                Units = JSONRec(reccount).Item("Units")
                MRP = JSONRec(reccount).Item("MRP")
                SaleRate = JSONRec(reccount).Item("SaleRate")
                Qty = JSONRec(reccount).Item("Qty")
                Gross = JSONRec(reccount).Item("Gross")
                SDiscount = JSONRec(reccount).Item("SpecialDiscount")
                TDiscount = JSONRec(reccount).Item("Tradediscount")
                DisAmount = JSONRec(reccount).Item("DiscountAmount")
                Vat = JSONRec(reccount).Item("Vat")
                VatAmount = JSONRec(reccount).Item("VatAmount")
                Net = JSONRec(reccount).Item("Net")
                Pcode = JSONRec(reccount).Item("ProductCode")
                FQty = JSONRec(reccount).Item("Free_Qty")
                TaxType = JSONRec(reccount).Item("Tax_type")
                Pack = JSONRec(reccount).Item("Pack")
                HSN = JSONRec(reccount).Item("HSN")
                MfgDate = JSONRec(reccount).Item("MfgDate")
                ExpDate = JSONRec(reccount).Item("ExpDate")
                BatchNo = JSONRec(reccount).Item("BatchNo")
                aslno = JSONRec(reccount).Item("adapterslno")
                bslno = JSONRec(reccount).Item("batteryslno")
                'Me.lstinv.AddItem JSONRec(reccount).Item("InvNo") & "-" & JSONRec(reccount).Item("InvDate")
                Set rec1 = db.OpenRecordset("select * from itemmaster where brand='" & JSONRec(reccount).Item("Brandname") & "' and item='" & JSONRec(reccount).Item("Itemname") & "' and producttype='" & JSONRec(reccount).Item("ProductType") & "' and itemtype='" & JSONRec(reccount).Item("ItemType") & "'")
                If Not rec1.EOF Then
                    db.Execute ("insert into Temp_Stockin (ItemName,Units,Pack,Qty,Mrp,PrRate,Amount,ProductCode,Vat,VatAmount,Net,Usname,Free_Qty,Discount,SpDiscount,Discount_amount,cd,batchno,mfgdate,expdate) values('" & IName & "','" & Units & "'," & Pack & "," & Qty & "," & MRP & "," & SaleRate & "," & Gross & "," & Pcode & "," & Vat & "," & VatAmount & "," & Net & ",'" & usname & "'," & FQty & "," & TDiscount & "," & SDiscount & "," & DisAmount & "," & Me.txtcd.Text & ",'" & BatchNo & "','" & MfgDate & "','" & ExpDate & "')")
                    Me.Data1.Refresh
                Else
                    'check product category exists or not
                    Set rec1 = db.OpenRecordset("select * from product where productname='" & PType & "'")
                    If rec1.EOF Then
                        db.Execute ("insert into product (productname) values('" & PType & "')")
                    End If
                    Set rec1 = db.OpenRecordset("select * from itemtype where producttype='" & PType & "' and item_type='" & IType & "'")
                    If rec1.EOF Then
                        db.Execute ("insert into itemtype (producttype,item_type) values('" & PType & "','" & IType & "')")
                    End If
                    Set rec1 = db.OpenRecordset("select * from brandmaster where brand='" & BName & "'")
                    If rec1.EOF Then
                        db.Execute ("insert into brandmaster (brand,purchase,sale,brandid) values('" & BName & "',0,0,0)")
                    End If
                    'check item category exists or not
                    Set rec1 = db.OpenRecordset("select max(productcode) as productid from ItemMaster")
                    If Not IsNull(rec1!productid) Then
                        Productcode = rec1!productid + 1
                    Else
                        Productcode = 1000
                    End If
                    db.Execute ("insert into ItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,Purchaserate,Salerate,Openingstock,ProductCode,tax_type,BrandId,CategoryId,hsn) values('" & PType & "','" & IType & "','" & BName & "','" & IName & "','" & IName & "','" & TSize & "','" & Unit & "',1," & MRP & "," & Vat & "," & SaleRate & "," & MRP & ",0," & Productcode & ",'Sales',0,0,'" & HSN & "')")
                    db.Execute ("insert into Stock (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,hsn) values('" & PType & "','" & IType & "','" & IName & "','" & BName & "','" & IName & "','" & TSize & "'," & MRP & "," & SaleRate & ",0," & Productcode & "," & Vat & ",1,'" & Unit & "'," & MRP & ",'" & HSN & "')")
                    db.Execute ("insert into Temp_Stockin (ItemName,Units,Pack,Qty,Mrp,PrRate,Amount,ProductCode,Vat,VatAmount,Net,Usname,Free_Qty,Discount,SpDiscount,Discount_amount,cd,batchno,mfgdate,expdate) values('" & IName & "','" & Units & "'," & Pack & "," & Qty & "," & MRP & "," & SaleRate & "," & Gross & "," & Pcode & "," & Vat & "," & VatAmount & "," & Net & ",'" & usname & "'," & FQty & "," & TDiscount & "," & SDiscount & "," & DisAmount & ",0,'" & BatchNo & "','" & MfgDate & "','" & ExpDate & "')")
                    Me.Data1.Refresh
                End If

                reccount = reccount + 1
            Wend
            Me.cboitemname.SetFocus
        End If
    Else
        MsgBox "Unregistered Users", vbCritical
    End If
End Sub
