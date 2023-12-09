VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalesReturn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Credit Note [Sales Return]"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   55
      Top             =   1680
      Width           =   14535
      Begin VB.TextBox cbounit 
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
         Left            =   6120
         TabIndex        =   81
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox cboitemname 
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
         Left            =   960
         TabIndex        =   80
         Top             =   360
         Width           =   4695
      End
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
         TabIndex        =   74
         Text            =   "cbobatch"
         Top             =   960
         Width           =   1500
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
         Left            =   13440
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   960
         Width           =   975
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
         Left            =   5400
         TabIndex        =   6
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtsalerate 
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
         Left            =   7680
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtproductcode 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtgross 
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txttradediscount 
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
         Left            =   10440
         TabIndex        =   11
         Text            =   "0"
         Top             =   960
         Width           =   615
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
         Left            =   11160
         TabIndex        =   12
         Text            =   "0"
         Top             =   960
         Width           =   615
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
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   960
         Width           =   615
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
         Left            =   6960
         TabIndex        =   7
         Text            =   "0"
         Top             =   960
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
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   240
         Width           =   1215
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
         Left            =   4680
         TabIndex        =   5
         Text            =   "0"
         Top             =   960
         Width           =   615
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
         Left            =   8640
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   960
         Width           =   735
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
         Left            =   12600
         TabIndex        =   56
         Text            =   "0.00"
         Top             =   960
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtexpdate 
         Height          =   315
         Left            =   3000
         TabIndex        =   75
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   1560
         TabIndex        =   76
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   3000
         TabIndex        =   79
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label38 
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
         Left            =   1560
         TabIndex        =   78
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch No"
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
         TabIndex        =   77
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Net"
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
         Left            =   13440
         TabIndex        =   72
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   6120
         TabIndex        =   71
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
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
         Left            =   5400
         TabIndex        =   70
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
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
         Left            =   960
         TabIndex        =   69
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label33 
         Caption         =   " Rate"
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
         Left            =   7680
         TabIndex        =   68
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Pr.Code"
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
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label19 
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
         Left            =   9480
         TabIndex        =   66
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Dis %"
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
         Left            =   10440
         TabIndex        =   65
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "S.Dis%"
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
         Left            =   11160
         TabIndex        =   64
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "Vat%"
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
         Left            =   11880
         TabIndex        =   63
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label16 
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
         Left            =   6960
         TabIndex        =   62
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label18 
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
         Left            =   4680
         TabIndex        =   61
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label21 
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
         Left            =   8640
         TabIndex        =   60
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Vat Amt"
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
         Left            =   12600
         TabIndex        =   59
         Top             =   720
         Width           =   735
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TempCreditNote"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   14535
      Begin VB.TextBox txtLrno 
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
         Left            =   9360
         TabIndex        =   4
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtTin 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cboParty 
         BackColor       =   &H00FFFFFF&
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
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox txtChalanNo 
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
         Left            =   9360
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtInvNo 
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "0"
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtChalanDate 
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
         Left            =   10560
         TabIndex        =   2
         Text            =   "##/##/####"
         Top             =   480
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   285
         Left            =   9360
         TabIndex        =   0
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
      Begin VB.Label LblAdr1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   54
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label30 
         Caption         =   "Lr No.Date"
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
         Left            =   8040
         TabIndex        =   53
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label28 
         Caption         =   "TIN /SRIN"
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
         Left            =   8040
         TabIndex        =   52
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Party"
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
         TabIndex        =   51
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Ch No Date"
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
         Left            =   8040
         TabIndex        =   50
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   8040
         TabIndex        =   49
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sl No"
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
         TabIndex        =   48
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   6780
      Width           =   14535
      Begin VB.TextBox txttotalspecialdiscount 
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H000000FF&
         Caption         =   "DELETE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9120
         TabIndex        =   31
         Top             =   1440
         Width           =   850
      End
      Begin VB.TextBox TxtNetAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox TxtVatAmount 
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   1560
         Width           =   1575
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
         Left            =   12720
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H0000FFFF&
         Caption         =   "EDIT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8160
         TabIndex        =   27
         Top             =   1440
         Width           =   850
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFC0FF&
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10080
         TabIndex        =   26
         Top             =   1440
         Width           =   850
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0000FF00&
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7200
         TabIndex        =   25
         Top             =   1440
         Width           =   850
      End
      Begin VB.TextBox txttotalgross 
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtTotalqty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtGrandtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txttotaltradediscount 
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtbalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtcrlimit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   720
         Width           =   1095
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
         Left            =   12720
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtmrpamount 
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
         Left            =   5400
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Special Discount"
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
         Left            =   3360
         TabIndex        =   44
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label20 
         Caption         =   "Net Amount"
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
         Left            =   11040
         TabIndex        =   43
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "GST Amount"
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
         Left            =   3360
         TabIndex        =   42
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Round Up"
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
         Left            =   11040
         TabIndex        =   41
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "Total Gross"
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
         Left            =   3360
         TabIndex        =   40
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label31 
         Caption         =   "Total Qty"
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
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "Grand Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   38
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label29 
         Caption         =   "Less Trade Discount"
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
         Left            =   3360
         TabIndex        =   37
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Cr.Limit"
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
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Balance"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Add Freight"
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
         Left            =   11040
         TabIndex        =   34
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "MRP Amount"
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
         Left            =   3360
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   14535
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmSalesReturn.frx":0000
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "frmSalesReturn.frx":0014
         TabIndex        =   73
         Top             =   240
         Width           =   14295
      End
   End
End
Attribute VB_Name = "frmSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset, rec As DAO.Recordset, rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset, rec5 As Recordset, rec6 As Recordset, rec7 As Recordset, tempamount, INV_TYPE, temp_discount_amount, TEMP_DEBTOR_GROUPID, DEL_CRNOTE, EDIT_INVOICE
Private Sub cboitemname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Me.txtfreight.SetFocus
    End If
    If KeyCode = 13 Then
        SEARCHWORD = Trim(Me.cboitemname.Text)
        frmproductlist.Show vbModal
    End If
End Sub
Private Sub cboParty_Change()
    Dim TEMP_SUBFIX
    Set rec1 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not IsNull(rec1!Address1) Then
        Me.LblAdr1.Caption = rec1("Address1")
    Else
        Me.LblAdr1.Caption = ""
    End If

    Me.cboParty.ToolTipText = Me.cboParty.ItemData(Me.cboParty.ListIndex)
    Set rec1 = db.OpenRecordset("select * from PartyDr where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not rec1.EOF Then
        If Not IsNull(rec1("Tin")) Then
            Me.txtTin.Text = rec1("Tin")
        Else
            Me.txtTin.Text = ""
        End If
        'Me.txttradediscount.Text = rec1("Discount")
        Me.txtcrlimit.Text = Format(rec1("CrLimit"), "########0.00")
    Else
        Me.txtTin.Text = ""
        Me.txttradediscount.Text = 0
        Me.txtcrlimit.Text = "0.00"
    End If

    Set rec2 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not rec2.EOF Then
        TEMP_DEBTOR_GROUPID = rec2("GroupId")
    End If
End Sub
Private Sub cboparty_Click()
    cboParty_Change
End Sub
Private Sub cboparty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmPartyDr.Show vbModal
    End If
    If KeyCode = 13 Then
        Me.txtLrno.SetFocus
    End If
End Sub

Private Sub cboprrate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtsalerate.SetFocus
    End If
End Sub

Private Sub cbosize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbounit.SetFocus
    End If
End Sub

Private Sub cbounit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtsalerate.SetFocus
    End If
End Sub

Private Sub CmdDelete_Click()
    db.Execute ("delete * from TempCreditNote")
    Data1.Refresh
    Me.txtInvNo.Locked = False
    Me.Label23.Enabled = True
    Me.txttotalgross.Text = "0.00"
    Me.txtChalanNo.Text = ""
    Me.txtTotalqty.Text = ""
    Me.TxtRoundup.Text = "0.00"

    Me.txtLrno.Text = ""
    Me.txttotaltradediscount.Text = "0.00"
    Me.txttotalspecialdiscount.Text = "0.00"


    Me.TxtVat.Text = 0
    Me.TxtVatAmount.Text = "0.00"
    Me.TxtNetAmount.Text = "0.00"
    Me.txtGrandtotal.Text = "0.00"
    Me.cboParty.ListIndex = 0
    Me.txtInvNo.SetFocus
    DEL_CRNOTE = "Y"
End Sub

Private Sub CmdEdit_Click()
    db.Execute ("delete * from TempCreditNote")
    Data1.Refresh
    Me.txtInvNo.Locked = False
    Me.Label23.Enabled = True
    Me.txttotalgross.Text = "0.00"
    Me.txtChalanNo.Text = ""
    Me.txtTotalqty.Text = ""
    Me.TxtRoundup.Text = "0.00"
    Me.txtLrno.Text = ""
    Me.txttotaltradediscount.Text = "0.00"
    Me.txttotalspecialdiscount.Text = "0.00"
    Me.TxtVatAmount.Text = "0.00"
    Me.TxtNetAmount.Text = "0.00"
    Me.txtGrandtotal.Text = "0.00"
    Me.cboParty.ListIndex = 0
    Me.txtInvNo.SetFocus
End Sub
Private Sub cmdprint_Click()
    FRMSALESRETURNPRINT.Show vbModal
End Sub

Private Sub CmdSave_Click()
On Error GoTo errtrap
  ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then

        X = NumberToWord(Trim(str(Round(Val(Me.txtGrandtotal.Text)))))

        Set rec1 = db.OpenRecordset("select * from Salesreturnhead where invno=" & Val(Me.txtInvNo.Text))
        If rec1.EOF Then
            temp_accid = 0
            db.Execute ("insert into Salesreturnhead (InvNo,InvDate,ChalanNo,ChalanDate,AccId,LrNo,Party,TotalQty,TotalGross,TradeDiscount,SpecialDiscount,VatAmount,Net,RndUp,GrandTotal,AmountInText,Freight,MrpAmount)values(" & Me.txtInvNo & ",'" & Me.txtInvDate.Text & "','" & Me.txtChalanNo.Text & "','" & Me.txtChalanDate.Text & "'," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ",'" & Me.txtLrno.Text & "','" & Me.cboParty.Text & "'," & Val(Me.txtTotalqty.Text) & "," & Val(Me.txttotalgross.Text) & "," & Val(Me.txttotaltradediscount.Text) & "," & Val(Me.txttotalspecialdiscount.Text) & "," & Val(Me.TxtVatAmount.Text) & "," & Val(Me.TxtNetAmount.Text) & "," & Val(Me.TxtRoundup.Text) & "," & Val(Me.txtGrandtotal.Text) & ",'" & AmountInText & "'," & Val(Me.txtfreight.Text) & "," & Val(Me.txtmrpamount.Text) & ")")
        Else
            temp_accid = rec1("AccId")
            db.Execute ("Update Salesreturnhead set InvDate='" & Me.txtInvDate.Text & "',ChalanNo='" & Me.txtChalanNo.Text & "',ChalanDate='" & Me.txtChalanDate.Text & "',Party='" & Me.cboParty.Text & "',AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ",TotalQty=" & Me.txtTotalqty.Text & ",TotalGross=" & Me.txttotalgross.Text & ",TradeDiscount=" & Val(Me.txttotaltradediscount.Text) & ",SpecialDiscount=" & Val(Me.txttotalspecialdiscount.Text) & ",VatAmount=" & Val(Me.TxtVatAmount.Text) & ",Net=" & Val(Me.TxtNetAmount.Text) & ",RndUp=" & Val(Me.TxtRoundup.Text) & ",GrandTotal=" & Val(Me.txtGrandtotal.Text) & ",AmountInText='" & X & "',Freight=" & Val(Me.txtfreight.Text) & ",MrpAmount=" & Val(Me.txtmrpamount.Text) & " where InvNo=" & Val(Me.txtInvNo.Text))
        End If
        '----------------------Checking &  Update previous entry---------------
        Set rec2 = db.OpenRecordset("select * from Salesreturndetails where InvNo=" & Me.txtInvNo.Text)
        If Not rec2.EOF Then
            While Not rec2.EOF
                stockqty = rec2("Qty") + rec2("Free_qty")
                db.Execute ("update stock set Qty=Qty-" & stockqty & " where ProductCode=" & rec2("ProductCode"))
                db.Execute ("update stockdetails set Qty=Qty-" & stockqty & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                rec2.MoveNext
            Wend
            db.Execute ("delete * from Salesreturndetails where InvNo=" & Val(Me.txtInvNo.Text))
        End If
        '-----------New Stock Entry------------
        Set rec2 = db.OpenRecordset("select * from TempCreditNote")
        If Not rec2.EOF Then
            While Not rec2.EOF
                db.Execute ("insert into Salesreturndetails (InvNo,Itemname,Units,MRP,SaleRate,Qty,Gross,SpecialDiscount,Tradediscount,DiscountAmount,Vat,VatAmount,Net,ProductCode,Free_Qty,Tax_type,mfgdate,expdate,batchno) values(" & Me.txtInvNo.Text & ",'" & Replace(rec2("Itemname"), "'", "''") & "','" & rec2("Units") & "'," & rec2("MRP") & "," & rec2("SaleRate") & "," & rec2("Qty") & "," & rec2("Gross") & "," & rec2("SpecialDiscount") & "," & rec2("Tradediscount") & "," & rec2("DiscountAmount") & "," & rec2("Vat") & "," & rec2("VatAmount") & "," & rec2("Net") & "," & rec2("ProductCode") & "," & rec2("Free_Qty") & ",'" & rec2("Tax_type") & "','" & rec2("mfgdate") & "','" & rec2("expdate") & "','" & rec2("batchno") & "')")
                stockqty = rec2("Qty") + rec2("Free_qty")
                db.Execute ("update stock set Qty=Qty+" & stockqty & " where ProductCode=" & rec2("ProductCode"))
                db.Execute ("update stockdetails set Qty=Qty+" & stockqty & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                rec2.MoveNext
            Wend
        End If
        
        db.Execute ("delete * from TempCreditNote")
        Data1.Refresh
        Me.txtInvNo.Text = Val(Me.txtInvNo.Text) + 1
        BillType = ""
        Me.txttotalgross.Text = "0.00"
        Me.txtChalanNo.Text = ""
        Me.txtTotalqty.Text = ""
        Me.TxtRoundup.Text = "0.00"
        Me.txtLrno.Text = ""
        Me.txttotalspecialdiscount.Text = "0.00"
        Me.txttotalspecialdiscount.Text = "0.00"
        Me.txtmrpamount.Text = "0.00"
        Me.TxtVat.Text = 0
        Me.TxtVatAmount.Text = "0.00"
        Me.TxtNetAmount.Text = "0.00"
        Me.txtGrandtotal.Text = "0.00"
        Me.txtInvNo.Locked = True
        Me.txtfreight.Text = "0.00"
        Me.cboParty.ListIndex = 0
        Me.txtInvDate.SetFocus
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Me.txttotalspecialdiscount.Text = Format(Val(Me.txttotalspecialdiscount.Text) - Val(Me.DBGrid1.Columns(13)), "#######0.00")
    Me.txttotaltradediscount.Text = Val(Me.txttotaltradediscount.Text) - Val(Me.DBGrid1.Columns(12))
    Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) - Val(Me.DBGrid1.Columns(10)), "######0.00")
    Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.DBGrid1.Columns(7)), "######0.00")
    Me.txtmrpamount.Text = Val(Me.txtmrpamount.Text) - ((Val(Me.DBGrid1.Columns(3))) * Val(Me.DBGrid1.Columns(5)))
    Me.txtTotalqty.Text = Val(Me.txtTotalqty.Text) - Val(Me.DBGrid1.Columns(3)) - Val(Me.DBGrid1.Columns(4))
    Me.TxtNetAmount.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.txttotaltradediscount.Text) - Val(Me.txttotalspecialdiscount.Text) + Val(Me.TxtVatAmount.Text), "#########0.00")
    Me.TxtRoundup.Text = Format(Round(Me.TxtNetAmount.Text) - Val(Me.TxtNetAmount.Text), "##0.00")
    Me.txtGrandtotal.Text = Round(Me.TxtNetAmount.Text)
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    FORMNAME = "SalesReturn"
    If usertype = "Admin" Then
        Me.CmdEdit.Enabled = True
        Me.CmdDelete.Enabled = True
    Else
        Me.CmdEdit.Enabled = False
        Me.CmdDelete.Enabled = False
    End If
    formid = 100
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Me.txtInvDate.Text = Format(Date, "dd/mm/yyyy")
    Me.txtChalanDate.Text = Format(Date, "dd/mm/yyyy")

    Set rec1 = db.OpenRecordset("select max(InvNo) as max_slno from Salesreturnhead")
    If Not IsNull(rec1!max_slno) Then
        Me.txtInvNo.Text = rec1!max_slno + 1
    Else
        Me.txtInvNo.Text = 1
    End If

    Set rec = db.OpenRecordset("select * from LedgerMaster where Groupname Like 'Sundry Debtor' or Groupname Like 'Cash-In-Hand'")
    While Not rec.EOF
        Me.cboParty.AddItem (rec("Accname"))
        Me.cboParty.ItemData(Me.cboParty.NewIndex) = rec("Accid")
        rec.MoveNext
    Wend
    If Me.cboParty.ListCount > 0 Then
        Me.cboParty.ListIndex = 0
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FORMNAME = ""
    formid = 0
    db.Execute ("delete * from TempCreditNote")
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
       
            Me.txtmrpamount.Text = Format(Val(Me.txtmrpamount.Text) + (Val(Me.txtQty.Text)) * Val(Me.txtmrp.Text), "######0.00")
            
            TradeDiscount = Round(Val(Me.txtgross.Text) * (Val(Me.txttradediscount.Text) / 100), 2)
           
            SpecialDiscount = Round((Val(Me.txtgross.Text) - TradeDiscount) * (Val(Me.txtspecialdiscount.Text) / 100), 2)
       
            Me.txtTotalqty.Text = Val(Me.txtTotalqty.Text) + Val(Me.txtQty.Text) + Val(Me.txtfree.Text)
            Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) + Val(Me.txtgross.Text), "######0.00")
            Me.txttotaltradediscount.Text = Format(Val(Me.txttotaltradediscount.Text) + TradeDiscount, "######0.00")
            Me.txttotalspecialdiscount.Text = Format(Val(Me.txttotalspecialdiscount.Text) + SpecialDiscount, "########0.00")
            Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) + Val(Me.txttaxamount.Text), "#####0.00")
            Me.TxtNetAmount.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.txttotaltradediscount.Text) - Val(Me.txttotalspecialdiscount.Text) + Val(Me.TxtVatAmount.Text), "#########0.00")
            Me.TxtRoundup.Text = Format(Round(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text) - Val(Me.TxtNetAmount.Text), "##0.00")
            Me.txtGrandtotal.Text = Round(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text))
            db.Execute ("insert into TempCreditNote (ItemName,ProductCode,Units,Qty,SaleRate,Gross,TradeDiscount,SpecialDiscount,DiscountAmount,MRP,Vat,VatAmount,Net,Free_Qty,Tax_type,mfgdate,expdate,batchno) values('" & Replace(Me.cboitemname.Text, "'", "''") & "'," & Me.txtproductcode.Text & ",'" & Me.cbounit.Text & "'," & Val(Me.txtQty.Text) & "," & Val(Me.txtsalerate.Text) & "," & Me.txtgross.Text & "," & Me.txttradediscount.Text & "," & Me.txtspecialdiscount.Text & "," & TradeDiscount + SpecialDiscount & "," & Val(Me.txtmrp.Text) & "," & Me.TxtVat.Text & "," & Me.txttaxamount.Text & "," & Me.txtamount.Text & "," & Me.txtfree.Text & ",'" & Me.txttaxtype.Text & "','" & Me.txtmfgdate.Text & "','" & Me.txtexpdate.Text & "','" & Me.cbobatch.Text & "')")
            Me.Data1.Refresh

            Me.txtQty.Text = 0
            Me.txtfree.Text = 0
            Me.txtmrp.Text = "0.00"
            Me.txttradediscount.Text = 0
            Me.txtspecialdiscount.Text = 0
            Me.txttaxamount.Text = "0.00"
            Me.txtpack.Text = 0
            TradeDiscount = 0
            SpecialDiscount = 0
           
            Me.cboitemname.SetFocus
    End If

End Sub

Private Sub txtChalanDate_GotFocus()
    Me.txtChalanDate.SelStart = 0
    Me.txtChalanDate.SelLength = Len(Me.txtChalanDate.Text)
End Sub

Private Sub txtChalanDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboParty.SetFocus
    End If
End Sub

Private Sub txtChalanNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtChalanDate.SetFocus
    End If
End Sub

Private Sub txtfree_GotFocus()
    Me.txtfree.SelStart = 0
    Me.txtfree.SelLength = Len(Me.txtfree.Text)
End Sub

Private Sub txtfree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtsalerate.SetFocus
    End If
End Sub

Private Sub txtfreight_GotFocus()
    Me.txtfreight.SelStart = 0
    Me.txtfreight.SelLength = Len(Me.txtfreight.Text)
End Sub

Private Sub txtfreight_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtRoundup.Text = Format(Round(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text)) - (Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text)), "##0.00")
        Me.txtGrandtotal.Text = Round(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text))
        Me.TxtRoundup.SetFocus
    End If
End Sub

Private Sub txtGrandtotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub

Private Sub txtInvDate_GotFocus()
    Me.txtInvDate.SelStart = 0
    Me.txtInvDate.SelLength = Len(Me.txtInvDate.Text)
End Sub
Private Sub txtInvdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtChalanNo.SetFocus
    End If
End Sub
Private Sub txtInvNo_GotFocus()
    Me.txtInvNo.SelStart = 0
    Me.txtInvNo.SelLength = Len(Me.txtInvNo.Text)
End Sub
Private Sub txtInvno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec = db.OpenRecordset("select * from Salesreturnhead where InvNo=" & Me.txtInvNo.Text)
        If Not rec.EOF Then
            '-----------Finding Name----------
            '            Set rec5 = db.OpenRecordset("select * from Partydr where accid=" & rec("Accid"))
            '            If Not rec5.EOF Then
            '                Set rec6 = db.OpenRecordset("select * from ZoneMaster where slno=" & rec5("ZoneCode"))
            '                If Not rec6.EOF Then
            '                    Me.CboZone.Text = Trim(rec6("ZoneName"))
            '                End If
            '            End If
            Set rec5 = db.OpenRecordset("select AccName from LedgerMaster where AccID=" & rec("AccId"))
            If Not rec5.EOF Then
                Me.cboParty.Text = Trim(rec5("AccName"))
            End If
            '---------------------------------
            db.Execute ("delete * from TempCreditNote")
            Set rec2 = db.OpenRecordset("select * from Salesreturndetails where InvNo=" & Me.txtInvNo.Text)
            If Not rec2.EOF Then
                While Not rec2.EOF
                    db.Execute ("insert into TempCreditNote (ItemName,ProductCode,Units,Qty,SaleRate,Gross,TradeDiscount,SpecialDiscount,DiscountAmount,MRP,Vat,VatAmount,Net,Free_Qty,Tax_type,mfgdate,expdate,batchno) values('" & Replace(rec2("ItemName"), "'", "''") & "'," & rec2("ProductCode") & ",'" & rec2("Units") & "'," & rec2("Qty") & "," & rec2("SaleRate") & "," & rec2("Gross") & "," & rec2("TradeDiscount") & "," & rec2("SpecialDiscount") & "," & rec2("DiscountAmount") & "," & rec2("MRP") & "," & rec2("Vat") & "," & rec2("VatAmount") & "," & rec2("Net") & "," & rec2("Free_Qty") & ",'" & rec2("Tax_type") & "','" & rec2("mfgdate") & "','" & rec2("expdate") & "','" & rec2("batchno") & "')")
                    rec2.MoveNext
                Wend
                Data1.Refresh
            End If
            Me.txtInvDate.Text = rec("InvDate")
            Me.txtChalanNo.Text = rec("ChalanNo")
            If Not IsNull(rec!CHALANDATE) Then
                Me.txtChalanDate.Text = rec("ChalanDate")
            Else
                Me.txtChalanDate.Text = "__/__/____"
            End If

            Me.txtLrno.Text = rec("LrNo")
            Me.txtTotalqty.Text = Format(rec("TotalQty"), "########0.00")
            Me.txttotalgross.Text = Format(rec("totalGross"), "#########0.00")
            Me.txtmrpamount.Text = Format(rec("MrpAmount"), "#######0.00")
            Me.txttotaltradediscount.Text = rec("tradediscount")
            Me.txttotalspecialdiscount.Text = Format(rec("specialdiscount"), "###########0.00")
            Me.TxtVatAmount.Text = Format(rec("VatAmount"), "##########0.00")
            Me.TxtNetAmount.Text = Format(rec("Net"), "###########0.00")
            Me.TxtRoundup.Text = rec("RndUp")
            Me.txtGrandtotal.Text = Format(rec("GrandTotal"), "################0.00")
            Me.txtInvDate.SetFocus
        End If

        '----------------------------------------Delete Invoice----------------------------
        If DEL_CRNOTE = "Y" Then
            ans = MsgBox("Confirm Delete?", vbYesNo)
            If ans = 6 Then
                '----------Update Stock--------------------------
                Set rec2 = db.OpenRecordset("select * from Salesreturndetails where InvNo=" & Me.txtInvNo.Text)
                If Not rec2.EOF Then
                    While Not rec2.EOF
                        stockqty = rec2("Qty") + rec2("Free_qty")
                        db.Execute ("update stock set Qty=Qty-" & stockqty & " where ProductCode=" & rec2("ProductCode"))
                        db.Execute ("update stockdetails set Qty=Qty-" & stockqty & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                        rec2.MoveNext
                    Wend
                    db.Execute ("delete * from Salesreturndetails where InvNo=" & Me.txtInvNo.Text)
                End If

                db.Execute ("delete * from Salesreturnhead where InvNo=" & Val(Me.txtInvNo.Text))
                DEL_CRNOTE = ""
                db.Execute ("delete * from TempCreditNote")
                Data1.Refresh
                Me.txtInvNo.Locked = True

                Me.txttotalgross.Text = "0.00"
                stockqty = 0
                Me.txtChalanNo.Text = ""
                Me.txtTotalqty.Text = ""
                Me.TxtRoundup.Text = "0.00"
                Me.txtLrno.Text = ""
                Me.TxtVat.Text = 0
                Me.TxtVatAmount.Text = "0.00"
                Me.TxtNetAmount.Text = "0.00"
                Me.txtGrandtotal.Text = "0.00"
                Me.cboParty.ListIndex = 0
                Me.txtInvNo.Text = 0
                Me.txtInvNo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtLrno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboitemname.SetFocus
    End If
End Sub

Private Sub txtmrp_GotFocus()
Me.txtmrp.SelStart = 0
Me.txtmrp.SelLength = Len(Me.txtmrp.Text)
End Sub

Private Sub txtmrp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txttradediscount.SetFocus
End If
End Sub

Private Sub txtNetamount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtRoundup.SetFocus
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
        Me.txtGrandtotal.Text = Format(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text) + Val(Me.TxtRoundup.Text), "###########0.00")
        txtGrandtotal.SetFocus
    End If
End Sub

Private Sub txtsalerate_GotFocus()
    Me.txtsalerate.SelStart = 0
    Me.txtsalerate.SelLength = Len(Me.txtsalerate.Text)
End Sub

Private Sub txtsalerate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("Select Salerate from ItemMaster where ProductCode=" & Me.txtproductcode.Text)
        If Not IsNull(rec1!SaleRate) Then
            If rec1!SaleRate <> Val(Me.txtsalerate.Text) Then
                ans = MsgBox("Update The Sale Rate!", vbYesNo)
                If ans = 6 Then
                    db.Execute ("Update ItemMaster set SaleRate=" & Val(Me.txtsalerate.Text) & " where ProductCode=" & Val(Me.txtproductcode.Text))
                Else
                    Me.txttradediscount.SetFocus
                End If
            End If
        End If
        Me.txtgross.Text = Format(Val(Me.txtQty.Text) * Val(Me.txtsalerate.Text), "######0.00")
        Me.txtmrp.SetFocus
    End If
End Sub

Private Sub txtspecialdiscount_GotFocus()
Me.txtspecialdiscount.SelStart = 0
Me.txtspecialdiscount.SelLength = Len(Me.txtspecialdiscount.Text)
End Sub

Private Sub txtspecialdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.TxtVat.SetFocus
End If
End Sub

Private Sub txttradediscount_GotFocus()
    Me.txttradediscount.SelStart = 0
    Me.txttradediscount.SelLength = Len(Me.txttradediscount.Text)
End Sub

Private Sub txttradediscount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtspecialdiscount.SetFocus
    End If
End Sub

Private Sub txtvat_GotFocus()
    Me.TxtVat.SelStart = 0
    Me.TxtVat.SelLength = Len(Me.TxtVat.Text)
End Sub

Private Sub TxtVat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

    TradeDiscount = Round(Val(Me.txtgross.Text) * (Val(Me.txttradediscount.Text) / 100), 2)
    SpecialDiscount = Round((Val(Me.txtgross.Text) - TradeDiscount) * (Val(Me.txtspecialdiscount.Text) / 100), 2)
            
            If Me.txttaxtype.Text = "MRP" Then
                Me.txttaxamount.Text = Round((Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text))) * (Val(Me.TxtVat.Text) / 100), 2)
            ElseIf Me.txttaxtype.Text = "SALES" Then
               'Me.txttaxamount.Text = Round((Val(Me.txtgross.Text) - TradeDiscount - SpecialDiscount) * (Val(Me.TxtVat.Text) / 100), 2)
               temp_rate = Val(Me.txtgross.Text) - TradeDiscount - SpecialDiscount
               Me.txttaxamount.Text = Val(temp_rate) - Format((temp_rate / ((Me.TxtVat.Text / 100) + 1)), "########0.00")
               Me.txtgross.Text = Format((temp_rate / ((Me.TxtVat.Text / 100) + 1)), "########0.00")
            ElseIf Me.txttaxtype.Text = "INCLUSIVE MRP" Then
                Me.txttaxamount.Text = Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) - Format((Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) / ((Me.TxtVat.Text / 100) + 1)), "########0.00")
            Else
                Me.txttaxamount.Text = "0.00"
            End If
            Me.txtamount.Text = Round(Val(Me.txtgross.Text) - TradeDiscount - SpecialDiscount + Val(Me.txttaxamount.Text), 2)
            Me.txtamount.SetFocus
End If

End Sub
