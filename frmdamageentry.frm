VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmdamageentry 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Purchase Return [ Debit Note]"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   44
      Top             =   1020
      Width           =   13815
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
         Height          =   285
         Left            =   1080
         TabIndex        =   70
         Top             =   360
         Width           =   3975
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
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   64
         Text            =   "cbobatch"
         Top             =   900
         Width           =   1500
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   4440
         TabIndex        =   2
         Text            =   "0"
         Top             =   900
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
         Left            =   5280
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   48
         Top             =   900
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   900
         Width           =   735
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   6720
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   900
         Width           =   855
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   8520
         TabIndex        =   47
         Text            =   "0.00"
         Top             =   900
         Width           =   855
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
         Left            =   3720
         TabIndex        =   1
         Text            =   "0"
         Top             =   900
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
         Left            =   60
         TabIndex        =   46
         Top             =   360
         Width           =   915
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   900
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Text            =   "0"
         Top             =   900
         Width           =   615
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   9480
         TabIndex        =   6
         Text            =   "0"
         Top             =   900
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
         TabIndex        =   45
         Top             =   120
         Width           =   1335
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   10200
         TabIndex        =   7
         Text            =   "0"
         Top             =   900
         Width           =   615
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   12600
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   900
         Width           =   975
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   11760
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   900
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtexpdate 
         Height          =   315
         Left            =   2700
         TabIndex        =   65
         Top             =   900
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
      Begin MSMask.MaskEdBox txtmfgdate 
         Height          =   315
         Left            =   1620
         TabIndex        =   66
         Top             =   900
         Width           =   1035
         _ExtentX        =   1826
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
         Left            =   60
         TabIndex        =   69
         Top             =   660
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
         Left            =   1620
         TabIndex        =   68
         Top             =   660
         Width           =   915
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
         Left            =   2700
         TabIndex        =   67
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label33 
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
         Left            =   1080
         TabIndex        =   62
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label32 
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
         Left            =   4440
         TabIndex        =   61
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label16 
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
         Left            =   5280
         TabIndex        =   60
         Top             =   660
         Width           =   1215
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
         Left            =   7680
         TabIndex        =   59
         Top             =   660
         Width           =   855
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
         Left            =   6720
         TabIndex        =   58
         Top             =   660
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
         Left            =   8520
         TabIndex        =   57
         Top             =   660
         Width           =   1215
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
         Left            =   3720
         TabIndex        =   56
         Top             =   660
         Width           =   615
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
         Left            =   11040
         TabIndex        =   55
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label15 
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
         Left            =   9480
         TabIndex        =   54
         Top             =   660
         Width           =   615
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
         Left            =   6000
         TabIndex        =   53
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label13 
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
         Left            =   10200
         TabIndex        =   52
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label12 
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
         Left            =   12600
         TabIndex        =   51
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label10 
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
         Left            =   11760
         TabIndex        =   49
         Top             =   660
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   13815
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
         Left            =   3360
         TabIndex        =   43
         Top             =   1440
         Width           =   975
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
         Left            =   2280
         TabIndex        =   30
         Top             =   1440
         Width           =   1000
      End
      Begin VB.TextBox TXTCST 
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
         TabIndex        =   29
         Text            =   "0"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtEtax 
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
         TabIndex        =   28
         Text            =   "0"
         Top             =   600
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
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   240
         Width           =   1455
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
         Left            =   6480
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox TxtCstAmount 
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
         Left            =   6480
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtEtAmount 
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
         Left            =   6480
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   600
         Width           =   1695
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
         Left            =   6480
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   240
         Width           =   1695
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
         Left            =   11760
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   1320
         Width           =   1695
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
         Left            =   11760
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   960
         Width           =   1695
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
         Left            =   1200
         TabIndex        =   20
         Top             =   1440
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   1680
         Width           =   1695
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
         TabIndex        =   18
         Top             =   1440
         Width           =   1000
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
         Left            =   11760
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   600
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
         Left            =   11760
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label21 
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
         Left            =   9960
         TabIndex        =   42
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         Left            =   9960
         TabIndex        =   41
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label22 
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
         Left            =   4560
         TabIndex        =   40
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "C.S.T Amount"
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
         Left            =   4560
         TabIndex        =   39
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "Entry Tax %"
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
         TabIndex        =   38
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label25 
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
         Left            =   4560
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "C.S.T %"
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
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Grand Total"
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
         Left            =   9960
         TabIndex        =   35
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "E. Tax Amount"
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
         Left            =   4560
         TabIndex        =   34
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label29 
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
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   9960
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Mrp Amount"
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
         Left            =   9960
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Temppurchasereturn"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   13815
      Begin VB.TextBox txtslno 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtStockindate 
         Height          =   315
         Left            =   11880
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
         Left            =   10920
         TabIndex        =   14
         Top             =   240
         Width           =   855
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
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmdamageentry.frx":0000
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "frmdamageentry.frx":0014
      TabIndex        =   63
      Top             =   2400
      Width           =   13815
   End
End
Attribute VB_Name = "frmdamageentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset, rec5 As Recordset, TAXAMOUNT, temp_lr_balance, temp_lr_slno, TEMP_GROUPID, temp_discount_amount, PURCHASE_RETURN_DELETE
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
Attribute PURCHASE_RETURN_DELETE.VB_VarUserMemId = 1073938432
Private Sub cboRate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtQty.SetFocus
    End If
End Sub

Private Sub cbounit_type_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtQty.SetFocus
    End If
End Sub


Private Sub cboitemname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        frmproductlist.Show vbModal
    End If
    If KeyCode = 27 Then
        Me.TxtEtax.SetFocus
    End If
End Sub

Private Sub CmdDelete_Click()
    Me.txtslno.Locked = False
    PURCHASE_RETURN_DELETE = "Y"
    db.Execute ("delete * from Temppurchasereturn")
    Data1.Refresh
   
    Me.txtTotalqty.Text = "0.00"
    Me.txttotalgross.Text = "0.00"
    Me.txtcst.Text = "0.00"
    Me.txtcstamount.Text = "0.00"
    Me.TxtEtax.Text = "0.00"
    Me.txtetamount.Text = "0.00"

    Me.TxtVat.Text = "0.00"
    Me.TxtVatAmount.Text = "0.00"

    Me.TxtNetAmount.Text = "0.00"
    Me.TxtRoundup.Text = "0.00"

    Me.txtGrandtotal.Text = "0.00"
    Me.txtslno.SetFocus

End Sub
Private Sub CmdEdit_Click()
    Me.txtslno.Locked = False
    db.Execute ("delete * from Temppurchasereturn")
    Data1.Refresh
    
    Me.txtTotalqty.Text = "0.00"
    Me.txttotalgross.Text = "0.00"
    Me.txtcst.Text = "0.00"
    Me.txtcstamount.Text = "0.00"
    Me.TxtEtax.Text = "0.00"
    Me.txtetamount.Text = "0.00"

    Me.TxtVat.Text = "0.00"
    Me.TxtVatAmount.Text = "0.00"

    Me.TxtNetAmount.Text = "0.00"
    Me.TxtRoundup.Text = "0.00"

    Me.txtGrandtotal.Text = "0.00"
    Me.txtslno.SetFocus
End Sub

Private Sub cmdprint_Click()
frmdamageprint.Show 0
End Sub

Private Sub CmdSave_Click()
On Error GoTo errtrap
    ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then

        temp_grandtotal = Me.txtGrandtotal.Text
        X = NumberToWord(Trim(str(Round(Val(temp_grandtotal)))))
        Set rec1 = db.OpenRecordset("select * from damageHead where Slno=" & Val(Me.txtslno.Text))
        If rec1.EOF Then
            AccId = 0
            db.Execute ("insert into DamageHead (Slno,damagedate,TotalQty,TotalGross,GrandTotal) values(" & Me.txtslno.Text & ",'" & Me.txtStockindate.Text & "'," & Me.txtTotalqty.Text & "," & Me.txttotalgross.Text & "," & Me.txtGrandtotal.Text & ")")
        Else
            
            db.Execute ("update DamageHead set damagedate='" & Me.txtStockindate.Text & "',TotalQty=" & Me.txtTotalqty.Text & ",TotalGross=" & Me.txttotalgross.Text & ",ETax=" & Me.TxtEtax.Text & ",ETaxAmount=" & Me.txtetamount.Text & ",CST=" & Me.txtcst.Text & ",CSTAmount=" & Me.txtcstamount.Text & ",VatAmount=" & Me.TxtVatAmount.Text & ",NetAmount=" & Me.TxtNetAmount.Text & ",RValue=" & Me.TxtRoundup.Text & ",GrandTotal=" & Me.txtGrandtotal.Text & " where Slno=" & Val(Me.txtslno.Text))
        End If

        'Check Existing Return (Update Stock)
        Set rec2 = db.OpenRecordset("select * from DamageDetails where Slno=" & Val(Me.txtslno.Text))
        If Not rec2.EOF Then
            While Not rec2.EOF
                Set rec3 = db.OpenRecordset("select * from Stock where productcode=" & rec2("Productcode"))
                If Not rec3.EOF Then
                    db.Execute ("update stock set Qty=Qty + " & rec2("Qty") & " where productcode=" & rec2("Productcode"))
                    db.Execute ("update stockdetails set Qty=Qty+" & rec2("Qty") & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                    db.Execute ("update damagestock set Qty=Qty-" & rec2("Qty") & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                End If
                rec2.MoveNext
            Wend
            db.Execute ("delete * from DamageDetails where SlNo=" & Val(Me.txtslno.Text))
        End If

        Set rec2 = db.OpenRecordset("select * from Temppurchasereturn")
        If Not rec2.EOF Then
            While Not rec2.EOF
                db.Execute ("insert into DamageDetails (Slno,ItemName,Units,Qty,MRP,PrRate,Amount,ProductCode,Vat,VatAmount,Net,mfgdate,expdate,batchno) values(" & Val(Me.txtslno.Text) & ",'" & Replace(rec2("ItemName"), "'", "''") & "','" & rec2("Units") & "'," & rec2("Qty") & "," & rec2("MRP") & "," & rec2("PrRate") & "," & rec2("Amount") & "," & rec2("ProductCode") & "," & rec2("Vat") & "," & rec2("VatAmount") & "," & rec2("Net") & ",'" & rec2("mfgdate") & "','" & rec2("expdate") & "','" & rec2("batchno") & "')")
                db.Execute ("update stock set Qty=Qty - " & (rec2("Qty") + rec2("Free_Qty")) & " where ProductCode=" & rec2("ProductCode"))
                db.Execute ("update stockdetails set Qty=Qty-" & (rec2("Qty") + rec2("Free_Qty")) & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                Set rec3 = db.OpenRecordset("select * from damagestock where productcode=" & rec2("Productcode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                    If Not rec3.EOF Then
                        db.Execute ("update damagestock set Qty=Qty + " & (rec2("Qty") + rec2("Free_Qty")) & " where productcode=" & rec2("Productcode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                    Else
                        Set rec1 = db.OpenRecordset("select * from itemmaster where productcode=" & rec2("productcode"))
                        If Not rec1.EOF Then
                            db.Execute ("insert into DamageStock (ProductType,itemtype,itemname,brand,barcode,size,MRP,PRate,Qty,ProductCode,Vat,Lose,UniyType,SaleRate,mfgdate,batchno,expdate,hsn) values('" & rec1("ProductType") & "','" & rec1("itemtype") & "','" & Replace(rec1("item"), "'", "''") & "','" & rec1("brand") & "','" & Replace(rec1("barcode"), "'", "''") & "','" & rec1("size") & "'," & rec1("mrp") & "," & rec1("purchaserate") & "," & rec2("qty") + rec2("free_qty") & "," & rec2("Productcode") & "," & rec2("vat") & "," & rec1("Lose") & ",'" & rec1("unittype") & "'," & rec1("SaleRate") & ",'" & rec2("mfgdate") & "','" & rec2("batchno") & "','" & rec2("expdate") & "','" & rec1("HSN") & "')")
                        End If
                    End If
                rec2.MoveNext
            Wend
        End If

        db.Execute ("delete * from Temppurchasereturn")
        Data1.Refresh
        AccId = 0
        Slno = 0
        prAccId = 0
        
        
        Me.txtslno.Locked = True
        Me.txtTotalqty.Text = "0.00"
        Me.txttotalgross.Text = "0.00"
        Me.txtcst.Text = "0.00"
        Me.txtcstamount.Text = "0.00"
        Me.TxtEtax.Text = "0.00"
        Me.txtetamount.Text = "0.00"
        Me.txtfreight.Text = "0.00"
        Me.TxtVatAmount.Text = "0.00"
        Me.TxtNetAmount.Text = "0.00"
        Me.TxtRoundup.Text = "0.00"
        Me.txtGrandtotal.Text = "0.00"
        PURCHASEDELETE = "N"
        Set rec3 = db.OpenRecordset("select max(Slno)  as max_no from DamageHead")
        If Not IsNull(rec3!max_no) Then
            Me.txtslno.Text = rec3!max_no + 1
        Else
            Me.txtslno.Text = 1
        End If
        Me.txtStockindate.SetFocus
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Me.txtTotalqty.Text = Format(Val(Me.txtTotalqty.Text) - Val(Me.DBGrid1.Columns(3)) - Val(Me.DBGrid1.Columns(5)), "############0.00")
    Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.DBGrid1.Columns(8)), "#########0.00")
    Me.txtmrpamount.Text = Format(Val(Me.txtmrpamount.Text) - (Val(Me.DBGrid1.Columns(3)) * Val(Me.DBGrid1.Columns(7))), "########0.00")
    Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) - Val(Me.DBGrid1.Columns(13)), "#######0.00")
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    FORMNAME = "DamageEntry"
    Data1.databasename = dbname
    formid = 3
    Me.Top = 0
    Me.Left = 0
    Me.txtStockindate.Text = Format(Date, "dd/mm/yyyy")

    Set rec1 = db.OpenRecordset("select max(slno) as max_no from DamageHead")
    If Not IsNull(rec1!max_no) Then
        Me.txtslno.Text = rec1!max_no + 1
    Else
        Me.txtslno.Text = 1
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db.Execute ("delete * from Temppurchasereturn")
End Sub

Private Sub Label34_Click()

End Sub

Private Sub txtAmount_GotFocus()
    Me.txtamount.SelStart = 0
    Me.txtamount.SelLength = Len(Me.txtamount.Text)
End Sub
Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

    End If
End Sub
Private Sub TXTCST_GotFocus()
    Me.txtcst.SelStart = 0
    Me.txtcst.SelLength = Len(Me.txtcst.Text)
End Sub

Private Sub TXTCST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtcstamount.Text = Format((Val(Me.txttotalgross.Text) / 100) * Me.txtcst.Text, "##########0.00")
        Me.txtfreight.SetFocus
    End If
End Sub

Private Sub TxtCstAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

    End If
End Sub
Private Sub txtentrytax_GotFocus()
    Me.TxtEtax.SelStart = 0
    Me.TxtEtax.SelLength = Len(Me.TxtEtax.Text)
End Sub
Private Sub txtentrytax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtcst.SetFocus
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

Private Sub TxtEtax_GotFocus()
    Me.TxtEtax.SelStart = 0
    Me.TxtEtax.SelLength = Len(Me.TxtEtax.Text)
End Sub
Private Sub TxtEtax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtetamount.Text = Format((Val(Me.txttotalgross.Text) / 100) * Me.TxtEtax.Text, "#######0.00")
        Me.txtcst.SetFocus
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
        Me.TxtNetAmount.Text = Format(Val(Me.txttotalgross.Text) + Val(Me.txtetamount.Text) + Val(Me.txtcstamount.Text) + Val(Me.TxtVatAmount.Text) + Val(Me.txtfreight.Text), "#########0.00")
        Me.TxtNetAmount.SetFocus
    End If
End Sub

Private Sub txtGrandtotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cmdSave.SetFocus
    End If
End Sub
Private Sub txtInvdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Me.cboSupplier.SetFocus
    End If
End Sub

Private Sub txtInvNo_GotFocus()
    'Me.txtInvno.SelStart = 0
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
        Me.txtdiscount.SetFocus
    End If
End Sub

Private Sub txtnet_GotFocus()
Me.txtnet.SelStart = 0
Me.txtnet.SelLength = Len(Me.txtnet.Text)
End Sub

Private Sub txtnet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        discount_amount = Round(Val(Me.txtamount.Text) * (Val(Me.txtdiscount.Text) / 100), 2)
        Special_Discount = Round((Val(Me.txtamount.Text) - discount_amount) * (Val(Me.txtspecialdiscount.Text) / 100), 2)
        db.Execute ("insert into Temppurchasereturn (ItemName,Units,Pack,Qty,Mrp,PrRate,Amount,ProductCode,Vat,VatAmount,Net,Usname,Free_Qty,Discount,SpDiscount,Discount_amount,mfgdate,expdate,batchno) values('" & Replace(Me.cboitemname.Text, "'", "''") & "','" & Me.cbounit.Text & "'," & Me.txtpack.Text & "," & (Val(Me.txtQty.Text)) & "," & Me.txtmrp.Text & "," & Val(Me.txtPrate.Text) & "," & Me.txtamount.Text & "," & Me.txtproductcode.Text & "," & Val(Me.TxtVat.Text) & "," & Me.txttaxamount.Text & "," & Me.txtnet.Text & ",'" & usname & "'," & Me.txtfree.Text & "," & Me.txtdiscount.Text & "," & Me.txtspecialdiscount.Text & "," & discount_amount + Special_Discount & ",'" & Me.txtmfgdate.Text & "','" & Me.txtexpdate.Text & "','" & Me.cbobatch.Text & "')")
        Data1.Refresh
        Me.txtTotalqty.Text = Val(Me.txtTotalqty.Text) + Val(Me.txtQty.Text) + Val(Me.txtfree.Text)
        Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) + Val(Me.txtamount.Text), "################0.00")
        
        Me.txtmrpamount.Text = Format(Val(Me.txtmrpamount.Text) + (Val(Me.txtmrp.Text) * Val(Me.txtQty.Text)), "########0.00")
        Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) + Val(Me.txttaxamount.Text), "########0.00")
        Me.txtmrp.Text = "0.00"
        Me.TxtVat.Text = 0
        Me.txttaxamount.Text = "0.00"
        Me.txtnet.Text = "0.0"
        Me.txtPrate.Text = "0.00"
        Me.txtamount.Text = "0.00"
        Me.txtdiscount.Text = 0
        discount_amount = 0
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
    If KeyCode = 13 Then
        db.Execute ("delete * from Temppurchasereturn")
        Set rec2 = db.OpenRecordset("select * from damageHead where Slno=" & Me.txtslno.Text)
        If Not rec2.EOF Then
            Me.txtStockindate.Text = rec2("damagedate")
            
            Me.txttotalgross.Text = Format(rec2("TotalGross"), "######0.00")
            'Me.txtmrpamount.Text = Format(rec2("TotalMrp"), "########0.00")
            Me.txtTotalqty.Text = Format(rec2("TotalQty"), "#####0.00")
            Me.txtcst.Text = rec2("CST")
            Me.txtcstamount.Text = Format(rec2("CSTAmount"), "#########0.00")
            Me.TxtVatAmount.Text = Format(rec2("VatAmount"), "#####0.00")
            Me.TxtNetAmount.Text = Format(rec2("NetAmount"), "######0.00")
            Me.txtGrandtotal.Text = Format(rec2("GrandTotal"), "######0.00")
            Me.TxtEtax.Text = rec2("ETax")
            Me.txtetamount.Text = Format(rec2("ETaxAmount"), "########0.00")
            Me.TxtRoundup.Text = rec2("RValue")

            'Me.txtfreight.Text = Format(rec2("Freight"), "#######0.00")
            
            Set rec1 = db.OpenRecordset("select * from damageDetails where SlNo=" & Me.txtslno.Text)
            If Not rec1.EOF Then
                While Not rec1.EOF
                    db.Execute ("insert into Temppurchasereturn (ItemName,Units,Qty,Mrp,PrRate,Amount,ProductCode,Vat,VatAmount,Net,mfgdate,expdate,batchno) values('" & Replace(rec1("ItemName"), "'", "''") & "','" & rec1("Units") & "'," & rec1("Qty") & "," & rec1("Mrp") & "," & rec1("PrRate") & "," & rec1("Amount") & "," & rec1("ProductCode") & "," & rec1("Vat") & "," & rec1("VatAmount") & "," & rec1("Net") & ",'" & rec1("mfgdate") & "','" & rec1("expdate") & "','" & rec1("batchno") & "')")
                    rec1.MoveNext
                Wend
                Me.txtslno.Locked = True
            End If
            Me.txtStockindate.SetFocus
        Else
            MsgBox "Not Found", vbCritical
        End If
        Me.Data1.Refresh
        '//--------------deleteing purchase-------------
        If PURCHASE_RETURN_DELETE = "Y" Then
            ans = MsgBox("delete this purchase Return?", vbYesNo)
            If ans = 6 Then
                '//Deleting Purchasereturn details-----------------
                Set rec2 = db.OpenRecordset("select * from damageDetails where Slno=" & Me.txtslno.Text)
                If Not rec2.EOF Then
                    While Not rec2.EOF
                        db.Execute ("update stock set Qty=Qty + " & rec2("Qty") & " where ProductCode=" & rec2("ProductCode"))
                        db.Execute ("update stockdetails set Qty=Qty+" & rec2("Qty") & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                        db.Execute ("update damagestock set Qty=Qty-" & rec2("Qty") & " where ProductCode=" & rec2("ProductCode") & " and mfgdate='" & rec2("mfgdate") & "' and expdate='" & rec2("expdate") & "' and batchno='" & rec2("batchno") & "'")
                        rec2.MoveNext
                    Wend
                    db.Execute ("delete * from damageDetails where SlNo=" & Me.txtslno.Text)
                End If
                db.Execute ("delete * from damageHead where Slno=" & Me.txtslno.Text)
            End If
            db.Execute ("delete * from Temppurchasereturn")
            Data1.Refresh
            AccId = 0
            prAccId = 0
            
            Me.txtslno.Locked = True
            Me.txtTotalqty.Text = "0.00"
            Me.txttotalgross.Text = "0.00"
            Me.txtcst.Text = "0.00"
            Me.txtcstamount.Text = "0.00"
            Me.TxtEtax.Text = "0.00"
            Me.txtetamount.Text = "0.00"
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
    Me.TxtVat.SetFocus
End If
End Sub

Private Sub txtStockindate_GotFocus()
    Me.txtStockindate.SelStart = 0
    Me.txtStockindate.SelLength = Len(Me.txtStockindate.Text)
End Sub
Private Sub txtStockindate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboitemname.SetFocus
    End If
End Sub
Private Sub TxtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtVat.SetFocus
    End If
End Sub


Private Sub txtvat_GotFocus()
    Me.TxtVat.SelStart = 0
    Me.TxtVat.SelLength = Len(Me.TxtVat.Text)
End Sub

Private Sub TxtVat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        discount_amount = Round(Val(Me.txtamount.Text) * (Val(Me.txtdiscount.Text) / 100), 2)
        Special_Discount = (Val(Me.txtamount.Text) - discount_amount) * (Val(Me.txtspecialdiscount.Text) / 100)
        
        If Me.txttaxtype.Text = "MRP" Then
            Me.txttaxamount.Text = Round((Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text))) * (Val(Me.TxtVat.Text) / 100), 2)
        End If
        If Me.txttaxtype.Text = "SALES" Then
            Me.txttaxamount.Text = Round((Val(Me.txtamount.Text) - discount_amount - Special_Discount) * (Val(Me.TxtVat.Text) / 100), 2)
        End If
        If Me.txttaxtype.Text = "INCLUSIVE MRP" Then
            Me.txttaxamount.Text = Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) - Format((Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) / ((Me.TxtVat.Text / 100) + 1)), "########0.00")
        End If
        Me.txtnet.Text = Format(Val(Me.txtamount.Text) - discount_amount - Special_Discount + Val(Me.txttaxamount.Text), "########0.00")
        Me.txtnet.SetFocus
End If



End Sub
