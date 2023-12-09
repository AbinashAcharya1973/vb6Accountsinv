VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInvoiceGR 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Invoice [Retail]"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   14715
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Height          =   3675
      Left            =   120
      TabIndex        =   57
      Top             =   3300
      Width           =   14535
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmInvoiceGR.frx":0000
         Height          =   3435
         Left            =   120
         OleObjectBlob   =   "frmInvoiceGR.frx":0014
         TabIndex        =   58
         Top             =   120
         Width           =   14295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E3E3E3&
      Height          =   1935
      Left            =   120
      TabIndex        =   48
      Top             =   6960
      Width           =   6015
      Begin VB.TextBox txtadd 
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
         Left            =   4200
         TabIndex        =   83
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox txttotalcase 
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
         TabIndex        =   82
         Text            =   "0"
         Top             =   480
         Width           =   1095
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
         Left            =   9660
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   540
         Width           =   1575
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
         Left            =   5280
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtcrlimit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
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
         Left            =   7500
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtbalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
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
         Left            =   7500
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "0.00"
         Top             =   540
         Width           =   1095
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
         Left            =   9660
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "0.00"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtGrandtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   1260
         Width           =   1695
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
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   120
         Width           =   1095
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   840
         Width           =   1095
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
         Left            =   840
         TabIndex        =   29
         Top             =   1500
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
         Left            =   3720
         TabIndex        =   32
         Top             =   1500
         Width           =   850
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H0000FFFF&
         Caption         =   "EDIT"
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
         Left            =   1800
         TabIndex        =   31
         Top             =   1500
         Width           =   850
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
         Left            =   4200
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   900
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   1200
         Width           =   1095
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   180
         Width           =   1695
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H000000FF&
         Caption         =   "DELETE"
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
         Left            =   2760
         TabIndex        =   30
         Top             =   1500
         Width           =   850
      End
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
         Left            =   9660
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Case/Pack"
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
         TabIndex        =   81
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Amount"
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
         Left            =   7620
         TabIndex        =   75
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Add /Ded"
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
         Left            =   2880
         TabIndex        =   72
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Left            =   6060
         TabIndex        =   71
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Cr.Limit"
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
         Left            =   6060
         TabIndex        =   70
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label29 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Less Trade Discount"
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
         Left            =   7200
         TabIndex        =   59
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
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
         Left            =   2880
         TabIndex        =   55
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Qty"
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
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Gross"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Round Up"
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
         Left            =   2880
         TabIndex        =   52
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
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
         Left            =   180
         TabIndex        =   51
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         Left            =   2880
         TabIndex        =   50
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Special Discount"
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
         Left            =   7620
         TabIndex        =   49
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E3E3E3&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   43
      Top             =   2280
      Width           =   14535
      Begin VB.TextBox TXTSTOCK 
         BackColor       =   &H00E3E3E3&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboitemname 
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   480
         Width           =   3735
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
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   480
         Width           =   735
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
         Left            =   7680
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   465
         Width           =   735
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
         TabIndex        =   9
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txttaxtype 
         BackColor       =   &H00E3E3E3&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   240
         Width           =   1215
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
         TabIndex        =   11
         Text            =   "0"
         Top             =   480
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
         TabIndex        =   17
         Text            =   "0"
         Top             =   480
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
         TabIndex        =   16
         Text            =   "0"
         Top             =   480
         Width           =   615
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
         TabIndex        =   15
         Text            =   "0"
         Top             =   480
         Width           =   615
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
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   480
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0"
         Top             =   480
         Width           =   735
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
         Left            =   8520
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   480
         Width           =   855
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
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cbounit 
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
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Text            =   "cbounit"
         Top             =   480
         Width           =   735
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
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "GST Amt"
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
         Height          =   255
         Left            =   12600
         TabIndex        =   79
         Top             =   240
         Width           =   735
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   7680
         TabIndex        =   78
         Top             =   225
         Width           =   855
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   4680
         TabIndex        =   77
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6960
         TabIndex        =   73
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "GST%"
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
         Height          =   255
         Left            =   11880
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   11160
         TabIndex        =   66
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CD"
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
         Height          =   255
         Left            =   10440
         TabIndex        =   65
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   9480
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   8520
         TabIndex        =   56
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   960
         TabIndex        =   47
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5400
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6120
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   13440
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E3E3E3&
      Height          =   2235
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   14535
      Begin VB.ComboBox cbopaymentmode 
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
         ItemData        =   "frmInvoiceGR.frx":244F
         Left            =   6360
         List            =   "frmInvoiceGR.frx":2451
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtorderno 
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
         Left            =   11340
         TabIndex        =   88
         Top             =   1440
         Width           =   1335
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
         Left            =   8520
         TabIndex        =   85
         Top             =   1200
         Width           =   495
      End
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
         Left            =   6360
         TabIndex        =   84
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cbocash 
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
         ItemData        =   "frmInvoiceGR.frx":2453
         Left            =   7440
         List            =   "frmInvoiceGR.frx":2455
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox CboZone 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtChalanDate 
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
         Left            =   12780
         TabIndex        =   4
         Top             =   720
         Width           =   1575
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
         Left            =   1440
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   1455
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
         Left            =   11340
         TabIndex        =   3
         Top             =   720
         Width           =   1335
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
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtTin 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   1575
      End
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
         Left            =   11340
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin VB.ComboBox CboInvType 
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
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   285
         Left            =   11340
         TabIndex        =   1
         Top             =   240
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
      Begin MSMask.MaskEdBox txtorderdate 
         Height          =   285
         Left            =   12780
         TabIndex        =   90
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label38 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Mode"
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
         Left            =   4920
         TabIndex        =   92
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Order NO - DT"
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
         Left            =   10020
         TabIndex        =   89
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash/Credit"
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
         Left            =   6360
         TabIndex        =   80
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
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
         TabIndex        =   61
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Inv No"
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
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
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
         Height          =   255
         Left            =   10020
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
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
         Left            =   10020
         TabIndex        =   40
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
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
         Left            =   6360
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
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
         Left            =   10020
         TabIndex        =   37
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax/Retail"
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
         Height          =   375
         Left            =   3360
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblAdr1 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         Top             =   1200
         Width           =   1935
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Temp_Invoice"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmInvoiceGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset, rec As DAO.Recordset, rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset, rec4 As Recordset, rec5 As Recordset, rec6 As Recordset, rec7 As Recordset, tempamount, INV_TYPE, temp_discount_amount, TEMP_DEBTOR_GROUPID, DEL_INVOICE, EDIT_INVOICE, dbl_click_edit As Boolean, temp_slno, temp_del_slno, EDIT_ITEM_SLNO
Attribute rec.VB_VarUserMemId = 1073938432
Attribute rec1.VB_VarUserMemId = 1073938432
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Attribute rec4.VB_VarUserMemId = 1073938432
Attribute rec5.VB_VarUserMemId = 1073938432
Attribute rec6.VB_VarUserMemId = 1073938432
Attribute rec7.VB_VarUserMemId = 1073938432
Attribute tempamount.VB_VarUserMemId = 1073938432
Attribute INV_TYPE.VB_VarUserMemId = 1073938432
Attribute temp_discount_amount.VB_VarUserMemId = 1073938432
Attribute TEMP_DEBTOR_GROUPID.VB_VarUserMemId = 1073938432
Attribute DEL_INVOICE.VB_VarUserMemId = 1073938432
Attribute EDIT_INVOICE.VB_VarUserMemId = 1073938432
Attribute dbl_click_edit.VB_VarUserMemId = 1073938432
Attribute temp_slno.VB_VarUserMemId = 1073938432
Attribute temp_del_slno.VB_VarUserMemId = 1073938432
Attribute EDIT_ITEM_SLNO.VB_VarUserMemId = 1073938432

Private Sub cbocash_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtInvDate.SetFocus
    End If
End Sub

Private Sub cboInvType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        INV_TYPE = Me.CboInvType.Text

        Set rec = db.OpenRecordset("select * from Invoicehead where InvNo=" & Me.txtInvNo.Text & " and InvType='" & Me.CboInvType.Text & "'")
        If Not rec.EOF Then
            Me.cbocash.Text = rec("BillType")

            Set rec7 = db.OpenRecordset("select * from LedgerMAster where AccID=" & rec("AccId"))
            If Not rec7.EOF Then
                i = 1
                'Me.cboParty.ListIndex = i
                Do While i < Me.cboParty.ListCount
                    'Me.cboParty.ListIndex = i
                    If Not IsNull(rec7("AccName")) Then
                        If Me.cboParty.ItemData(i) = rec("Accid") Then
                            Me.cboParty.ListIndex = i
                            Exit Do
                        Else
                            i = i + 1
                        End If
                    End If
                Loop
                'Me.cboParty.Text = Trim(rec5("AccName"))
            End If
            '---------------------------------
            db.Execute ("delete * from Temp_invoice")
            Set rec2 = db.OpenRecordset("select * from invoicedetails where InvNo=" & Me.txtInvNo.Text & " and InvType='" & INV_TYPE & "' order by slno")
            If Not rec2.EOF Then
                While Not rec2.EOF
                    db.Execute ("insert into Temp_Invoice (ItemName,ProductCode,Units,Qty,SaleRate,Gross,TradeDiscount,SpecialDiscount,DiscountAmount,MRP,Vat,VatAmount,Net,Free_Qty,Tax_type,pack,slno) values('" & rec2("ItemName") & "'," & rec2("ProductCode") & ",'" & rec2("Units") & "'," & rec2("Qty") & "," & rec2("SaleRate") & "," & rec2("Gross") & "," & rec2("TradeDiscount") & "," & rec2("SpecialDiscount") & "," & rec2("DiscountAmount") & "," & rec2("MRP") & "," & rec2("Vat") & "," & rec2("VatAmount") & "," & rec2("Net") & "," & rec2("Free_Qty") & ",'" & rec2("Tax_type") & "'," & rec2("pack") & "," & rec2("slno") & ")")
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

            If Not IsNull(rec!orderdate) Then
                Me.txtorderdate.Text = rec("orderdate")
            Else
                Me.txtorderdate.Text = "__/__/____"
            End If

            If Not IsNull(rec!orderno) Then
                Me.txtorderno.Text = rec("orderno")
            Else
                Me.txtorderno.Text = ""
            End If

            Set rec2 = db.OpenRecordset("select max(slno) as max_slno from temp_invoice")
            If Not IsNull(rec2!max_slno) Then
                temp_slno = rec2!max_slno
            Else
                temp_slno = 0
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
            Me.txtfreight.Text = rec("Freight")
            If Not IsNull(rec("TotalCase")) Then
                Me.txttotalcase.Text = rec("TotalCase")
            End If
            If Not IsNull(rec("add_ded")) Then
                Me.txtadd.Text = rec("add_ded")
            End If
            Me.txtGrandtotal.Text = Format(rec("GrandTotal"), "################0.00")
            Me.txtInvDate.SetFocus
        End If

        '----------------------------------------Delete Invoice----------------------------
        If DEL_INVOICE = "Y" Then
            ans = MsgBox("Confirm Delete?", vbYesNo)
            If ans = 6 Then
                '----------Update Stock--------------------------
                Set rec2 = db.OpenRecordset("select * from invoicedetails where InvNo=" & Me.txtInvNo.Text & " and InvType='" & INV_TYPE & "'")
                If Not rec2.EOF Then
                    While Not rec2.EOF

                        stockqty = rec2("Qty") + rec2("Free_qty")

                        db.Execute ("update stock set Qty=Qty+" & stockqty & " where ProductCode=" & rec2("ProductCode"))

                        rec2.MoveNext
                    Wend
                    db.Execute ("delete * from invoicedetails where InvNo=" & Me.txtInvNo.Text & " and InvType='" & INV_TYPE & "'")
                End If
                '-----------------------Delete * From LedgerTran--------------------------

                db.Execute ("delete * from LedgerTran Where VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)


                db.Execute ("delete * from invoicehead where InvNo=" & Me.txtInvNo.Text & " and InvType='" & INV_TYPE & "'")
                DEL_INVOICE = ""
                db.Execute ("delete * from Temp_Invoice")
                temp_slno = 0
                Me.Data1.RecordSource = "select * from Temp_Invoice order by slno desc"
                Data1.Refresh
                Me.CboInvType.Clear
                Me.CboInvType.Enabled = False
                Me.txtInvNo.Locked = True
                Me.Label6.Enabled = True
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

Private Sub cboitemname_Change()
'cboitemname.Text = Trim(Me.DBGrid1.Columns(1))
    txtproductcode.Text = Me.cboitemname.ItemData(Me.cboitemname.ListIndex)
    Set rec1 = db.OpenRecordset("SELECT * FROM ITEMMASTER WHERE PRODUCTCODE=" & Me.txtproductcode.Text)
    If Not rec1.EOF Then
        txtmrp.Text = rec1("MRP")
        txtsalerate.Text = rec1("SALERATE")
        TxtVat.Text = rec1("TAX")
        cbounit.Text = rec1("UNITTYPE")
        txttaxtype.Text = rec1("TAX_TYPE")
        Set rec2 = db.OpenRecordset("SELECT QTY FROM STOCK WHERE PRODUCTCODE=" & Me.txtproductcode.Text)
        If Not rec2.EOF Then
            Me.TXTSTOCK.Text = "Stock:: " & rec2("qty")
        End If
    End If
End Sub

Private Sub cboitemname_Click()
    cboitemname_Change
End Sub

Private Sub cboitemname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Me.txttotalcase.SetFocus
    End If
    If KeyCode = 13 Then
        '        SEARCHWORD = Trim(Me.cboitemname.Text)
        '        frmproductlist.Show vbModal
        Me.txtpack.SetFocus

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
            Me.cboitemname.ListIndex = 1
        End If
    End If
    If KeyCode = vbKeyF1 Then
        formid = "Invoice"
        frmnewitemmaster.Show vbModal
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
            If Not IsNull(rec1("statecode")) Then
                Set rec2 = db.OpenRecordset("select * from statecode where stcode=" & rec1("statecode"))
            End If
            If Not rec2.EOF Then
                Me.txtstate.Text = rec2("statename")
                Me.txtstatecode.Text = rec2("stcode")
            End If
        Else
            Me.txtTin.Text = ""
            Me.txtstate.Text = ""
            Me.txtstatecode.Text = ""
        End If
        Me.txttradediscount.Text = rec1("Discount")
        Me.txtcrlimit.Text = Format(rec1("CrLimit"), "########0.00")
    Else
        Me.txtTin.Text = ""
        Me.txttradediscount.Text = 0
        Me.txtcrlimit.Text = "0.00"
    End If

    If Me.CboInvType.Enabled = False Then

        '        TEMP_SUBFIX = Left(Me.txtTin.Text, 5)
        '        If Len(Me.txtTin.Text) > 0 Then
        '            INV_TYPE = "TAX"
        '        Else
        '            INV_TYPE = "RETAIL"
        '        End If
        '        If TEMP_SUBFIX = "GSTIN" Then
        '            INV_TYPE = "TAX"
        '        End If
        '        If TEMP_SUBFIX = "" Then
        '            INV_TYPE = "RETAIL"
        '        End If
        INV_TYPE = "TAX"
        Set rec1 = db.OpenRecordset("select max(InvNo) as max_slno from InvoiceHead where InvType='" & INV_TYPE & "'")
        If Not IsNull(rec1!max_slno) Then
            Me.txtInvNo.Text = rec1!max_slno + 1
            If SoftwareVersion = "Demo" And Val(Me.txtInvNo.Text) > 2 Then
                MsgBox "Demo Expired", vbCritical
                'Unload frmMain
                End
                End
            End If
        Else
            Me.txtInvNo.Text = 1
        End If
    End If



    Set rec2 = db.OpenRecordset("select * from LedgerMaster where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not rec2.EOF Then
        TEMP_DEBTOR_GROUPID = rec2("GroupId")
    End If

    Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not IsNull(rs!max_dr) Then
        temp_dr = rs!max_dr
    End If
    Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not IsNull(rs!max_cr) Then
        temp_cr = rs!max_cr
    End If
    Set rs = db.OpenRecordset("select * from ledgermaster where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not rs.EOF Then
        If rs("BalanceType") = "Dr" Then
            temp_dr = temp_dr + rs("OBalance")
        End If
        If rs("BalanceType") = "Cr" Then
            temp_cr = temp_cr + rs("OBalance")
        End If
        Me.txtbalance.Text = Format(temp_dr - temp_cr, "#######0.00")
    Else
        Me.txtbalance.Text = 0
    End If
    If Me.cboParty.Text Like "CASH*" Then
        Me.txtcrlimit.Text = 10000000
    End If


    If Me.CboZone.ListCount > 0 Then
        Me.CboZone.Clear
    End If
    Set rec4 = db.OpenRecordset("select * from PartyDr where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex))
    If Not rec4.EOF Then
        Set rec5 = db.OpenRecordset("select * from ZoneMaster where slno=" & rec4("ZoneCode"))
        If Not rec5.EOF Then
            Me.CboZone.AddItem (rec5("ZoneName"))
            Me.CboZone.ItemData(Me.CboZone.NewIndex) = rec5("Slno")
        End If
    End If

    If Me.cboParty.Text Like "CASH*" Then
        Me.CboZone.AddItem ("CASH")
        Me.CboZone.ItemData(Me.CboZone.NewIndex) = 250
    End If
    If Me.CboZone.ListCount = 0 Then
        Me.CboZone.AddItem ("CASH")
        Me.CboZone.ItemData(Me.CboZone.NewIndex) = 250
    End If
    If Me.CboZone.ListCount > 0 Then
        Me.CboZone.ListIndex = 0
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
        Me.CboZone.SetFocus
    End If
End Sub
'Private Sub cboproducttype_Change()
'Set rec1 = db.OpenRecordset("select distinct ItemType from ItemMaster where ProductType='" & Me.cboproducttype.Text & "'")
'Me.cboitemtype.Clear
'If Not rec1.EOF Then
'    While Not rec1.EOF
'    Me.cboitemtype.AddItem (rec1("ItemType"))
'    rec1.MoveNext
'    Wend
'    If Me.cboitemtype.ListCount > 0 Then
'    Me.cboitemtype.ListIndex = 0
'    End If
'End If
'End Sub
'Private Sub cboproducttype_Click()
'cboproducttype_Change
'End Sub
'Private Sub cboproducttype_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then
'Me.txtlessdiscount.SetFocus
'End If
'If KeyCode = 13 Then
'Me.cboitemtype.SetFocus
'End If
'End Sub

'Private Sub cboprrate_Change()
'Set rec1 = db.OpenRecordset("select Salerate from ItemMaster where ProductType='" & Me.cboproducttype.Text & "' and ItemType='" & Me.cboitemtype.Text & "' and Brand='" & Me.cbobrandname.Text & "' and Item='" & Me.cboitemname.Text & "' and Size='" & Me.cbosize.Text & "' and MRP=" & Me.txtmrp.Text & " and Purchaserate=" & Me.cboprrate.Text)
'If Not rec1.EOF Then
'Me.txtsalerate.Text = Format(rec1("Salerate"), "############0.00")
'End If
'Set rec1 = db.OpenRecordset("select Qty from Stock where ProductType='" & Me.cboproducttype.Text & "' and ItemType='" & Me.cboitemtype.Text & "' and Brand='" & Me.cbobrandname.Text & "' and Itemname='" & Me.cboitemname.Text & "' and Size='" & Me.cbosize.Text & "' and MRP=" & Me.txtmrp.Text & " and Prate=" & Me.cboprrate.Text)
'If Not rec1.EOF Then
'Me.txtQty.Text = rec1("Qty")
'End If
'End Sub
'Private Sub cboprrate_Click()
'cboprrate_Change
'End Sub

Private Sub cboprrate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtsalerate.SetFocus
    End If
End Sub

'Private Sub cbosize_Change()
'Set rec1 = db.OpenRecordset("select distinct MRP from stock where ProductType='" & Me.cboproducttype.Text & "' and itemtype='" & Me.cboitemtype.Text & "' and brand='" & Me.cbobrandname.Text & "' and itemname='" & Me.cboitemname.Text & "' and size='" & Me.cbosize.Text & "'")
'Me.txtmrp.Clear
'If Not rec1.EOF Then
'    While Not rec1.EOF
'    Me.txtmrp.AddItem (rec1("mrp"))
'    rec1.MoveNext
'    Wend
'End If
'If Me.txtmrp.ListCount > 0 Then
'Me.txtmrp.ListIndex = 0
'End If
'End Sub
'Private Sub cbosize_Click()
'cbosize_Change
'End Sub
Private Sub cbosize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbounit.SetFocus
    End If
End Sub

Private Sub cbopaymentmode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cboitemname.SetFocus
        'Me.txtbarcode.SetFocus
    End If
End Sub

Private Sub cbounit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtsalerate.SetFocus
    End If
End Sub

Private Sub CboZone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtLrno.SetFocus
    End If
End Sub

Private Sub CmdDelete_Click()
    db.Execute ("delete * from Temp_Invoice")
    Data1.Refresh
    Me.CboInvType.Enabled = True
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
    DEL_INVOICE = "Y"
End Sub

Private Sub CmdEdit_Click()
    db.Execute ("delete * from Temp_Invoice")
    Data1.Refresh
    Me.CboInvType.Enabled = True
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
    EDIT_INVOICE = "y"
    DEL_INVOICE = "n"
    Me.txtInvNo.SetFocus
End Sub
Private Sub cmdprint_Click()
    FrmInvPrint.Show vbModal
End Sub
Private Sub cmdSave_Click()
    ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then
        temp_day = Left((Me.txtInvDate.Text), 2)
        temp_month = Mid((Me.txtInvDate.Text), 4, 2)
        temp_year = Right((Me.txtInvDate.Text), 4)

        Accperiod_day = Left(AccountingPeriod, 2)
        Accperiod_month = Mid(AccountingPeriod, 4, 2)
        Accperiod_year = Right(AccountingPeriod, 4)

        X = NumberToWord(Trim(Str(Round(Val(Me.txtGrandtotal.Text)))))

        BillType = Me.cbocash.Text
        If BillType = "CASH" Then
            '            i = 0
            '            cash_Party = Me.cboParty.Text
            '            Do While i < Me.cboParty.ListCount
            '               If Me.cboParty.List(i) = "CASH" Then
            '                    Me.cboParty.ListIndex = i
            '                    Exit Do
            '               Else
            '                    i = i + 1
            '               End If
            '            Loop
        End If
        Set rec1 = db.OpenRecordset("select * from Invoicehead where invno=" & Val(Me.txtInvNo.Text) & " and INVTYPE='" & INV_TYPE & "'")
        If rec1.EOF Then
            temp_accid = 0
            db.Execute ("insert into InvoiceHead (InvNo,InvDate,ChalanNo,ChalanDate,InvType,AccId,LrNo,Party,TotalQty,TotalGross,TradeDiscount,SpecialDiscount,VatAmount,Net,RndUp,GrandTotal,AmountInText,Freight,MrpAmount,BillType,TotalCase,add_ded,outstanding,orderno,orderdate,paymentmode)values(" & Me.txtInvNo.Text & ",'" & Me.txtInvDate.Text & "','" & Me.txtChalanNo.Text & "','" & Me.txtChalanDate.Text & "','" & INV_TYPE & "'," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ",'" & Me.txtLrno.Text & "','" & Me.cboParty.Text & "'," & Val(Me.txtTotalqty.Text) & "," & Val(Me.txttotalgross.Text) & "," & Val(Me.txttotaltradediscount.Text) & "," & Val(Me.txttotalspecialdiscount.Text) & "," & Val(Me.TxtVatAmount.Text) & "," & Val(Me.TxtNetAmount.Text) & "," & Val(Me.TxtRoundup.Text) & "," & Val(Me.txtGrandtotal.Text) & ",'" & X & "'," & Val(Me.txtfreight.Text) & "," & Val(Me.txtmrpamount.Text) & ",'" & Me.cbocash.Text & "'," & Me.txttotalcase.Text & ",'" & Me.txtadd.Text & "'," & Me.txtbalance.Text & ",'" & _
                        Me.txtorderno.Text & "','" & Me.txtorderdate.Text & "','" & Me.cbopaymentmode.Text & "')")
        Else
            temp_accid = rec1("AccId")
            db.Execute ("UPDATE INVOICEHEAD SET InvDate='" & Me.txtInvDate.Text & "',ChalanNo='" & Me.txtChalanNo.Text & "',ChalanDate='" & Me.txtChalanDate.Text & "',Party='" & Me.cboParty.Text & "',AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ",TotalQty=" & Me.txtTotalqty.Text & ",TotalGross=" & Me.txttotalgross.Text & ",TradeDiscount=" & Val(Me.txttotaltradediscount.Text) & ",SpecialDiscount=" & Val(Me.txttotalspecialdiscount.Text) & ",VatAmount=" & Val(Me.TxtVatAmount.Text) & ",Net=" & Val(Me.TxtNetAmount.Text) & ",RndUp=" & Val(Me.TxtRoundup.Text) & ",GrandTotal=" & Val(Me.txtGrandtotal.Text) & ",AmountInText='" & X & "',Freight=" & Val(Me.txtfreight.Text) & ",MrpAmount=" & Val(Me.txtmrpamount.Text) & ",BillType='" & Me.cbocash.Text & "',TotalCase=" & Val(Me.txttotalcase.Text) & ",add_ded='" & Me.txtadd.Text & "',orderno='" & Me.txtorderno.Text & "',orderdate='" & Me.txtorderdate.Text & "',paymentmode='" & _
                        Me.cbopaymentmode.Text & "' where InvNo=" & Val(Me.txtInvNo.Text) & " and InvType='" & INV_TYPE & "'")
        End If
        '----------------------Checking &  Update previous entry---------------
        Set rec2 = db.OpenRecordset("select * from invoicedetails where InvNo=" & Val(Me.txtInvNo.Text) & " and InvType='" & INV_TYPE & "'")
        If Not rec2.EOF Then
            While Not rec2.EOF
                stockqty = rec2("Qty") + rec2("Free_qty")
                db.Execute ("update stock set Qty=Qty+" & stockqty & " where ProductCode=" & rec2("ProductCode"))
                rec2.MoveNext
            Wend
            db.Execute ("delete * from invoicedetails where InvNo=" & Me.txtInvNo.Text & " and InvType='" & INV_TYPE & "'")
        End If
        '-----------New Stock Entry------------
        Set rec2 = db.OpenRecordset("select * from Temp_Invoice")
        If Not rec2.EOF Then
            While Not rec2.EOF
                db.Execute ("insert into InvoiceDetails (InvNo,InvType,Itemname,Units,MRP,SaleRate,Qty,Gross,SpecialDiscount,Tradediscount,DiscountAmount,Vat,VatAmount,Net,ProductCode,Free_Qty,Tax_type,pack,slno) values(" & Me.txtInvNo.Text & ",'" & INV_TYPE & "','" & rec2("Itemname") & "','" & rec2("Units") & "'," & rec2("MRP") & "," & rec2("SaleRate") & "," & rec2("Qty") & "," & rec2("Gross") & "," & rec2("SpecialDiscount") & "," & rec2("Tradediscount") & "," & rec2("DiscountAmount") & "," & rec2("Vat") & "," & rec2("VatAmount") & "," & rec2("Net") & "," & rec2("ProductCode") & "," & rec2("Free_Qty") & ",'" & rec2("Tax_type") & "'," & rec2("pack") & "," & rec2("slno") & ")")
                stockqty = rec2("Qty") + rec2("Free_qty")
                db.Execute ("update stock set Qty=Qty-" & stockqty & " where ProductCode=" & rec2("ProductCode"))
                rec2.MoveNext
            Wend
        End If
        If BillType = "CASH" Then
            'db.Execute ("update invoicehead set cashparty='" & cash_Party & "' where InvNo=" & Val(Me.txtInvno.Text) & " and InvType='" & INV_TYPE & "'")
        End If
        '---------Deleting Previous Entry of Account Ledger------------------
        db.Execute ("delete * from LedgerTran Where VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)

        '---------NEW TRANSACTION ENTRY------------------------------------
        Set rec4 = db.OpenRecordset("select * from LedgerMaster where AccName like 'Sales*'")
        If Not rec4.EOF Then
            SalesAccId = rec4("Accid")
            Set rec3 = db.OpenRecordset("select * from LedgerTran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & " and Slno=(select max(SlNo) from LedgerTran where AccId=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
            If Not rec3.EOF Then
                LEDGERBALANCE = rec3("Balance")
                LEDGERSLNO = rec3("Slno") + 1
            Else
                LEDGERSLNO = 1
                LEDGERBALANCE = 0
            End If
            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & LEDGERSLNO & ",'" & Me.txtInvDate.Text & "','To Sales A/c'," & Val(Me.txtGrandtotal.Text) & ",0," & LEDGERBALANCE + Val(Me.txtGrandtotal.Text) & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','" & INV_TYPE & " Invoice'," & Me.txtInvNo.Text & "," & TEMP_DEBTOR_GROUPID & "," & SalesAccId & ")")
        End If
        '---------Deleting Previous Entry of Sales Ledger Ledger------------------
        Set rec3 = db.OpenRecordset("select * from LedgerMaster where AccName like 'Sales*'")
        If Not rec3.EOF Then
            Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec3("AccId") & " and  VoucherType='" & INV_TYPE & " Invoice' and VoucherSlNo=" & Me.txtInvNo.Text)
            If Not rec4.EOF Then
                db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec4("Cr") & " where AccId=" & rec3("AccId") & " and SlNo>=" & rec4("SlNo"))
                db.Execute ("delete * from LedgerTran Where AccId=" & rec3("AccId") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
            End If
            '--------New Transaction entry to sales ledger-------------
            Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId= " & rec3("Accid") & " and SlNo=(select max(Slno) from LedgerTran where AccId=" & rec3("AccId") & ")")
            If Not rec4.EOF Then
                Salesledger_Balance = rec4("Balance")
                Salesledger_slno = rec4("Slno") + 1
            Else
                Salesledger_Balance = 0
                Salesledger_slno = 1
            End If
            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Salesledger_slno & ",'" & Me.txtInvDate.Text & "','By " & Me.cboParty.Text & "',0," & Val(Me.txtGrandtotal.Text) - Val(Me.TxtVatAmount.Text) & "," & Salesledger_Balance + Val(Me.txtGrandtotal.Text) - Val(Me.TxtVatAmount.Text) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','" & INV_TYPE & " Invoice'," & Me.txtInvNo.Text & "," & rec3("GroupId") & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
        End If

        '        '---------Deleting Previous vat ledger transaction----------
        '        Set rec3 = db.OpenRecordset("select * from ledgermaster where AccName like 'VAT*'")
        '        If Not rec3.EOF Then
        '            Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec3("Accid") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
        '            If Not rec4.EOF Then
        '                db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec4("Cr") & " where AccId=" & rec3("AccId") & " and SlNo>=" & rec4("SlNo"))
        '                db.Execute ("delete * from LedgerTran Where AccId=" & rec3("AccId") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
        '            End If
        '            '---------New Transaction Vat tax ledger----------
        '            If Val(Me.TxtVatAmount.Text) > 0 Then
        '                Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId= " & rec3("Accid") & " and SlNo=(select max(Slno) from LedgerTran where AccId=" & rec3("AccId") & ")")
        '                If Not rec4.EOF Then
        '                    Vat_Balance = rec4("Balance")
        '                    Vat_slno = rec4("Slno") + 1
        '                Else
        '                    Vat_Balance = 0
        '                    Vat_slno = 1
        '                End If
        '                db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Vat_slno & ",'" & Me.txtInvDate.Text & "','By " & Me.cboParty.Text & "',0," & Me.TxtVatAmount.Text & "," & Vat_Balance + Val(Me.TxtVatAmount.Text) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','" & INV_TYPE & " Invoice'," & Me.txtInvNo.Text & "," & rec3("GroupId") & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
        '            End If
        '        End If
        '---------Deleting Previous SGST ledger transaction----------
        If Val(Me.txtstatecode.Text) = 21 Then
            Set rec3 = db.OpenRecordset("select * from ledgermaster where AccName like 'SGST'")
            If Not rec3.EOF Then
                Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec3("Accid") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
                If Not rec4.EOF Then
                    db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec4("Cr") & " where AccId=" & rec3("AccId") & " and SlNo>=" & rec4("SlNo"))
                    db.Execute ("delete * from LedgerTran Where AccId=" & rec3("AccId") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
                End If
                '---------New Transaction Vat tax ledger----------
                If Val(Me.TxtVatAmount.Text) > 0 Then
                    Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId= " & rec3("Accid") & " and SlNo=(select max(Slno) from LedgerTran where AccId=" & rec3("AccId") & ")")
                    If Not rec4.EOF Then
                        Vat_Balance = rec4("Balance")
                        Vat_slno = rec4("Slno") + 1
                    Else
                        Vat_Balance = 0
                        Vat_slno = 1
                    End If
                    db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Vat_slno & ",'" & Me.txtInvDate.Text & "','By " & Me.cboParty.Text & "',0," & Val(Me.TxtVatAmount.Text) / 2 & "," & Vat_Balance + (Val(Me.TxtVatAmount.Text) / 2) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','" & INV_TYPE & " Invoice'," & Me.txtInvNo.Text & "," & rec3("GroupId") & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
                End If
            End If
        End If
        '---------Deleting Previous CGST ledger transaction----------
        If Val(Me.txtstatecode.Text) = 21 Then
            Set rec3 = db.OpenRecordset("select * from ledgermaster where AccName like 'CGST'")
            If Not rec3.EOF Then
                Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec3("Accid") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
                If Not rec4.EOF Then
                    db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec4("Cr") & " where AccId=" & rec3("AccId") & " and SlNo>=" & rec4("SlNo"))
                    db.Execute ("delete * from LedgerTran Where AccId=" & rec3("AccId") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
                End If
                '---------New Transaction Vat tax ledger----------
                If Val(Me.TxtVatAmount.Text) > 0 Then
                    Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId= " & rec3("Accid") & " and SlNo=(select max(Slno) from LedgerTran where AccId=" & rec3("AccId") & ")")
                    If Not rec4.EOF Then
                        Vat_Balance = rec4("Balance")
                        Vat_slno = rec4("Slno") + 1
                    Else
                        Vat_Balance = 0
                        Vat_slno = 1
                    End If
                    db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Vat_slno & ",'" & Me.txtInvDate.Text & "','By " & Me.cboParty.Text & "',0," & Val(Me.TxtVatAmount.Text) / 2 & "," & Vat_Balance + (Val(Me.TxtVatAmount.Text) / 2) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','" & INV_TYPE & " Invoice'," & Me.txtInvNo.Text & "," & rec3("GroupId") & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
                End If
            End If
        End If
        '---------Deleting Previous IGST ledger transaction----------
        If Val(Me.txtstatecode.Text) <> 21 Then
            Set rec3 = db.OpenRecordset("select * from ledgermaster where AccName like 'IGST'")
            If Not rec3.EOF Then
                Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec3("Accid") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
                If Not rec4.EOF Then
                    db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec4("Cr") & " where AccId=" & rec3("AccId") & " and SlNo>=" & rec4("SlNo"))
                    db.Execute ("delete * from LedgerTran Where AccId=" & rec3("AccId") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
                End If
                '---------New Transaction Vat tax ledger----------
                If Val(Me.TxtVatAmount.Text) > 0 Then
                    Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId= " & rec3("Accid") & " and SlNo=(select max(Slno) from LedgerTran where AccId=" & rec3("AccId") & ")")
                    If Not rec4.EOF Then
                        Vat_Balance = rec4("Balance")
                        Vat_slno = rec4("Slno") + 1
                    Else
                        Vat_Balance = 0
                        Vat_slno = 1
                    End If
                    db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & Vat_slno & ",'" & Me.txtInvDate.Text & "','By " & Me.cboParty.Text & "',0," & Val(Me.TxtVatAmount.Text) / 2 & "," & Vat_Balance + (Val(Me.TxtVatAmount.Text) / 2) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','" & INV_TYPE & " Invoice'," & Me.txtInvNo.Text & "," & rec3("GroupId") & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
                End If
            End If
        End If
        '//---------SHORTAGE/EXCESS
        ''  Set rec3 = db.OpenRecordset("select * from ledgermaster where AccName like 'SHORT/EXCESS*'")
        ''  If Not rec3.EOF Then
        ''        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId=" & rec3("accid") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
        ''        If Not rec4.EOF Then
        ''        db.Execute ("update LedgerTran set SLno=SlNo-1,Balance=Balance-" & rec4("Cr") & " where AccId=" & rec3("AccId") & " and SlNo>=" & rec4("SlNo"))
        ''        db.Execute ("delete * from LedgerTran Where AccId=" & rec3("AccId") & " and VoucherType='" & INV_TYPE & " Invoice' and VoucherSlno=" & Me.txtInvNo.Text)
        ''        End If
        ''    '---------New Transaction Vat tax ledger----------
        ''    If Me.TxtRoundup.Text <> 0 Then
        ''        Set rec4 = db.OpenRecordset("select * from LedgerTran where AccId= " & rec3("Accid") & " and SlNo=(select max(Slno) from LedgerTran where AccId=" & rec3("AccId") & ")")
        ''        If Not rec4.EOF Then
        ''        short_Balance = rec4("Balance")
        ''        short_slno = rec4("Slno") + 1
        ''        Else
        ''        short_Balance = 0
        ''        short_slno = 1
        ''        End If
        ''        db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & short_slno & ",'" & Me.txtInvDate.Text & "','By " & Me.cboParty.Text & "',0," & Me.TxtRoundup.Text & "," & short_Balance + Val(Me.TxtRoundup.Text) & "," & rec3("AccId") & ",'Inv No:" & Me.txtInvNo.Text & "-" & Me.txtInvDate.Text & "','" & INV_TYPE & " Invoice'," & Me.txtInvNo.Text & "," & rec3("GroupId") & "," & Me.cboParty.ItemData(Me.cboParty.ListIndex) & ")")
        ''    End If
        ''  End If

        db.Execute ("delete * from temp_invoice")
        temp_slno = 0
        Data1.Refresh
        SalesAccId = 0
        InvNo = 0
        BillType = ""
        Me.txttotalgross.Text = "0.00"
        Me.txttotalcase.Text = "0"
        Me.txtChalanNo.Text = ""
        Me.txtTotalqty.Text = ""
        Me.TxtRoundup.Text = "0.00"
        Me.txtLrno.Text = ""
        Me.txttotaltradediscount.Text = "0.00"
        Me.txttotalspecialdiscount.Text = "0.00"
        Me.txtmrpamount.Text = "0.00"
        Me.TxtVat.Text = 0
        Me.TxtVatAmount.Text = "0.00"
        Me.TxtNetAmount.Text = "0.00"
        Me.txtGrandtotal.Text = "0.00"
        Me.txtadd.Text = ""
        Me.txtpack.Text = "0.00"
        Salesledger_Balance = 0
        Salesledger_slno = 0
        LEDGERBALANCE = 0
        LEDGERSLNO = 0
        temp_accid = 0
        Me.CboInvType.Enabled = False
        Me.txtfreight.Text = "0.00"
        Me.cboParty.ListIndex = 0
        EDIT_INVOICE = "n"
        DEL_INVOICE = "n"
        Me.txtInvDate.SetFocus
    End If

End Sub


Private Sub DBGrid1_AfterDelete()
    db.Execute ("update Temp_Invoice set slno=slno-1 where slno>" & temp_del_slno)
    temp_del_slno = 0
    temp_slno = temp_slno - 1
    Me.Data1.RecordSource = "select * from Temp_Invoice order by slno desc"
    Me.Data1.Refresh

End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    TradeDiscount = Round(Val(Me.DBGrid1.Columns(9)) * (Val(Me.DBGrid1.Columns(15)) / 100), 2)
    SpecialDiscount = Round((Val(Me.DBGrid1.Columns(9)) - TradeDiscount) * (Val(Me.DBGrid1.Columns(16)) / 100), 2)

    Me.txttotalspecialdiscount.Text = Format(Val(Me.txttotalspecialdiscount.Text) - Val(SpecialDiscount), "#######0.00")
    Me.txttotaltradediscount.Text = Val(Me.txttotaltradediscount.Text) - Val(TradeDiscount)
    Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) - Val(Me.DBGrid1.Columns(13)), "######0.00")
    Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.DBGrid1.Columns(9)), "######0.00")
    Me.txtmrpamount.Text = Val(Me.txtmrpamount.Text) - ((Val(Me.DBGrid1.Columns(4))) * Val(Me.DBGrid1.Columns(7)))
    Me.txtTotalqty.Text = Format(Val(Me.txtTotalqty.Text) - (Val(Me.DBGrid1.Columns(4)) + Val(Me.DBGrid1.Columns(6))), "######0.00")
    Me.txttotalcase.Text = Val(Me.txttotalcase.Text) - Val(Me.DBGrid1.Columns(3))
    Me.TxtNetAmount.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.txttotaltradediscount.Text) - Val(Me.txttotalspecialdiscount.Text) + Val(Me.TxtVatAmount.Text), "#########0.00")
    Me.TxtRoundup.Text = Format(Round(Me.TxtNetAmount.Text) - Val(Me.TxtNetAmount.Text), "##0.00")
    Me.txtGrandtotal.Text = Round(Me.TxtNetAmount.Text)
    temp_del_slno = Me.DBGrid1.Columns(0)

End Sub

Private Sub DBGrid1_DblClick()
    dbl_click_edit = True
    product_code = Me.DBGrid1.Columns(1)
    EDIT_ITEM_SLNO = Me.DBGrid1.Columns(0)
    Me.cboitemname.ListIndex = 1
    i = 1
    Do While i < Me.cboitemname.ListCount
        If Me.cboitemname.ItemData(i) = product_code Then
            Me.cboitemname.ListIndex = i
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    Me.txtpack.Text = Val(Me.DBGrid1.Columns(3))
    Me.txtQty.Text = Me.DBGrid1.Columns(4)
    Me.cbounit.Text = Me.DBGrid1.Columns(5)
    Me.txtfree.Text = Me.DBGrid1.Columns(6)
    Me.txtmrp.Text = Me.DBGrid1.Columns(7)
    Me.txtsalerate.Text = Me.DBGrid1.Columns(8)
    Me.txtgross.Text = Me.DBGrid1.Columns(9)
    Me.TxtVat.Text = Me.DBGrid1.Columns(12)
    Me.txttaxamount.Text = Me.DBGrid1.Columns(13)
    Me.txtamount.Text = Me.DBGrid1.Columns(14)
    Me.txttradediscount.Text = Me.DBGrid1.Columns(15)
    Me.txtspecialdiscount.Text = Me.DBGrid1.Columns(16)
    Me.txtpack.SetFocus
End Sub

Private Sub Form_Load()
    FORMNAME = "InvoiceR"
    dbl_click_edit = False
    temp_slno = 0
    temp_del_slno = 0
    formid = 103
    Me.Top = 0
    Me.Left = 0
    Data1.databasename = dbname
    Me.txtInvDate.Text = Format(Date, "dd/mm/yyyy")
    Me.txtChalanDate.Text = Format(Date, "dd/mm/yyyy")
    Me.txtorderdate.Text = Format(Date, "dd/mm/yyyy")

    Me.cbocash.AddItem ("CREDIT")
    Me.cbocash.AddItem ("CASH")
    Me.cbocash.AddItem ("REVISED")
    Me.cbocash.ListIndex = 0

    Me.cbopaymentmode.AddItem "CASH"
    Me.cbopaymentmode.AddItem "CHEQUE"
    Me.cbopaymentmode.AddItem "CARD"
    Me.cbopaymentmode.ListIndex = 0

    Set rec = db.OpenRecordset("select * from LedgerMaster where Groupname Like 'Sundry Debtor' or Groupname Like 'Cash-In-Hand'")
    While Not rec.EOF
        Me.cboParty.AddItem (rec("Accname"))
        Me.cboParty.ItemData(Me.cboParty.NewIndex) = rec("Accid")
        rec.MoveNext
    Wend
    If Me.cboParty.ListCount > 0 Then
        Me.cboParty.ListIndex = 0
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FORMNAME = ""
    formid = 0
    db.Execute ("delete * from Temp_Invoice")
End Sub


Private Sub txtadd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtfreight.SetFocus
    End If
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dbl_click_edit = True Then
            t_totalcase = Val(Me.DBGrid1.Columns(3))
            t_txtqty = Val(Me.DBGrid1.Columns(4))
            t_txtfree = Val(Me.DBGrid1.Columns(6))
            t_txtmrp = Val(Me.DBGrid1.Columns(7))
            t_txtsalerate = Val(Me.DBGrid1.Columns(8))
            t_txtgross = Val(Me.DBGrid1.Columns(9))
            t_TxtVat = Val(Me.DBGrid1.Columns(12))
            t_txttaxamount = Val(Me.DBGrid1.Columns(13))
            t_txtamount = Val(Me.DBGrid1.Columns(14))
            t_txttradediscount = Val(Me.DBGrid1.Columns(15))
            t_txtspecialdiscount = Val(Me.DBGrid1.Columns(16))
            de_mrp = Val(t_txtqty) * Val(t_txtmrp)
            new_mrp = Val(Me.txtQty.Text) * Val(Me.txtmrp.Text)
            Me.txtmrpamount.Text = Format((Val(Me.txtmrpamount.Text)) - de_mrp + new_mrp, "######0.00")
            Me.txttotalcase.Text = Val(Me.txttotalcase.Text) - t_totalcase
            TradeDiscount = Round(Val(Me.txtgross.Text) * (Val(Me.txttradediscount.Text) / 100), 2)
            t_TradeDiscount = Round(Val(t_txtgross) * (Val(t_txttradediscount) / 100), 2)
            SpecialDiscount = Round((Val(Me.txtgross.Text) - TradeDiscount) * (Val(Me.txtspecialdiscount.Text) / 100), 2)
            t_SpecialDiscount = Round((Val(t_txtgross) - t_TradeDiscount) * (Val(t_txtspecialdiscount) / 100), 2)

            Me.txtTotalqty.Text = Val(Me.txtTotalqty.Text) - (Val(t_txtqty) + Val(t_txtfree)) + (Val(Me.txtQty.Text) + Val(Me.txtfree.Text))
            Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) - (t_txtgross) + Val(Me.txtgross.Text), "######0.00")
            Me.txttotaltradediscount.Text = Format(Val(Me.txttotaltradediscount.Text) - (t_TradeDiscount) + TradeDiscount, "######0.00")
            Me.txttotalspecialdiscount.Text = Format(Val(Me.txttotalspecialdiscount.Text) - (t_SpecialDiscount) + SpecialDiscount, "########0.00")
            Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) - (t_txttaxamount) + Val(Me.txttaxamount.Text), "#####0.00")
            Me.TxtNetAmount.Text = Format(Val(Me.txttotalgross.Text) - Val(Me.txttotaltradediscount.Text) - Val(Me.txttotalspecialdiscount.Text) + Val(Me.TxtVatAmount.Text), "#########0.00")
            Me.TxtRoundup.Text = Format(Round(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text) - Val(Me.TxtNetAmount.Text), "##0.00")
            Me.txtGrandtotal.Text = Round(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text))
            db.Execute ("update Temp_Invoice set Units='" & Me.cbounit.Text & "',Qty=" & Me.txtQty.Text & ",SaleRate=" & Val(Me.txtsalerate.Text) & ",Gross=" & Me.txtgross.Text & ",TradeDiscount=" & Me.txttradediscount.Text & ",SpecialDiscount=" & Me.txtspecialdiscount.Text & ",DiscountAmount=" & TradeDiscount + SpecialDiscount & ",MRP=" & Me.txtmrp.Text & ",Vat=" & Me.TxtVat.Text & ",VatAmount=" & Me.txttaxamount.Text & ",Net=" & Me.txtamount.Text & ",Free_Qty=" & Me.txtfree.Text & ",Tax_type='" & Me.txttaxtype.Text & "',pack=" & Me.txtpack.Text & " where ProductCode=" & Me.txtproductcode.Text & " and SLNO=" & EDIT_ITEM_SLNO)
            EDIT_ITEM_SLNO = 0
        Else
            temp_slno = temp_slno + 1
            Me.txtmrpamount.Text = Format(Val(Me.txtmrpamount.Text) + (Val(Me.txtQty.Text)) * Val(Me.txtmrp.Text), "######0.00")
            TradeDiscount = Round(Val(Me.txtgross.Text) * (Val(Me.txttradediscount.Text) / 100), 2)
            SpecialDiscount = Round((Val(Me.txtgross.Text) - TradeDiscount) * (Val(Me.txtspecialdiscount.Text) / 100), 2)

            Me.txtTotalqty.Text = Val(Me.txtTotalqty.Text) + Val(Me.txtQty.Text) + Val(Me.txtfree.Text)
            Me.txttotalgross.Text = Format(Val(Me.txttotalgross.Text) + Val(Me.txtgross.Text), "######0.00")
            Me.txttotaltradediscount.Text = Format(Val(Me.txttotaltradediscount.Text) + TradeDiscount, "######0.00")
            Me.txttotalspecialdiscount.Text = Format(Val(Me.txttotalspecialdiscount.Text) + SpecialDiscount, "########0.00")
            Me.TxtVatAmount.Text = Format(Val(Me.TxtVatAmount.Text) + Val(Me.txttaxamount.Text), "#####0.00")
            Me.TxtNetAmount.Text = Format(Val(Me.txttotalgross.Text) - (Val(Me.txttotaltradediscount.Text) + Val(Me.txttotalspecialdiscount.Text)) + Val(Me.TxtVatAmount.Text), "#########0.00")
            Me.TxtRoundup.Text = Format(Round(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text) - Val(Me.TxtNetAmount.Text), "##0.00")
            Me.txtGrandtotal.Text = Round(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text))
            db.Execute ("insert into Temp_Invoice (ItemName,ProductCode,Units,Qty,SaleRate,Gross,TradeDiscount,SpecialDiscount,DiscountAmount,MRP,Vat,VatAmount,Net,Free_Qty,Tax_type,pack,slno) values('" & Me.cboitemname.Text & "'," & Me.txtproductcode.Text & ",'" & Me.cbounit.Text & "'," & Val(Me.txtQty.Text) & "," & Val(Me.txtsalerate.Text) & "," & Me.txtgross.Text & "," & Me.txttradediscount.Text & "," & Me.txtspecialdiscount.Text & "," & TradeDiscount + SpecialDiscount & "," & Val(Me.txtmrp.Text) & "," & Me.TxtVat.Text & "," & Me.txttaxamount.Text & "," & Me.txtamount.Text & "," & Me.txtfree.Text & ",'" & Me.txttaxtype.Text & "'," & Me.txtpack.Text & "," & temp_slno & ")")
        End If
        Me.Data1.RecordSource = "select * from Temp_Invoice order by slno desc"
        Me.Data1.Refresh
        Me.txttotalcase.Text = Val(Me.txttotalcase.Text) + Me.txtpack.Text
        Me.txtQty.Text = 0
        Me.txtfree.Text = 0
        Me.txtmrp.Text = "0.00"
        de_mrp = 0
        new_mrp = 0
        Me.txttradediscount.Text = 0
        Me.txtspecialdiscount.Text = 0
        Me.txttaxamount.Text = "0.00"
        Me.txtpack.Text = 0
        TradeDiscount = 0
        SpecialDiscount = 0
        dbl_click_edit = False
        Me.cboitemname.SetFocus
    End If

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

Private Sub txtfree_Change()
    If Not ValidateNumeric(Me.txtfree.Text) Then
        Me.txtfree.Text = 0
        txtfree_GotFocus
    End If

End Sub

Private Sub txtfree_GotFocus()
    Me.txtfree.SelStart = 0
    Me.txtfree.SelLength = Len(Me.txtfree.Text)
End Sub

Private Sub txtfree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Me.txtsalerate.SetFocus
        Me.txtmrp.SetFocus
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
        Set rec1 = db.OpenRecordset("select * from Invoicehead where InvNo=" & Me.txtInvNo.Text)
        Me.CboInvType.Clear
        If Not rec1.EOF Then
            While Not rec1.EOF
                Me.CboInvType.AddItem (rec1("InvType"))
                rec1.MoveNext
            Wend
            Me.CboInvType.ListIndex = 0
            Me.CboInvType.SetFocus
        End If
    End If
End Sub
Private Sub txtLrno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtorderno.SetFocus
    End If
End Sub

Private Sub txtmrp_GotFocus()
    Me.txtmrp.SelStart = 0
    Me.txtmrp.SelLength = Len(Me.txtmrp.Text)
End Sub

Private Sub txtmrp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("Select mrp from ItemMaster where ProductCode=" & Me.txtproductcode.Text)
        If Not IsNull(rec1!mRP) Then
            If rec1!mRP <> Val(Me.txtmrp.Text) Then
                ans = MsgBox("Update The MRP!", vbYesNo)
                If ans = 6 Then
                    db.Execute ("Update ItemMaster set mRP=" & Val(Me.txtmrp.Text) & " where ProductCode=" & Val(Me.txtproductcode.Text))
                Else
                    'Me.txttradediscount.SetFocus
                    Me.txtsalerate.SetFocus
                End If
            End If
        End If
        'Me.txttradediscount.SetFocus
        Me.txtsalerate.Text = Val(Me.txtmrp.Text) / (1 + Val(Me.TxtVat.Text) / 100)
        Me.txtsalerate.SetFocus
    End If
End Sub

Private Sub txtNetamount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtRoundup.SetFocus
    End If
End Sub

Private Sub txtorderdate_GotFocus()
    Me.txtorderdate.SelStart = 0
    Me.txtorderdate.SelLength = Len(Me.txtorderdate.Text)
End Sub

Private Sub txtorderdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbopaymentmode.SetFocus
    End If
End Sub

Private Sub txtorderno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtorderdate.SetFocus
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
        Me.txtGrandtotal.Text = Format(Val(Me.TxtNetAmount.Text) + Val(Me.txtfreight.Text) + Val(Me.TxtRoundup.Text), "###########0.00")
        txtGrandtotal.SetFocus
    End If
End Sub

Private Sub txtsalerate_Change()
    If Not ValidateNumeric(Me.txtsalerate.Text) Then
        Me.txtsalerate.Text = 0
        txtsalerate_GotFocus
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
        'Me.txtmrp.SetFocus
        'Me.txtgross.SetFocus
        'Me.txttradediscount.SetFocus
        Me.TxtVat.SetFocus
    End If
End Sub

Private Sub txtspecialdiscount_Change()
    If Not ValidateNumeric(Me.txtspecialdiscount.Text) Then
        Me.txtspecialdiscount.Text = 0
        txtspecialdiscount_GotFocus
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

Private Sub txttotalcase_GotFocus()
    Me.txttotalcase.SelStart = 0
    Me.txttotalcase.SelLength = Len(Me.txttotalcase.Text)
End Sub

Private Sub txttotalcase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Set rec1 = db.OpenRecordset("select sum(pack) as total_pack from temp_invoice")
        If Not IsNull(rec1!total_pack) Then
            Me.txttotalcase.Text = rec1!total_pack
        End If
        Set rec1 = db.OpenRecordset("select sum(gross) as total_gross from temp_invoice")
        If Not IsNull(rec1!total_gross) Then
            Me.txttotalgross.Text = rec1!total_gross
        End If
        Set rec1 = db.OpenRecordset("select sum(net) as total_net from temp_invoice")
        If Not IsNull(rec1!total_net) Then
            Me.TxtNetAmount.Text = rec1!total_net
        End If
        Set rec1 = db.OpenRecordset("select sum(vatamount) as total_vat from temp_invoice")
        If Not IsNull(rec1!total_vat) Then
            Me.TxtVatAmount.Text = rec1!total_vat
        End If
        Me.txtadd.SetFocus
    End If
End Sub

Private Sub txttradediscount_Change()
    If Not ValidateNumeric(Me.txttradediscount.Text) Then
        Me.txttradediscount.Text = 0
        txttradediscount_GotFocus
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

Private Sub TxtVat_Change()
    If Not ValidateNumeric(Me.TxtVat.Text) Then
        Me.TxtVat.Text = 0
        txtvat_GotFocus
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
            Me.txttaxamount.Text = Round((Val(Me.txtgross.Text) - TradeDiscount - SpecialDiscount) * (Val(Me.TxtVat.Text) / 100), 2)
        ElseIf Me.txttaxtype.Text = "INCLUSIVE MRP" Then
            Me.txttaxamount.Text = Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) - Format((Val(Me.txtmrp.Text) * (Val(Me.txtQty.Text)) / ((Me.TxtVat.Text / 100) + 1)), "########0.00")
        Else
            Me.txttaxamount.Text = "0.00"
        End If
        Me.txtamount.Text = Round(Val(Me.txtgross.Text) - TradeDiscount - SpecialDiscount + Val(Me.txttaxamount.Text), 2)
        Me.txtamount.SetFocus
    End If
End Sub
